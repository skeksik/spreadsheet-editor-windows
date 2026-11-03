using System.Globalization;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using X = DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using Xm = DocumentFormat.OpenXml.Office.Excel;

namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>
/// Moves a worksheet's rows by a <see cref="RowPlan"/> (docs/spec.md §9.20), after its formulas were rewritten
/// (<see cref="FormulaRewriter"/>): the row elements with their cells, heights and styles; what belongs to a row (merges, hyperlinks,
/// comments, drawing anchors), which moves with its card; and the ranges that do not (conditional formatting, data validation, the
/// dimension, filters, tables, page breaks, ignored errors, protected ranges), which move by the plan's row shifts. The pictures the app
/// shows in removed cards are removed with their rows; every other drawing object in deleted rows moves to the first row after them.
/// </summary>
internal static class RowMover
{
    private const string VmlExcel = "urn:schemas-microsoft-com:office:excel";

    /// <summary>The rows of the last column that hold the picture layout (docs/spec.md §5.2 step 2).</summary>
    private const int MetadataRows = 3;

    /// <summary>Refuses a change of rows on a worksheet that holds parts the writer cannot move rows in.</summary>
    public static void Check(WorksheetPart part, string sheetName)
    {
        if (part.PivotTableParts.Any())
        {
            throw Unsupported(sheetName, "a pivot table");
        }

        if (part.GetPartsOfType<SlicersPart>().Any() || part.GetPartsOfType<TimeLinePart>().Any())
        {
            throw Unsupported(sheetName, "slicers or timelines");
        }

        if (part.GetPartsOfType<WorksheetThreadedCommentsPart>().Any())
        {
            throw Unsupported(sheetName, "threaded comments");
        }

        if (part.Worksheet?.GetFirstChild<X.Controls>() is not null || part.ControlPropertiesParts.Any())
        {
            throw Unsupported(sheetName, "form controls");
        }

        if (part.TableDefinitionParts.Any(t => t.QueryTableParts.Any()))
        {
            throw Unsupported(sheetName, "a query table");
        }
    }

    public static void Move(WorksheetPart part, RowPlan plan, SheetSource origin, string sheetName)
    {
        var worksheet = part.Worksheet!;
        MoveRows(worksheet, plan);
        MoveMerges(worksheet, plan);
        MoveHyperlinks(worksheet, plan);
        ShiftSqref(worksheet.Elements<X.ConditionalFormatting>(), plan, grow: true);
        if (worksheet.GetFirstChild<X.DataValidations>() is { } validations)
        {
            ShiftSqref(validations.Elements<X.DataValidation>(), plan, grow: true);
            var count = validations.Elements<X.DataValidation>().Count();
            if (count == 0)
            {
                validations.Remove();
            }
            else
            {
                validations.Count = (uint)count;
            }
        }

        ShiftSqref(worksheet.Descendants<X.IgnoredError>(), plan);
        ShiftSqref(worksheet.Descendants<X.ProtectedRange>(), plan);
        if (worksheet.GetFirstChild<X.WorksheetExtensionList>() is { } extensions)
        {
            foreach (var sqref in extensions.Descendants<Xm.ReferenceSequence>().ToList())
            {
                // Treated as conditional formatting's and data validation's ranges (x14), which is what the writer expects to find here;
                // not verified that no other extension uses xm:sqref.
                if (TryGrow(sqref.Text, plan) is { } shifted)
                {
                    sqref.Text = shifted;
                }
                else
                {
                    sqref.Parent?.Remove();
                }
            }
        }

        if (worksheet.GetFirstChild<X.SheetDimension>() is { Reference.Value: { } dimension } element
            && TryShift(dimension, plan) is { } newDimension)
        {
            element.Reference = newDimension;
        }

        ShiftFilter(worksheet.GetFirstChild<X.AutoFilter>(), worksheet.GetFirstChild<X.SortState>(), plan);
        foreach (var table in part.TableDefinitionParts.Select(t => t.Table).OfType<X.Table>())
        {
            table.Reference = TryShift(table.Reference?.Value, plan)
                ?? throw Unsupported(sheetName, $"the table \"{table.Name?.Value}\" in rows the save deletes");
            ShiftFilter(table.AutoFilter, table.SortState, plan);
        }

        ShiftBreaks(worksheet.GetFirstChild<X.RowBreaks>(), plan);
        MoveComments(part, plan, sheetName);
        MoveDrawing(part, plan, origin);
    }

    private static void MoveRows(X.Worksheet worksheet, RowPlan plan)
    {
        if (worksheet.GetFirstChild<X.SheetData>() is not { } data)
        {
            return;
        }

        var rows = new List<(int Row, X.Row Element)>();
        foreach (var row in data.Elements<X.Row>().ToList())
        {
            var old = (int)row.RowIndex!.Value;
            row.Remove();
            if (plan.Map(old) is not { } moved)
            {
                continue;
            }

            if (moved != old && old <= MetadataRows)
            {
                // The metadata cells belong to rows 1 to 3 of the last column, not to the row's card (docs/spec.md §5.2 step 2): the
                // writer writes them again where they belong.
                foreach (var cell in row.Elements<X.Cell>().Where(c => CellReference.TryParse(c.CellReference?.Value, out var column, out _) && column == CellReference.MaxColumn).ToList())
                {
                    cell.Remove();
                }
            }

            if (moved != old)
            {
                row.RowIndex = (uint)moved;
                foreach (var cell in row.Elements<X.Cell>())
                {
                    if (CellReference.TryParse(cell.CellReference?.Value, out var column, out _))
                    {
                        cell.CellReference = CellIndex.Name(moved, column);
                    }
                }
            }

            rows.Add((moved, row));
        }

        foreach (var (_, row) in rows.OrderBy(r => r.Row))
        {
            data.AppendChild(row);
        }
    }

    private static void MoveMerges(X.Worksheet worksheet, RowPlan plan)
    {
        if (worksheet.GetFirstChild<X.MergeCells>() is not { } merges)
        {
            return;
        }

        foreach (var merge in merges.Elements<X.MergeCell>().ToList())
        {
            if (merge.Reference?.Value is not { } reference)
            {
                continue;
            }

            if (TryMove(reference, plan) is { } moved)
            {
                merge.Reference = moved;
            }
            else
            {
                merge.Remove();
            }
        }

        var count = merges.Elements<X.MergeCell>().Count();
        if (count == 0)
        {
            merges.Remove();
        }
        else
        {
            merges.Count = (uint)count;
        }
    }

    private static void MoveHyperlinks(X.Worksheet worksheet, RowPlan plan)
    {
        if (worksheet.GetFirstChild<X.Hyperlinks>() is not { } hyperlinks)
        {
            return;
        }

        foreach (var hyperlink in hyperlinks.Elements<X.Hyperlink>().ToList())
        {
            if (hyperlink.Reference?.Value is not { } reference)
            {
                continue;
            }

            if (TryMove(reference, plan) is { } moved)
            {
                hyperlink.Reference = moved;
            }
            else
            {
                hyperlink.Remove();
            }
        }

        if (!hyperlinks.HasChildren)
        {
            hyperlinks.Remove();
        }
    }

    /// <summary>Shifts the <c>sqref</c> of each element; an element whose cells are all deleted is removed.</summary>
    private static void ShiftSqref(IEnumerable<OpenXmlElement> elements, RowPlan plan, bool grow = false)
    {
        foreach (var element in elements.ToList())
        {
            var attribute = element.GetAttributes().FirstOrDefault(a => a.LocalName == "sqref" && string.IsNullOrEmpty(a.NamespaceUri));
            if (attribute.Value is not { Length: > 0 } sqref)
            {
                continue;
            }

            var shifted = grow ? TryGrow(sqref, plan) : TryShift(sqref, plan);
            if (shifted is null)
            {
                element.Remove();
            }
            else if (!string.Equals(shifted, sqref, StringComparison.Ordinal))
            {
                element.SetAttribute(new OpenXmlAttribute(attribute.Prefix, attribute.LocalName, attribute.NamespaceUri, shifted));
            }
        }
    }

    private static void ShiftFilter(X.AutoFilter? filter, X.SortState? sort, RowPlan plan)
    {
        if (filter?.Reference?.Value is { } reference && TryShift(reference, plan) is { } shifted)
        {
            filter.Reference = shifted;
        }

        foreach (var state in new[] { sort, filter?.GetFirstChild<X.SortState>() }.OfType<X.SortState>())
        {
            if (state.Reference?.Value is { } sortReference && TryShift(sortReference, plan) is { } shiftedSort)
            {
                state.Reference = shiftedSort;
            }

            foreach (var condition in state.Elements<X.SortCondition>())
            {
                if (condition.Reference?.Value is { } conditionReference && TryShift(conditionReference, plan) is { } shiftedCondition)
                {
                    condition.Reference = shiftedCondition;
                }
            }
        }
    }

    /// <summary>
    /// Page breaks below rows that move. A break's <c>id</c> is taken as the row it follows (ECMA-376's exact meaning not verified); a
    /// break in deleted rows moves to the first row after them.
    /// </summary>
    private static void ShiftBreaks(X.RowBreaks? breaks, RowPlan plan)
    {
        foreach (var brk in breaks?.Elements<X.Break>() ?? [])
        {
            if (brk.Id?.Value is { } id && id > 0)
            {
                brk.Id = (uint)plan.ShiftPointClamped((int)id);
            }
        }
    }

    /// <summary>
    /// Moves the comments with their cells: the comments part's <c>ref</c> and, in the VML drawing, each note's row and anchor. A comment
    /// on a deleted row is removed with its note.
    /// </summary>
    private static void MoveComments(WorksheetPart part, RowPlan plan, string sheetName)
    {
        if (part.WorksheetCommentsPart?.Comments?.CommentList is not { } list)
        {
            return;
        }

        var moves = new Dictionary<(int Row, int Column), int?>();
        foreach (var comment in list.Elements<X.Comment>().ToList())
        {
            if (!CellReference.TryParse(comment.Reference?.Value, out var column, out var row))
            {
                continue;
            }

            var moved = TryMove(CellIndex.Name(row, column), plan);
            int? newRow = moved is not null && CellReference.TryParse(moved, out _, out var r) ? r : null;
            moves[(row, column)] = newRow;
            if (newRow is { } target)
            {
                if (target != row)
                {
                    comment.Reference = CellIndex.Name(target, column);
                }
            }
            else
            {
                comment.Remove();
            }
        }

        if (moves.All(m => m.Value == m.Key.Row))
        {
            return;
        }

        foreach (var vml in part.VmlDrawingParts)
        {
            MoveNotes(vml, moves, sheetName);
        }
    }

    private static void MoveNotes(VmlDrawingPart vml, Dictionary<(int Row, int Column), int?> moves, string sheetName)
    {
        XDocument document;
        try
        {
            using var stream = vml.GetStream(FileMode.Open, FileAccess.Read);
            document = XDocument.Load(stream, LoadOptions.PreserveWhitespace);
        }
        catch (XmlException ex)
        {
            throw SpreadsheetWriteException.NotSupported(
                $"The comments on the sheet \"{sheetName}\" cannot be moved with their rows: their drawing cannot be read ({ex.Message}).");
        }

        XNamespace x = VmlExcel;
        var changed = false;
        foreach (var data in document.Descendants(x + "ClientData").ToList())
        {
            if (!string.Equals((string?)data.Attribute("ObjectType"), "Note", StringComparison.Ordinal))
            {
                throw Unsupported(sheetName, "form controls");
            }

            if (!int.TryParse(data.Element(x + "Row")?.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var row)
                || !int.TryParse(data.Element(x + "Column")?.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var column)
                || !moves.TryGetValue((row + 1, column + 1), out var moved))
            {
                continue;
            }

            if (moved is not { } target)
            {
                data.Parent?.Remove();
                changed = true;
                continue;
            }

            var offset = target - (row + 1);
            if (offset == 0)
            {
                continue;
            }

            data.Element(x + "Row")!.Value = (row + offset).ToString(CultureInfo.InvariantCulture);
            if (data.Element(x + "Anchor") is { } anchor)
            {
                var parts = anchor.Value.Split(',');
                if (parts.Length == 8
                    && int.TryParse(parts[2].Trim(), NumberStyles.None, CultureInfo.InvariantCulture, out var top)
                    && int.TryParse(parts[6].Trim(), NumberStyles.None, CultureInfo.InvariantCulture, out var bottom))
                {
                    parts[2] = " " + (top + offset).ToString(CultureInfo.InvariantCulture);
                    parts[6] = " " + (bottom + offset).ToString(CultureInfo.InvariantCulture);
                    anchor.Value = string.Join(",", parts);
                }
            }

            changed = true;
        }

        if (!changed)
        {
            return;
        }

        using var output = new MemoryStream();
        using (var writer = XmlWriter.Create(output, new XmlWriterSettings { OmitXmlDeclaration = true }))
        {
            document.Save(writer);
        }

        output.Position = 0;
        vml.FeedData(output);
    }

    /// <summary>
    /// Moves the drawing's cell anchors: an object whose top-left cell lies in a kept card moves with the card; any other moves by the
    /// shifts, and one in deleted rows moves to the first row after them. A picture the app shows (a picture of its own anchored in a
    /// table column of a removed card, docs/spec.md §4.1 step 6) is removed with the card. Absolute anchors stay.
    /// </summary>
    private static void MoveDrawing(WorksheetPart part, RowPlan plan, SheetSource origin)
    {
        if (part.DrawingsPart is not { WorksheetDrawing: { } root } drawing)
        {
            return;
        }

        var removedAny = false;
        foreach (var anchor in root.ChildElements.ToList())
        {
            var (from, to) = anchor switch
            {
                Xdr.TwoCellAnchor two => (two.FromMarker, (OpenXmlCompositeElement?)two.ToMarker),
                Xdr.OneCellAnchor one => (one.FromMarker, null),
                _ => (null, null),
            };
            if (from is null || !TryMarker(from.RowId, out var fromRow))
            {
                continue;
            }

            var row = fromRow + 1;
            if (plan.IsRemoved(row) && IsOwnedPicture(anchor, from, origin))
            {
                anchor.Remove();
                removedAny = true;
                continue;
            }

            var toId = to?.GetFirstChild<Xdr.RowId>();
            var hasTo = TryMarker(toId, out var toRow);
            if (plan.CardOffset(row, row) is { } offset)
            {
                SetMarker(from.RowId!, fromRow + offset);
                if (hasTo)
                {
                    SetMarker(toId!, toRow + offset);
                }

                continue;
            }

            SetMarker(from.RowId!, plan.ShiftPointClamped(row) - 1);
            if (hasTo)
            {
                SetMarker(toId!, plan.ShiftPointClamped(toRow + 1) - 1);
            }
        }

        if (!removedAny)
        {
            return;
        }

        var used = root.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().Select(b => b.Embed?.Value).OfType<string>().ToHashSet(StringComparer.Ordinal);
        foreach (var pair in drawing.Parts.Where(p => p.OpenXmlPart is ImagePart && !used.Contains(p.RelationshipId)).ToList())
        {
            drawing.DeletePart(pair.RelationshipId);
        }
    }

    private static bool IsOwnedPicture(OpenXmlElement anchor, Xdr.FromMarker from, SheetSource origin) =>
        anchor.GetFirstChild<Xdr.Picture>()?.BlipFill?.Blip?.Embed?.Value is { Length: > 0 }
        && TryMarker(from.ColumnId, out var column)
        && column + 1 >= origin.FirstColumn && column + 1 <= origin.LastColumn;

    private static bool TryMarker(OpenXmlLeafTextElement? element, out int value) =>
        int.TryParse(element?.Text, NumberStyles.None, CultureInfo.InvariantCulture, out value);

    private static void SetMarker(OpenXmlLeafTextElement element, int value)
    {
        var text = Math.Max(0, value).ToString(CultureInfo.InvariantCulture);
        if (!string.Equals(element.Text, text, StringComparison.Ordinal))
        {
            element.Text = text;
        }
    }

    private static string? TryMove(string range, RowPlan plan)
    {
        try
        {
            return plan.MoveRange(range);
        }
        catch (FormatException)
        {
            return range;
        }
    }

    private static string? TryGrow(string range, RowPlan plan)
    {
        try
        {
            return plan.GrowRange(range);
        }
        catch (FormatException)
        {
            return range;
        }
    }

    private static string? TryShift(string? range, RowPlan plan)
    {
        if (range is null)
        {
            return null;
        }

        try
        {
            return plan.ShiftRange(range);
        }
        catch (FormatException)
        {
            return range;
        }
    }

    private static SpreadsheetWriteException Unsupported(string sheetName, string what) =>
        SpreadsheetWriteException.NotSupported(
            $"Rows cannot be inserted or deleted on the sheet \"{sheetName}\" yet: it holds {what}, which the save cannot move with its rows.");
}
