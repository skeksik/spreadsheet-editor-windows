using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using SpreadsheetEditor.Core.Formulas;
using SpreadsheetEditor.Core.Imaging;
using SpreadsheetEditor.Core.Model;
using SpreadsheetEditor.Core.Values;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>
/// Writes one sheet's table into its worksheet's DOM (docs/spec.md §5.2, §5.5–5.6 with §9.2–9.3, §9.7, §9.16–9.19). The model is compared
/// with the worksheet as it was read (<see cref="SourceWorksheet"/>): a cell is changed only when its value, style or position changed,
/// so everything else stays as read. The worksheet's rows were already moved by the sheet's <see cref="RowPlan"/> (<see cref="RowMover"/>):
/// each card is written into its rows after the move, and a lookup in the worksheet as read goes through the plan to the row as read.
/// </summary>
internal sealed class WorksheetWriter
{
    private readonly Sheet sheet;
    private readonly X.Worksheet worksheet;
    private readonly CellIndex cells;
    private readonly SourceWorksheet source;
    private readonly WriteContext context;
    private readonly RowPlan? plan;
    private readonly SheetPictures pictures;
    private readonly DrawingWriter drawing;

    /// <summary>The table columns (0-based) holding a value whose pictures were laid out again: they are widened for them.</summary>
    private readonly SortedSet<int> laidOutColumns = [];

    /// <summary>The table's columns as read; cells to their right that are not part of the old table are not overwritten.</summary>
    private readonly int sourceLastColumn;

    private WorksheetWriter(
        Sheet sheet, WorksheetPart part, SourceWorksheet source, WriteContext context, RowPlan? plan, SheetPictures pictures)
    {
        this.sheet = sheet;
        this.plan = plan;
        this.pictures = pictures;
        worksheet = part.Worksheet!;
        drawing = new DrawingWriter(part, sheet.Name);
        this.source = source;
        this.context = context;
        cells = new CellIndex(worksheet);
        sourceLastColumn = sheet.Source?.LastColumn ?? int.MaxValue;
    }

    /// <summary>
    /// Writes a sheet read from <c>.xlsx</c> into its table's rows as <paramref name="plan"/> moved them, or a new sheet (no plan) into an
    /// empty worksheet, with the pictures <paramref name="pictures"/> lays out again.
    /// </summary>
    public static void Write(Sheet sheet, WorksheetPart part, SourceWorksheet source, WriteContext context, RowPlan? plan, SheetPictures pictures)
    {
        Validate(sheet);
        var writer = new WorksheetWriter(sheet, part, source, context, plan, pictures);
        if (sheet.Source is { } origin && plan is not null)
        {
            writer.WriteOverSource(origin, plan);
        }
        else
        {
            writer.WriteNew();
        }

        writer.WidenPictureColumns();
        writer.drawing.Finish();
        writer.WriteMetadata();
    }

    /// <summary>Refuses a sheet without entries, or a card without one value per entry.</summary>
    public static void Validate(Sheet sheet)
    {
        if (sheet.Columns.Count == 0)
        {
            throw SpreadsheetWriteException.Invalid($"The sheet \"{sheet.Name}\" has no entries; a sheet needs at least one.");
        }

        foreach (var card in sheet.Cards)
        {
            if (card.Entries.Count != sheet.Columns.Count)
            {
                throw SpreadsheetWriteException.Invalid($"A card on the sheet \"{sheet.Name}\" does not have one value per entry.");
            }
        }
    }

    private void WriteOverSource(SheetSource origin, RowPlan rows)
    {
        if (sheet.FirstColumn != origin.FirstColumn)
        {
            throw SpreadsheetWriteException.NotSupported($"The table on the sheet \"{sheet.Name}\" was moved to another column.");
        }

        WriteName(origin, rows);
        WriteHeader(origin, rows);

        for (var i = 0; i < rows.Cards.Count; i++)
        {
            WriteCard(sheet.Cards[i], rows.Cards[i], origin);
        }

        ClearRemovedColumns(origin, rows);
    }

    private void WriteName(SheetSource origin, RowPlan rows)
    {
        var name = sheet.TableName ?? "";
        var firstTableRow = sheet.HeaderRow ?? sheet.FirstDataRow;
        if (rows.NameRow is not { } row)
        {
            return;
        }

        var column = sheet.TableNameColumn;
        var original = "";
        if (origin.FirstRow < firstTableRow)
        {
            var read = source.Grid.Get(origin.FirstRow, column);
            original = read is null || read.IsBlank ? "" : read.Value.DisplayText;
        }

        if (!string.Equals(name, original, StringComparison.Ordinal))
        {
            // A cleared name clears its cell; the row stays (§9.17). A name given to a table without a name row goes in the row inserted
            // above its header (§9.18).
            WriteText(row, column, name);
        }
    }

    private void WriteHeader(SheetSource origin, RowPlan rows)
    {
        if (sheet.HeaderRow is not { } row || rows.HeaderRow is not { } target)
        {
            return;
        }

        for (var j = 0; j < sheet.Columns.Count; j++)
        {
            var column = sheet.Columns[j];
            var targetColumn = sheet.FirstColumn + j;
            var name = IsUnnamed(column.Name) ? "" : column.Name;
            if (column.SourceIndex is { } index)
            {
                var from = origin.FirstColumn + index;
                var read = source.Grid.Get(row, from);
                var original = read is null || read.IsBlank ? "" : read.Value.DisplayText;
                if (string.Equals(name, original, StringComparison.Ordinal))
                {
                    // An unchanged header keeps its cell as it is (a number stays a number); a moved one is copied (§9.3).
                    if (from != targetColumn)
                    {
                        CopyCell(row, from, target, targetColumn);
                    }

                    continue;
                }
            }

            WriteText(target, targetColumn, name);
        }
    }

    private void WriteCard(Card card, RowPlan.CardRows rows, SheetSource origin)
    {
        var row = rows.FirstRow;
        for (var j = 0; j < sheet.Columns.Count; j++)
        {
            var entry = card.Entries[j];
            var target = sheet.FirstColumn + j;
            // A duplicated card's values come from the card it copies, as read: unedited ones are copied (formulas translated).
            (int Row, int Column)? from = (rows.SourceFirstRow ?? card.CopiedFromRow) is { } first && sheet.Columns[j].SourceIndex is { } index
                ? (first, origin.FirstColumn + index)
                : null;
            // A kept card's cells moved with its rows: a value in its own column is in place.
            WriteEntry(entry, row, target, from, inPlace: rows.IsKept && from?.Column == target);
        }

        if (rows.Rows > rows.OwnRows && rows.Rows > 1)
        {
            MergeCard(rows);
        }

        WritePictures(card, rows, origin);
    }

    /// <summary>
    /// Lays out again the pictures of the card's values that do not keep theirs as read (<see cref="SheetPictures.Keeps"/>): the pictures
    /// the app owns in the value's cell, over the card's rows, are removed, and the value's pictures are anchored as
    /// <see cref="PictureLayout.ForCard"/> places them (§5.4). When any of them has pictures, the card's rows get their heights (§5.3 step 5).
    /// </summary>
    private void WritePictures(Card card, RowPlan.CardRows rows, SheetSource? origin)
    {
        var layout = pictures.ForCard(card);
        var lastRow = rows.FirstRow + rows.Rows - 1;
        var laidOut = false;
        for (var j = 0; j < sheet.Columns.Count; j++)
        {
            if (pictures.Keeps(card, rows.SourceFirstRow, j))
            {
                continue;
            }

            var column = sheet.FirstColumn + j;
            if (origin is not null && column >= origin.FirstColumn && column <= origin.LastColumn)
            {
                // Only pictures in the table's columns as read are the app's: a picture beside the table is not removed.
                drawing.RemoveOwned(rows.FirstRow, lastRow, column, column);
            }

            foreach (var placement in layout.Placements[j])
            {
                drawing.Add(placement, rows.FirstRow + placement.RowOffset, column);
            }

            if (layout.Placements[j].Count > 0)
            {
                laidOut = true;
                laidOutColumns.Add(j);
            }
        }

        if (!laidOut)
        {
            return;
        }

        var heights = layout.RowHeightsPoints;
        for (var i = 0; i < heights.Count && i < rows.Rows; i++)
        {
            if (heights[i] is { } points)
            {
                var row = cells.Row(rows.FirstRow + i);
                row.Height = points;
                row.CustomHeight = true;
            }
        }
    }

    /// <summary>Widens the table columns whose pictures were laid out again to their pictures' width (§5.2 step 8).</summary>
    private void WidenPictureColumns()
    {
        foreach (var j in laidOutColumns)
        {
            if (pictures.Layout.Columns[j].WidthTargetPx is { } target)
            {
                ColumnWidthWriter.Widen(worksheet, sheet.FirstColumn + j, target);
            }
        }
    }

    /// <summary>
    /// Merges each table column over a card's rows when the card gained rows (a new card of several rows, or a card whose span grew,
    /// docs/spec.md §5.3 step 3 with §9.7 and §9.29): a column is not merged over a value in the card's later rows, nor where another merge
    /// already covers part of the card's rows; the card's own vertical merges of a column are replaced by the longer one.
    /// </summary>
    private void MergeCard(RowPlan.CardRows rows)
    {
        var first = rows.FirstRow;
        var last = rows.FirstRow + rows.Rows - 1;
        var merges = worksheet.GetFirstChild<X.MergeCells>();
        var existing = merges?.Elements<X.MergeCell>()
            .Select(m => (Element: m, Ok: CellReference.TryParseRange(m.Reference?.Value, out var range), Range: range))
            .Where(m => m.Ok && m.Range.FirstRow <= last && m.Range.LastRow >= first)
            .ToList() ?? [];
        for (var j = 0; j < sheet.Columns.Count; j++)
        {
            var column = sheet.FirstColumn + j;
            var inColumn = existing.Where(m => m.Range.FirstColumn <= column && m.Range.LastColumn >= column).ToList();
            if (inColumn.Exists(m => m.Range.FirstColumn != m.Range.LastColumn || m.Range.FirstRow < first || m.Range.LastRow > last))
            {
                continue;
            }

            var hasLaterValue = false;
            for (var row = first + 1; row <= last && !hasLaterValue; row++)
            {
                hasLaterValue = cells.Get(row, column) is { } cell && HasContent(cell);
            }

            if (hasLaterValue)
            {
                continue;
            }

            foreach (var (element, _, _) in inColumn)
            {
                element.Remove();
            }

            merges ??= InsertMergeCells();
            merges.AppendChild(new X.MergeCell { Reference = CellIndex.Name(first, column) + ":" + CellIndex.Name(last, column) });
        }

        if (merges is not null)
        {
            merges.Count = (uint)merges.Elements<X.MergeCell>().Count();
        }
    }

    private X.MergeCells InsertMergeCells()
    {
        var merges = new X.MergeCells();
        // mergeCells follows sheetData and the elements that may come between them (ECMA-376 Part 1 §18.3.1.99, not verified in full).
        OpenXmlElement after = worksheet.GetFirstChild<X.SheetData>()!;
        foreach (var name in new[] { "sheetCalcPr", "sheetProtection", "protectedRanges", "scenarios", "autoFilter", "sortState", "dataConsolidate", "customSheetViews" })
        {
            if (worksheet.ChildElements.FirstOrDefault(e => e.LocalName == name) is { } element)
            {
                after = element;
            }
        }

        worksheet.InsertAfter(merges, after);
        return merges;
    }

    private void WriteEntry(Entry entry, int row, int column, (int Row, int Column)? from, bool inPlace)
    {
        var value = EntryValue.Resolve(entry, context.Culture);
        // A value written elsewhere takes the style of the cell it came from: a moved card takes its formatting along.
        uint? style = !inPlace && from is { } origin ? source.Style(origin.Row, origin.Column) : null;

        if (value is not null && ReferenceEquals(value, entry.Original))
        {
            if (inPlace)
            {
                // An unedited value in its own cell: the cell is left as it is, formula and type included (§9.2–9.3).
                ApplyStyle(row, column, null, entry.Style, multiline: false);
                return;
            }

            if (from is { } position)
            {
                CopyCell(position.Row, position.Column, row, column);
                ApplyStyle(row, column, null, entry.Style, multiline: false);
                return;
            }
        }

        if (value is null)
        {
            Clear(row, column, style);
            ApplyStyle(row, column, null, entry.Style, multiline: false);
            return;
        }

        if (value.Formula is not null && ReferenceEquals(value, entry.Original))
        {
            // A copy of a card with no row as read (a card copied from a sheet without a source) holds the formula but not where it was
            // read from, so it cannot be translated.
            throw SpreadsheetWriteException.NotSupported(
                $"A copied card on the sheet \"{sheet.Name}\" holds a formula, which would go to the cell {CellIndex.Name(row, column)}; a copied formula cannot be written yet.");
        }

        var written = WriteValue(row, column, value, entry.Text, style);
        ApplyStyle(row, column, written, entry.Style, entry.Text.Contains('\n', StringComparison.Ordinal));
    }

    /// <summary>Writes a typed value; returns the number format to apply, or null to keep the cell's.</summary>
    private NumberFormat? WriteValue(int row, int column, CellValue value, string text, uint? style)
    {
        var cell = Prepare(row, column, style);
        if (value.Formula is { } formula)
        {
            // The formula without a cached value: the workbook is marked for a full calculation when it is opened (WriteContext).
            cell.CellFormula = new X.CellFormula(formula);
            return value.Format;
        }

        switch (value.Kind)
        {
            case CellValueKind.Number when value.Number is { } number:
                cell.CellValue = new X.CellValue(number.ToString("R", CultureInfo.InvariantCulture));
                return value.Format;
            case CellValueKind.Boolean when value.Boolean is { } flag:
                cell.DataType = X.CellValues.Boolean;
                cell.CellValue = new X.CellValue(flag ? "1" : "0");
                return value.Format;
            case CellValueKind.Date:
                double? serial = value.Number;
                if (serial is null && value.Date is { } date && DateSerial.TryFromDateTime(date, context.DateSystem, out var s))
                {
                    serial = s;
                }

                if (serial is { } days)
                {
                    cell.CellValue = new X.CellValue(days.ToString("R", CultureInfo.InvariantCulture));
                    return value.Format;
                }

                // A date before the date system's first day has no serial: it is kept as the text that was typed.
                SetText(cell, text);
                return null;
            case CellValueKind.Error:
                cell.DataType = X.CellValues.Error;
                cell.CellValue = new X.CellValue(value.DisplayText);
                return value.Format;
            default:
                SetText(cell, value.DisplayText);
                return value.Format;
        }
    }

    private void WriteText(int row, int column, string text)
    {
        if (text.Length == 0)
        {
            Clear(row, column, null);
            return;
        }

        SetText(Prepare(row, column, null), text);
    }

    private void SetText(X.Cell cell, string text)
    {
        cell.DataType = X.CellValues.SharedString;
        cell.CellValue = new X.CellValue(context.Strings.Add(text).ToString(CultureInfo.InvariantCulture));
    }

    /// <summary>
    /// The cell, emptied of its value and ready for a new one, with its style replaced when one is given. A cell that did not exist gets
    /// the style it showed (its row's or column's).
    /// </summary>
    private X.Cell Prepare(int row, int column, uint? style)
    {
        CheckChange(row, column);
        var existing = cells.Get(row, column);
        var cell = existing ?? cells.GetOrCreate(row, column);
        Empty(cell, row, column);
        SetStyle(cell, row, column, existing is null ? style ?? SourceStyle(row, column) : style);
        context.ValueChanged = true;
        return cell;
    }

    private void Clear(int row, int column, uint? style)
    {
        var cell = cells.Get(row, column);
        if (SourceHasValue(row, column) || (cell is not null && HasContent(cell)))
        {
            CheckChange(row, column);
            context.ValueChanged = true;
        }

        if (cell is null)
        {
            if (style is { } index && index != FallbackStyle(row, column))
            {
                SetStyle(cells.GetOrCreate(row, column), row, column, index);
            }

            return;
        }

        Empty(cell, row, column);
        SetStyle(cell, row, column, style);
    }

    /// <summary>
    /// Copies a cell as read (<paramref name="fromRow"/>, <paramref name="fromColumn"/>) to another cell after the move, with its value,
    /// type and style. A formula is translated by the distance it moves beyond where the row shifts alone would have put it, so its
    /// relative references move with it as in a copy; its cached value is dropped and the workbook is calculated when it is opened.
    /// </summary>
    private void CopyCell(int fromRow, int fromColumn, int row, int column)
    {
        var original = source.Element(fromRow, fromColumn);
        if (original is null)
        {
            Clear(row, column, source.Style(fromRow, fromColumn));
            return;
        }

        if (original.CellFormula is { } sourceFormula)
        {
            var type = sourceFormula.FormulaType?.Value;
            var text = source.Grid.Get(fromRow, fromColumn)?.Formula
                ?? (string.IsNullOrEmpty(sourceFormula.Text) ? null : sourceFormula.Text);
            if (type == X.CellFormulaValues.Array || type == X.CellFormulaValues.DataTable || text is null)
            {
                throw SpreadsheetWriteException.NotSupported(
                    $"An array or data table formula on the sheet \"{sheet.Name}\" would move to the cell {CellIndex.Name(row, column)}; it cannot be moved apart from its range.");
            }

            // A duplicated card may copy a card whose rows the save deletes: its formula is first moved, as a copy, to the first row after
            // them, so that its relative references (to its own row, say) are not lost with the deleted rows.
            var anchor = fromRow;
            while (plan is not null && plan.IsRemoved(anchor))
            {
                anchor++;
            }

            var moved = anchor == fromRow ? text : FormulaReferences.Translate(text, anchor - fromRow, 0);
            var shifted = context.References.Apply(moved, sheet.Name);
            var positional = plan?.Positional(anchor) ?? anchor;
            var formulaCell = Prepare(row, column, source.Style(fromRow, fromColumn));
            formulaCell.CellFormula = new X.CellFormula(FormulaReferences.Translate(shifted, row - positional, column - fromColumn));
            return;
        }

        var cell = Prepare(row, column, source.Style(fromRow, fromColumn));
        cell.DataType = original.DataType?.Value;
        if (original.CellValue is { } value)
        {
            cell.CellValue = (X.CellValue)value.CloneNode(true);
        }

        if (original.InlineString is { } inline)
        {
            cell.InlineString = (X.InlineString)inline.CloneNode(true);
        }
    }

    /// <summary>
    /// Refuses a change to a cell that would break other cells' formulas, or overwrite a value beside the table, and notes a changed
    /// formula cell (the calculation chain is then dropped).
    /// </summary>
    private void CheckChange(int row, int column)
    {
        if (Old(row) is { } old && source.IsProtected(old, column))
        {
            throw SpreadsheetWriteException.NotSupported(
                $"The cell {CellIndex.Name(row, column)} on the sheet \"{sheet.Name}\" is part of a shared or array formula of several cells; changing it needs the formula rewritten.");
        }

        if (column > sourceLastColumn && column != CellReference.MaxColumn && SourceHasValue(row, column))
        {
            throw SpreadsheetWriteException.NotSupported(
                $"An entry added to the table on the sheet \"{sheet.Name}\" would overwrite the cell {CellIndex.Name(row, column)} beside it.");
        }

        if (Old(row) is { } formulaRow && source.HasFormula(formulaRow, column))
        {
            context.FormulaChanged = true;
        }
    }

    /// <summary>The row as read whose content is at <paramref name="row"/> now, or null for a new row.</summary>
    private int? Old(int row) => plan is null ? row : plan.OldRow(row);

    /// <summary>True when the cell now at this position held a value or formula as read.</summary>
    private bool SourceHasValue(int row, int column) => Old(row) is { } old && source.HasValue(old, column);

    /// <summary>The style the cell now at this position showed as read; a new row shows its column's.</summary>
    private uint SourceStyle(int row, int column) => Old(row) is { } old ? source.Style(old, column) : source.ColumnStyle(column);

    /// <summary>The style a cell without its own style shows now: its row's (which moved with it), else its column's.</summary>
    private uint FallbackStyle(int row, int column) => Old(row) is { } old ? source.FallbackStyle(old, column) : source.ColumnStyle(column);

    private void ApplyStyle(int row, int column, NumberFormat? format, CellStyle? style, bool multiline)
    {
        if (format is null && style is null && !multiline)
        {
            return;
        }

        // A cell without its own style shows its row's or column's, as the reader reads it.
        var cell = cells.Get(row, column);
        var current = cell?.StyleIndex?.Value ?? FallbackStyle(row, column);
        var changed = context.Styles.Apply(current, format, style, multiline);
        if (changed != current)
        {
            SetStyle(cell ?? cells.GetOrCreate(row, column), row, column, changed);
        }
    }

    /// <summary>Sets a cell's style; style 0 is written explicitly only where the row or column would otherwise give another.</summary>
    private void SetStyle(X.Cell cell, int row, int column, uint? style)
    {
        if (style is { } index)
        {
            cell.StyleIndex = index == 0 && FallbackStyle(row, column) == 0 ? null : index;
        }
    }

    private static bool HasContent(X.Cell cell) => cell.CellValue is not null || cell.CellFormula is not null || cell.InlineString is not null;

    /// <summary>
    /// Removes the cell's value, formula and type; its style stays (as Excel's Clear Contents keeps formatting). A shared formula's master
    /// first gives the other cells of its group their own formula, so they keep theirs.
    /// </summary>
    private void Empty(X.Cell cell, int row, int column)
    {
        if (cell.CellFormula is { FormulaType.Value: var type, SharedIndex.Value: { } index } master
            && type == X.CellFormulaValues.Shared && !string.IsNullOrEmpty(master.Text))
        {
            foreach (var (r, c, other) in FormulaRewriter.Cells(worksheet))
            {
                if (other != cell && other.CellFormula is { SharedIndex.Value: var si } dependent && si == index && string.IsNullOrEmpty(dependent.Text))
                {
                    FormulaRewriter.Unshare(dependent, FormulaReferences.Translate(master.Text, r - row, c - column));
                }
            }

            context.FormulaChanged = true;
        }

        cell.CellFormula = null;
        cell.CellValue = null;
        cell.InlineString = null;
        cell.DataType = null;
        cell.CellMetaIndex = null;
        cell.ValueMetaIndex = null;
    }

    /// <summary>
    /// Clears what removed entries leave behind (§9.19): the columns the table had to the right of its new last column, over the table's
    /// rows from the header (or first card) down. Values go, styles stay; merged ranges that reach into the cleared area are removed.
    /// </summary>
    private void ClearRemovedColumns(SheetSource origin, RowPlan rows)
    {
        var firstTableRow = rows.HeaderRow ?? rows.FirstDataRow;
        var lastTableRow = rows.LastRow;
        var firstCleared = sheet.FirstColumn + sheet.Columns.Count;
        if (firstCleared > origin.LastColumn)
        {
            return;
        }

        for (var row = firstTableRow; row <= lastTableRow; row++)
        {
            for (var column = firstCleared; column <= origin.LastColumn; column++)
            {
                Clear(row, column, null);
            }
        }

        // The removed entries' pictures go with them (§9.19); pictures in the name or header row are not the app's and stay.
        drawing.RemoveOwned(rows.FirstDataRow, lastTableRow, firstCleared, origin.LastColumn);

        if (worksheet.GetFirstChild<X.MergeCells>() is not { } merges)
        {
            return;
        }

        foreach (var merge in merges.Elements<X.MergeCell>().ToList())
        {
            if (CellReference.TryParseRange(merge.Reference?.Value, out var range)
                && range.FirstRow <= lastTableRow && range.LastRow >= firstTableRow
                && range.FirstColumn <= origin.LastColumn && range.LastColumn >= firstCleared)
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

    /// <summary>A sheet without a source: the name row (when named), the header row (when shown), then one row per card.</summary>
    private void WriteNew()
    {
        var row = 1;
        if (!string.IsNullOrEmpty(sheet.TableName))
        {
            WriteText(row, sheet.TableNameColumn, sheet.TableName);
            row++;
        }

        if (sheet.HeaderRow is not null)
        {
            for (var j = 0; j < sheet.Columns.Count; j++)
            {
                var name = sheet.Columns[j].Name;
                WriteText(row, sheet.FirstColumn + j, IsUnnamed(name) ? "" : name);
            }

            row++;
        }

        foreach (var card in sheet.Cards)
        {
            for (var j = 0; j < sheet.Columns.Count; j++)
            {
                WriteEntry(card.Entries[j], row, sheet.FirstColumn + j, null, inPlace: false);
            }

            // A card with pictures takes the rows they need, merged in the table columns (§5.3 step 3); any other card takes one row.
            var layout = pictures.ForCard(card);
            var rows = new RowPlan.CardRows(row, layout.HasPictures ? Math.Max(1, layout.WorksheetRowsNeeded) : 1, null, 0);
            if (rows.Rows > 1)
            {
                MergeCard(rows);
            }

            WritePictures(card, rows, null);
            row += rows.Rows;
        }
    }

    /// <summary>The picture layout in the hidden cells of the last column, rows 1 to 3, with that column hidden (§5.2 step 2).</summary>
    private void WriteMetadata()
    {
        var layout = sheet.ImageLayout;
        var perRow = layout.ImagesPerRow * layout.Width > PictureLimits.MaxRowWidth
            ? PictureLimits.MaxImagesPerRow(layout.Width)
            : layout.ImagesPerRow;
        string[] texts =
        [
            LayoutMetadata.HeightPrefix + layout.Height.ToString(CultureInfo.InvariantCulture),
            LayoutMetadata.ImagesPerRowPrefix + perRow.ToString(CultureInfo.InvariantCulture),
            LayoutMetadata.WidthPrefix + layout.Width.ToString(CultureInfo.InvariantCulture),
        ];
        // The metadata cells are the app's own: writing them alone does not ask for the workbook's formulas to be calculated again.
        var valueChanged = context.ValueChanged;
        for (var i = 0; i < texts.Length; i++)
        {
            var row = i + 1;
            // A metadata cell that moved with its row was removed (RowMover): it is written again in its row.
            var original = Old(row) == row ? source.Grid.Get(row, CellReference.MaxColumn)?.PlainText : null;
            if (!string.Equals(original, texts[i], StringComparison.Ordinal))
            {
                WriteText(row, CellReference.MaxColumn, texts[i]);
            }
        }

        context.ValueChanged = valueChanged;

        HideLastColumn();
    }

    private void HideLastColumn()
    {
        const uint last = CellReference.MaxColumn;
        var columns = worksheet.GetFirstChild<X.Columns>();
        var covering = columns?.Elements<X.Column>().FirstOrDefault(c => c.Min?.Value <= last && c.Max?.Value >= last);
        if (covering is not null)
        {
            if (covering.Hidden?.Value == true)
            {
                return;
            }

            if (covering.Min!.Value < last)
            {
                // Split the range: the columns before the last keep their settings, the last becomes hidden with the same settings.
                var hidden = (X.Column)covering.CloneNode(true);
                covering.Max = last - 1;
                hidden.Min = last;
                hidden.Hidden = true;
                covering.InsertAfterSelf(hidden);
                return;
            }

            covering.Hidden = true;
            return;
        }

        if (columns is null)
        {
            columns = new X.Columns();
            worksheet.InsertBefore(columns, worksheet.GetFirstChild<X.SheetData>());
        }

        // Width 9.140625 is the width a column of the default font gets (not verified); a hidden column shows no width anyway.
        columns.AppendChild(new X.Column { Min = last, Max = last, Width = 9.140625, Hidden = true, CustomWidth = true });
    }

    /// <summary>An entry name that counts as unnamed: empty, or beginning with <c>Unnamed:</c> (docs/spec.md §3.3).</summary>
    private static bool IsUnnamed(string name) => name.Length == 0 || name.StartsWith("Unnamed:", StringComparison.Ordinal);
}
