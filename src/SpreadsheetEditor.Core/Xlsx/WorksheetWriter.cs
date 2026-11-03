using System.Globalization;
using DocumentFormat.OpenXml;
using SpreadsheetEditor.Core.Model;
using SpreadsheetEditor.Core.Values;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>
/// Writes one sheet's table into its worksheet's DOM (docs/spec.md §5.2, §5.5–5.6 with §9.2–9.3, §9.7, §9.16–9.19). The model is compared
/// with the worksheet as it was read (<see cref="SourceWorksheet"/>): a cell is changed only when its value, style or position changed,
/// so everything else stays as read. Cards are written over the rows they were read from; a change that needs rows inserted or deleted is
/// refused with <see cref="SpreadsheetWriteError.NotSupported"/> (part B of the writer).
/// </summary>
internal sealed class WorksheetWriter
{
    private readonly Sheet sheet;
    private readonly X.Worksheet worksheet;
    private readonly CellIndex cells;
    private readonly SourceWorksheet source;
    private readonly WriteContext context;

    /// <summary>The table's columns as read; cells to their right that are not part of the old table are not overwritten.</summary>
    private readonly int sourceLastColumn;

    private WorksheetWriter(Sheet sheet, X.Worksheet worksheet, SourceWorksheet source, WriteContext context)
    {
        this.sheet = sheet;
        this.worksheet = worksheet;
        this.source = source;
        this.context = context;
        cells = new CellIndex(worksheet);
        sourceLastColumn = sheet.Source?.LastColumn ?? int.MaxValue;
    }

    /// <summary>Writes a sheet read from <c>.xlsx</c> over its table as read, or a new sheet into an empty worksheet.</summary>
    public static void Write(Sheet sheet, X.Worksheet worksheet, SourceWorksheet source, WriteContext context)
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

        var writer = new WorksheetWriter(sheet, worksheet, source, context);
        if (sheet.Source is { } origin)
        {
            writer.WriteOverSource(origin);
        }
        else
        {
            writer.WriteNew();
        }

        writer.WriteMetadata();
    }

    private void WriteOverSource(SheetSource origin)
    {
        if (sheet.FirstColumn != origin.FirstColumn)
        {
            throw SpreadsheetWriteException.NotSupported($"The table on the sheet \"{sheet.Name}\" was moved to another column.");
        }

        var slots = CardGrouping.Group(sheet.FirstDataRow, origin.LastRow, origin.FirstColumn, origin.LastColumn, source.Grid.Merges);
        CheckRows(slots);

        var firstTableRow = sheet.HeaderRow ?? sheet.FirstDataRow;
        WriteName(origin, firstTableRow);
        WriteHeader(origin);

        for (var i = 0; i < slots.Count; i++)
        {
            WriteCard(sheet.Cards[i], slots[i].FirstRow, slots[i].RowSpan, origin);
        }

        ClearRemovedColumns(origin, firstTableRow);
    }

    /// <summary>
    /// Part A writes each card into one of the table's rows as read: the cards' row spans must be the spans of the table's cards as read,
    /// in order. Moving cards of equal spans is just writing values; anything else needs rows inserted or deleted.
    /// </summary>
    private void CheckRows(List<(int FirstRow, int RowSpan)> slots)
    {
        // A table with a header and no card was read with one empty card that has no row in the file; it can stay empty.
        if (slots.Count == 0 && sheet.Cards is [{ SourceFirstRow: null } only] && only.Entries.TrueForAll(e => e.Text.Length == 0))
        {
            return;
        }

        var matches = slots.Count == sheet.Cards.Count;
        for (var i = 0; matches && i < slots.Count; i++)
        {
            matches = slots[i].RowSpan == sheet.Cards[i].SourceRowSpan;
        }

        if (!matches)
        {
            throw SpreadsheetWriteException.NotSupported(
                $"The cards on the sheet \"{sheet.Name}\" were added, removed or reordered so that worksheet rows must be inserted or deleted.");
        }
    }

    private void WriteName(SheetSource origin, int firstTableRow)
    {
        var name = sheet.TableName ?? "";
        if (origin.FirstRow < firstTableRow)
        {
            var row = origin.FirstRow;
            var column = sheet.TableNameColumn;
            var read = source.Grid.Get(row, column);
            var original = read is null || read.IsBlank ? "" : read.Value.DisplayText;
            if (!string.Equals(name, original, StringComparison.Ordinal))
            {
                // A cleared name clears its cell; the row stays (§9.17).
                WriteText(row, column, name);
            }
        }
        else if (name.Length > 0)
        {
            throw SpreadsheetWriteException.NotSupported(
                $"The table on the sheet \"{sheet.Name}\" was given a name; a row must be inserted above its header for it.");
        }
    }

    private void WriteHeader(SheetSource origin)
    {
        if (sheet.HeaderRow is not { } row)
        {
            return;
        }

        for (var j = 0; j < sheet.Columns.Count; j++)
        {
            var column = sheet.Columns[j];
            var target = sheet.FirstColumn + j;
            var name = IsUnnamed(column.Name) ? "" : column.Name;
            if (column.SourceIndex is { } index)
            {
                var from = origin.FirstColumn + index;
                var read = source.Grid.Get(row, from);
                var original = read is null || read.IsBlank ? "" : read.Value.DisplayText;
                if (string.Equals(name, original, StringComparison.Ordinal))
                {
                    // An unchanged header keeps its cell as it is (a number stays a number); a moved one is copied (§9.3).
                    if (from != target)
                    {
                        CopyCell(row, from, row, target);
                    }

                    continue;
                }
            }

            WriteText(row, target, name);
        }
    }

    private void WriteCard(Card card, int row, int rowSpan, SheetSource origin)
    {
        var inOwnRows = card.SourceFirstRow == row;
        if (!inOwnRows && rowSpan > 1 && (HasLaterRowValues(row, rowSpan, origin) || (card.SourceFirstRow is { } sourceRow && HasLaterRowValues(sourceRow, rowSpan, origin))))
        {
            // Values in a card's later rows belong to it but are not in the model (§9.7); moving them needs whole rows moved (part B).
            throw SpreadsheetWriteException.NotSupported(
                $"A card on the sheet \"{sheet.Name}\" that spans several rows with values in its later rows was moved.");
        }

        for (var j = 0; j < sheet.Columns.Count; j++)
        {
            var entry = card.Entries[j];
            var target = sheet.FirstColumn + j;
            if (entry.Images.Count > 0 && (!inOwnRows || sheet.Columns[j].SourceIndex != j))
            {
                // Pictures are anchored in the drawing, not in the cell: moving them with their card or entry is part B.
                throw SpreadsheetWriteException.NotSupported(
                    $"A card or entry with pictures on the sheet \"{sheet.Name}\" was moved; pictures cannot be moved yet.");
            }

            (int Row, int Column)? from = card.SourceFirstRow is { } first && sheet.Columns[j].SourceIndex is { } index
                ? (first, origin.FirstColumn + index)
                : null;
            WriteEntry(entry, row, target, from);
        }
    }

    private bool HasLaterRowValues(int row, int rowSpan, SheetSource origin)
    {
        for (var r = row + 1; r < row + rowSpan; r++)
        {
            for (var c = origin.FirstColumn; c <= origin.LastColumn; c++)
            {
                if (source.HasValue(r, c))
                {
                    return true;
                }
            }
        }

        return false;
    }

    private void WriteEntry(Entry entry, int row, int column, (int Row, int Column)? from)
    {
        var value = EntryValue.Resolve(entry, context.Culture);
        var inPlace = from == (row, column);
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
            throw MovedFormula(row, column);
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
        Empty(cell);
        SetStyle(cell, row, column, existing is null ? style ?? source.Style(row, column) : style);
        context.ValueChanged = true;
        return cell;
    }

    private void Clear(int row, int column, uint? style)
    {
        var cell = cells.Get(row, column);
        if (source.HasValue(row, column) || (cell is not null && HasContent(cell)))
        {
            CheckChange(row, column);
            context.ValueChanged = true;
        }

        if (cell is null)
        {
            if (style is { } index && index != source.FallbackStyle(row, column))
            {
                SetStyle(cells.GetOrCreate(row, column), row, column, index);
            }

            return;
        }

        Empty(cell);
        SetStyle(cell, row, column, style);
    }

    private void CopyCell(int fromRow, int fromColumn, int row, int column)
    {
        if (source.HasFormula(fromRow, fromColumn))
        {
            throw MovedFormula(row, column);
        }

        var original = source.Element(fromRow, fromColumn);
        if (original is null)
        {
            Clear(row, column, source.Style(fromRow, fromColumn));
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
        if (source.IsProtected(row, column))
        {
            throw SpreadsheetWriteException.NotSupported(
                $"The cell {CellIndex.Name(row, column)} on the sheet \"{sheet.Name}\" is part of a shared or array formula of several cells; changing it needs the formula rewritten.");
        }

        if (column > sourceLastColumn && column != CellReference.MaxColumn && source.HasValue(row, column))
        {
            throw SpreadsheetWriteException.NotSupported(
                $"An entry added to the table on the sheet \"{sheet.Name}\" would overwrite the cell {CellIndex.Name(row, column)} beside it.");
        }

        if (source.HasFormula(row, column))
        {
            context.FormulaChanged = true;
        }
    }

    private SpreadsheetWriteException MovedFormula(int row, int column) =>
        SpreadsheetWriteException.NotSupported(
            $"A formula on the sheet \"{sheet.Name}\" would move to the cell {CellIndex.Name(row, column)}; its references would have to be rewritten.");

    private void ApplyStyle(int row, int column, NumberFormat? format, CellStyle? style, bool multiline)
    {
        if (format is null && style is null && !multiline)
        {
            return;
        }

        // A cell without its own style shows its row's or column's, as the reader reads it.
        var cell = cells.Get(row, column);
        var current = cell?.StyleIndex?.Value ?? source.FallbackStyle(row, column);
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
            cell.StyleIndex = index == 0 && source.FallbackStyle(row, column) == 0 ? null : index;
        }
    }

    private static bool HasContent(X.Cell cell) => cell.CellValue is not null || cell.CellFormula is not null || cell.InlineString is not null;

    /// <summary>Removes the cell's value, formula and type; its style stays (as Excel's Clear Contents keeps formatting).</summary>
    private static void Empty(X.Cell cell)
    {
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
    private void ClearRemovedColumns(SheetSource origin, int firstTableRow)
    {
        var firstCleared = sheet.FirstColumn + sheet.Columns.Count;
        if (firstCleared > origin.LastColumn)
        {
            return;
        }

        for (var row = firstTableRow; row <= origin.LastRow; row++)
        {
            for (var column = firstCleared; column <= origin.LastColumn; column++)
            {
                Clear(row, column, null);
            }
        }

        if (worksheet.GetFirstChild<X.MergeCells>() is not { } merges)
        {
            return;
        }

        foreach (var merge in merges.Elements<X.MergeCell>().ToList())
        {
            if (CellReference.TryParseRange(merge.Reference?.Value, out var range)
                && range.FirstRow <= origin.LastRow && range.LastRow >= firstTableRow
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
                WriteEntry(card.Entries[j], row, sheet.FirstColumn + j, null);
            }

            row++;
        }
    }

    /// <summary>The picture layout in the hidden cells of the last column, rows 1 to 3, with that column hidden (§5.2 step 2).</summary>
    private void WriteMetadata()
    {
        var layout = sheet.ImageLayout;
        var perRow = layout.ImagesPerRow * layout.Width > LayoutMetadata.MaxRowWidth
            ? LayoutMetadata.MaxImagesPerRow(layout.Width)
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
            if (!string.Equals(source.Grid.Get(row, CellReference.MaxColumn)?.PlainText, texts[i], StringComparison.Ordinal))
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
