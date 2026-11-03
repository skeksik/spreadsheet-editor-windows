using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>Builds a <see cref="Sheet"/> from a worksheet's cells: its table, entries, cards and picture layout (docs/spec.md §4.1–4.5).</summary>
internal static class SheetBuilder
{
    /// <summary>The sheet, or null when the worksheet has no table (no occupied cell).</summary>
    public static Sheet? Build(string name, string relationshipId, uint sheetId, CellGrid grid)
    {
        var table = TableDetector.Detect(grid, []);
        if (table is null)
        {
            return null;
        }

        var sheet = new Sheet(name)
        {
            TableName = table.Name?.Value.DisplayText,
            TableNameColumn = table.Name?.Column ?? table.FirstColumn,
            FirstColumn = table.FirstColumn,
            HeaderRow = table.HeadersHidden ? null : table.HeaderRow,
            FirstDataRow = table.FirstDataRow,
            ImageLayout = LayoutMetadata.Read(
                grid.Get(1, CellReference.MaxColumn)?.PlainText,
                grid.Get(2, CellReference.MaxColumn)?.PlainText,
                grid.Get(3, CellReference.MaxColumn)?.PlainText),
        };

        for (var column = table.FirstColumn; column <= table.LastColumn; column++)
        {
            // Hidden headers: the entries are known by position only, and the header row is the first card (docs/spec.md §3.3).
            var header = table.HeadersHidden ? null : grid.Get(table.HeaderRow, column);
            sheet.Columns.Add(new Column(header is null || header.IsBlank ? "" : header.Value.DisplayText) { SourceIndex = column - table.FirstColumn });
        }

        // Values come from a card's first row only; its later rows stay in the file as they are (§9.7).
        var lastRow = table.LastRow;
        foreach (var (firstRow, rowSpan) in CardGrouping.Group(table.FirstDataRow, table.LastRow, table.FirstColumn, table.LastColumn, grid.Merges))
        {
            var card = new Card { SourceFirstRow = firstRow, SourceRowSpan = rowSpan };
            for (var column = table.FirstColumn; column <= table.LastColumn; column++)
            {
                card.Entries.Add(CreateEntry(grid.Get(firstRow, column)));
            }
            sheet.Cards.Add(card);
            lastRow = Math.Max(lastRow, firstRow + rowSpan - 1);
        }

        if (sheet.Cards.Count == 0)
        {
            // A table with a header and no data row still shows one empty card; it was not in the file, so it has no source row.
            var card = new Card();
            for (var column = table.FirstColumn; column <= table.LastColumn; column++)
            {
                card.Entries.Add(new Entry());
            }
            sheet.Cards.Add(card);
        }

        sheet.Source = new SheetSource(relationshipId, sheetId, table.FirstRow, lastRow, table.FirstColumn, table.LastColumn);
        return sheet;
    }

    private static Entry CreateEntry(ReadCell? cell) =>
        cell is null ? new Entry() : new Entry { Text = cell.Value.DisplayText, Original = cell.Value };
}
