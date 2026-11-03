using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>Builds a <see cref="Sheet"/> from a worksheet's cells: its table, entries, cards and picture layout (docs/spec.md §4.1–4.5).</summary>
internal static class SheetBuilder
{
    /// <summary>
    /// The sheet. A worksheet without a table (no occupied cell outside the metadata column) gives a new sheet's defaults with the picture
    /// layout of its metadata cells (docs/spec.md §2.3, §9.10).
    /// </summary>
    /// <param name="name">The worksheet's name.</param>
    /// <param name="relationshipId">The workbook part's relationship id of the worksheet part.</param>
    /// <param name="sheetId">The worksheet's <c>sheetId</c>.</param>
    /// <param name="grid">The worksheet's cells, merges and style indexes.</param>
    /// <param name="styles">The workbook's cell styles.</param>
    /// <param name="pictures">The pictures anchored to the worksheet's cells, in file order.</param>
    public static Sheet Build(string name, string relationshipId, uint sheetId, CellGrid grid, StyleTable styles, IReadOnlyList<AnchoredPicture> pictures)
    {
        var layout = LayoutMetadata.Read(
            grid.Get(1, CellReference.MaxColumn)?.PlainText,
            grid.Get(2, CellReference.MaxColumn)?.PlainText,
            grid.Get(3, CellReference.MaxColumn)?.PlainText);
        var table = TableDetector.Detect(grid, pictures.Select(p => (p.Row, p.Column)));
        if (table is null)
        {
            return BuildEmpty(name, relationshipId, sheetId, grid, layout);
        }

        var sheet = new Sheet(name)
        {
            TableName = table.Name?.Value.DisplayText,
            TableNameColumn = table.Name?.Column ?? table.FirstColumn,
            FirstColumn = table.FirstColumn,
            HeaderRow = table.HeadersHidden ? null : table.HeaderRow,
            FirstDataRow = table.FirstDataRow,
            ImageLayout = layout,
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
                var entry = CreateEntry(grid.Get(firstRow, column));
                entry.Style = styles.Style(grid.StyleIndex(firstRow, column));
                card.Entries.Add(entry);
            }
            sheet.Cards.Add(card);
            lastRow = Math.Max(lastRow, firstRow + rowSpan - 1);
        }

        AssignPictures(sheet, table, pictures);

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

        sheet.Source = new SheetSource(relationshipId, sheetId, table.FirstRow, lastRow, table.FirstColumn, table.LastColumn) { SharedFormulas = grid.SharedFormulas };
        return sheet;
    }

    /// <summary>
    /// A new sheet's defaults (§2.3): one unnamed entry and one empty card, headers shown, no table name, the table at A1. The source marks
    /// the worksheet as empty, so that a save without edits leaves it as it was read (§9.10).
    /// </summary>
    private static Sheet BuildEmpty(string name, string relationshipId, uint sheetId, CellGrid grid, ImageLayout layout)
    {
        var sheet = new Sheet(name) { ImageLayout = layout };
        sheet.Columns.Add(new Column(""));
        var card = new Card();
        card.Entries.Add(new Entry());
        sheet.Cards.Add(card);
        sheet.Source = new SheetSource(relationshipId, sheetId, 1, 1, 1, 1) { SharedFormulas = grid.SharedFormulas, EmptyLayout = layout };
        return sheet;
    }

    /// <summary>
    /// Gives each picture anchored inside the column block to the entry of its column on the card whose rows hold its anchor row, in
    /// worksheet-row order and then file order (§4.1 step 6, §4.2 step 3). A column that receives a picture becomes an image column.
    /// Pictures anchored outside the block, or in the name or header row, are not shown and stay in the package.
    /// </summary>
    private static void AssignPictures(Sheet sheet, DetectedTable table, IReadOnlyList<AnchoredPicture> pictures)
    {
        if (pictures.Count == 0 || sheet.Cards.Count == 0)
        {
            return;
        }
        var firstRows = sheet.Cards.Select(c => c.SourceFirstRow!.Value).ToArray();
        // OrderBy is a stable sort: pictures on the same row keep their file order.
        foreach (var picture in pictures.OrderBy(p => p.Row))
        {
            if (picture.Column < table.FirstColumn || picture.Column > table.LastColumn)
            {
                continue;
            }
            var index = Array.BinarySearch(firstRows, picture.Row);
            if (index < 0)
            {
                index = ~index - 1;
            }
            if (index < 0)
            {
                continue;
            }
            var card = sheet.Cards[index];
            if (picture.Row >= card.SourceFirstRow + card.SourceRowSpan)
            {
                continue;
            }
            var column = picture.Column - table.FirstColumn;
            card.Entries[column].Images.Add(picture.Image);
            sheet.Columns[column].IsImageColumn = true;
        }
    }

    private static Entry CreateEntry(ReadCell? cell) =>
        cell is null ? new Entry() : new Entry { Text = cell.Value.DisplayText, Original = cell.Value };
}
