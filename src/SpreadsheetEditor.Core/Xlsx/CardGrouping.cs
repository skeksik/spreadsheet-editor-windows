namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>
/// Groups a table's worksheet rows into cards by vertical merges (docs/spec.md §4.3, §9.7). Only a merged range that lies wholly inside
/// the table's columns groups rows: a range that reaches outside them, even partly, is ignored here and kept as it is in the file.
/// </summary>
internal static class CardGrouping
{
    /// <summary>The cards from <paramref name="firstRow"/> through <paramref name="lastRow"/>, as (first row, row count) pairs.</summary>
    public static List<(int FirstRow, int RowSpan)> Group(
        int firstRow, int lastRow, int firstColumn, int lastColumn, IEnumerable<MergeRange> merges)
    {
        var vertical = merges
            .Where(m => m.FirstColumn >= firstColumn && m.LastColumn <= lastColumn && m.LastRow > m.FirstRow && m.LastRow >= firstRow)
            .OrderBy(m => m.FirstRow)
            .ToList();
        var cards = new List<(int, int)>();
        var next = 0;
        var row = firstRow;
        while (row <= lastRow)
        {
            // Every range that overlaps rows row..end and reaches below end extends the card, until none does. Ranges are visited in
            // order of their first row; one that ends above the card can never overlap a later card, so each is looked at once.
            var end = row;
            while (next < vertical.Count && vertical[next].FirstRow <= end)
            {
                if (vertical[next].LastRow > end)
                {
                    end = vertical[next].LastRow;
                }
                next++;
            }
            cards.Add((row, end - row + 1));
            row = end + 1;
        }
        return cards;
    }
}
