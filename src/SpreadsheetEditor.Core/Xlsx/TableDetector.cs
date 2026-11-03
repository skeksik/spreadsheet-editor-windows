using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>Finds the table on a worksheet (docs/spec.md §4.1, with §9.1 and §9.9).</summary>
internal static class TableDetector
{
    /// <summary>More than this many rows between two occupied rows ends the row block (up to 9 empty rows are still one block).</summary>
    public const int MaxRowStep = 10;

    /// <summary>More than this many columns between two occupied columns ends the column block (up to 2 empty columns are still one block).</summary>
    public const int MaxColumnStep = 3;

    /// <summary>
    /// The table, or null when the worksheet has no occupied cell. A cell is occupied when it holds a value, except in the last column
    /// (<c>XFD</c>, where the layout metadata lives) and except a cell holding a metadata string (§9.1). The top-left cells of pictures
    /// anchored to cells are occupied too (<paramref name="anchors"/>).
    /// </summary>
    public static DetectedTable? Detect(CellGrid grid, IEnumerable<(int Row, int Column)> anchors)
    {
        var occupied = new List<(int Row, int Column)>();
        foreach (var cell in grid.Cells)
        {
            if (IsOccupied(cell))
            {
                occupied.Add((cell.Row, cell.Column));
            }
        }
        foreach (var anchor in anchors)
        {
            if (anchor.Column < CellReference.MaxColumn)
            {
                occupied.Add(anchor);
            }
        }
        if (occupied.Count == 0)
        {
            return null;
        }

        var rows = Block(occupied.Select(p => p.Row), MaxRowStep);
        var columns = Block(occupied.Where(p => p.Row >= rows.First && p.Row <= rows.Last).Select(p => p.Column), MaxColumnStep);

        // Table name: the first row has exactly one non-blank cell, there is a row below it in the block, and that row has more than one.
        var named = rows.First < rows.Last && NonBlank(grid, rows.First, columns) is [var name] && NonBlank(grid, rows.First + 1, columns).Count > 1
            ? name
            : null;
        var headerRow = named is null ? rows.First : rows.First + 1;

        // Hidden headers: every non-blank header cell is a number. Booleans, dates, text and errors are not numbers; a formula counts by
        // its cached value (§9.9).
        var header = NonBlank(grid, headerRow, columns);
        var hidden = header.Count > 0 && header.TrueForAll(c => c.Value.Kind == CellValueKind.Number);

        return new DetectedTable(rows.First, rows.Last, columns.First, columns.Last, named, headerRow, hidden);
    }

    private static bool IsOccupied(ReadCell cell) =>
        cell.Column < CellReference.MaxColumn && !LayoutMetadata.IsMetadata(cell.PlainText);

    /// <summary>From the smallest value up to the last one before the first step of more than <paramref name="maxStep"/>.</summary>
    private static (int First, int Last) Block(IEnumerable<int> values, int maxStep)
    {
        var sorted = values.Distinct().Order().ToList();
        var last = sorted[0];
        for (var i = 1; i < sorted.Count && sorted[i] - last <= maxStep; i++)
        {
            last = sorted[i];
        }
        return (sorted[0], last);
    }

    /// <summary>The non-blank cells of a row inside the column block, ignoring metadata cells.</summary>
    private static List<ReadCell> NonBlank(CellGrid grid, int row, (int First, int Last) columns)
    {
        var cells = new List<ReadCell>();
        for (var column = columns.First; column <= columns.Last; column++)
        {
            if (grid.Get(row, column) is { IsBlank: false } cell && IsOccupied(cell))
            {
                cells.Add(cell);
            }
        }
        return cells;
    }
}
