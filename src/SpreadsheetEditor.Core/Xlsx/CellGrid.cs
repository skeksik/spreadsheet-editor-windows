namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>
/// The non-empty cells of one worksheet, looked up by position, and the merged ranges. Only the cells that hold a value are kept, as
/// small records; the worksheet's XML itself is not kept in memory.
/// </summary>
internal sealed class CellGrid
{
    private readonly Dictionary<long, ReadCell> cells = [];

    private readonly Dictionary<long, uint> cellStyles = [];
    private readonly Dictionary<int, uint> rowStyles = [];
    private readonly List<(int First, int Last, uint Style)> columnStyles = [];

    public List<MergeRange> Merges { get; } = [];

    public int Count => cells.Count;

    public IEnumerable<ReadCell> Cells => cells.Values;

    public void Add(ReadCell cell) => cells[Key(cell.Row, cell.Column)] = cell;

    public ReadCell? Get(int row, int column) => cells.GetValueOrDefault(Key(row, column));

    /// <summary>Records a cell's <c>s</c> attribute, whether or not the cell holds a value.</summary>
    public void SetCellStyle(int row, int column, uint style) => cellStyles[Key(row, column)] = style;

    /// <summary>Records a row's style (<c>s</c> with <c>customFormat="1"</c>), which its cells without their own style take.</summary>
    public void SetRowStyle(int row, uint style) => rowStyles[row] = style;

    /// <summary>Records a <c>col</c> element's style for columns <paramref name="first"/> to <paramref name="last"/>.</summary>
    public void SetColumnStyle(int first, int last, uint style) => columnStyles.Add((first, last, style));

    /// <summary>
    /// The cell style index (<c>cellXfs</c>) a cell shows: its own <c>s</c>, else its row's style when the row has a custom format, else
    /// its column's style, else 0. The order follows how a cell without its own style takes its row's or column's formatting (not
    /// verified against ECMA-376's text).
    /// </summary>
    public uint StyleIndex(int row, int column)
    {
        if (cellStyles.TryGetValue(Key(row, column), out var style) || rowStyles.TryGetValue(row, out style))
        {
            return style;
        }
        foreach (var (first, last, columnStyle) in columnStyles)
        {
            if (column >= first && column <= last)
            {
                return columnStyle;
            }
        }
        return 0;
    }

    private static long Key(int row, int column) => ((long)row << 16) | (uint)column;
}
