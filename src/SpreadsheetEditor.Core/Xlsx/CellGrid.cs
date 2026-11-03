namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>
/// The non-empty cells of one worksheet, looked up by position, and the merged ranges. Only the cells that hold a value are kept, as
/// small records; the worksheet's XML itself is not kept in memory.
/// </summary>
internal sealed class CellGrid
{
    private readonly Dictionary<long, ReadCell> cells = [];

    public List<MergeRange> Merges { get; } = [];

    public int Count => cells.Count;

    public IEnumerable<ReadCell> Cells => cells.Values;

    public void Add(ReadCell cell) => cells[Key(cell.Row, cell.Column)] = cell;

    public ReadCell? Get(int row, int column) => cells.GetValueOrDefault(Key(row, column));

    private static long Key(int row, int column) => ((long)row << 16) | (uint)column;
}
