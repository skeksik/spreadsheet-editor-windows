namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>
/// The shared formula groups of one worksheet and the cells in them, by worksheet position as read (1-based), for the whole sheet and
/// not only the table. The reader shows every cell of a group as its own formula (<see cref="WorksheetCells"/>); the writer uses this
/// map to tell a group's cells from plain formula cells: an unedited cell stays as it was, an edited dependent leaves the group, and
/// an edited master must give the group's other cells their own formula.
/// </summary>
internal sealed class SharedFormulaMap
{
    private readonly Dictionary<uint, SharedFormulaGroup> groups = [];
    private readonly Dictionary<long, SharedFormulaCell> cells = [];

    /// <summary>A map without groups.</summary>
    public static SharedFormulaMap Empty { get; } = new();

    /// <summary>The groups whose master was found, by <c>si</c>.</summary>
    public IReadOnlyDictionary<uint, SharedFormulaGroup> Groups => groups;

    /// <summary>The number of cells that belong to a group, masters included.</summary>
    public int CellCount => cells.Count;

    /// <summary>
    /// The group membership of the cell at <paramref name="row"/>, <paramref name="column"/>, or null when it is not in a group. A
    /// dependent whose master the worksheet does not have is still listed, with a <c>si</c> missing from <see cref="Groups"/>.
    /// </summary>
    public SharedFormulaCell? Get(int row, int column) =>
        cells.TryGetValue(Key(row, column), out var cell) ? cell : null;

    /// <summary>The group with this <c>si</c>, or null when its master was not found.</summary>
    public SharedFormulaGroup? Group(uint index) => groups.GetValueOrDefault(index);

    /// <summary>Adds a group; the first master read for a <c>si</c> wins. Returns whether it was added.</summary>
    internal bool TryAddGroup(SharedFormulaGroup group) => groups.TryAdd(group.Index, group);

    internal void AddCell(int row, int column, SharedFormulaCell cell) => cells[Key(row, column)] = cell;

    private static long Key(int row, int column) => ((long)row << 16) | (uint)column;
}
