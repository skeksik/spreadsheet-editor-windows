using DocumentFormat.OpenXml.Packaging;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>
/// A worksheet as it is in the source package, never changed, for the writer to compare the model with: its cells as the reader reads
/// them (values and display texts), its cell elements (to copy a moved value with its type and style), and its formula ranges.
/// </summary>
internal sealed class SourceWorksheet
{
    private readonly Dictionary<long, X.Cell> elements = [];
    private readonly List<MergeRange> protectedRanges = [];

    public SourceWorksheet(WorksheetPart part, WorksheetCells reader, CancellationToken cancellationToken)
    {
        Grid = reader.Read(part, cancellationToken);
        var index = new CellIndex(part.Worksheet ?? new X.Worksheet());
        foreach (var (row, column, cell) in index.All)
        {
            elements[Key(row, column)] = cell;
            if (cell.CellFormula is not { } formula)
            {
                continue;
            }

            // An array or data table formula fills its whole range: changing one of its cells would break the others. (A shared
            // formula's master gives the other cells of its group their own formula when it is changed, WorksheetWriter.)
            var type = formula.FormulaType?.Value;
            if ((type == X.CellFormulaValues.Array || type == X.CellFormulaValues.DataTable)
                && CellReference.TryParseRange(formula.Reference?.Value, out var range)
                && (range.FirstRow != range.LastRow || range.FirstColumn != range.LastColumn))
            {
                protectedRanges.Add(range);
            }
        }
    }

    /// <summary>An empty worksheet, for a sheet that has no source.</summary>
    private SourceWorksheet() => Grid = new CellGrid();

    public static SourceWorksheet Empty { get; } = new();

    public CellGrid Grid { get; }

    /// <summary>The style a cell shows as read: its own, else its row's or column's (<see cref="CellGrid.StyleIndex"/>).</summary>
    public uint Style(int row, int column) => Grid.StyleIndex(row, column);

    /// <summary>The style a cell without its own style shows: its row's or column's, else 0.</summary>
    public uint FallbackStyle(int row, int column) => Grid.FallbackStyleIndex(row, column);

    /// <summary>The style a cell of a new row shows: its column's, else 0.</summary>
    public uint ColumnStyle(int column) => Grid.FallbackStyleIndex(0, column);

    /// <summary>The cell element at a position, or null.</summary>
    public X.Cell? Element(int row, int column) => elements.GetValueOrDefault(Key(row, column));

    /// <summary>True when the cell holds a value or a formula.</summary>
    public bool HasValue(int row, int column) => Grid.Get(row, column) is not null || Element(row, column)?.CellFormula is not null;

    /// <summary>True when the cell has a formula, including a cell of a shared formula that carries no text of its own.</summary>
    public bool HasFormula(int row, int column) => Element(row, column)?.CellFormula is not null;

    /// <summary>
    /// True when changing the cell would break a formula of other cells: it lies in an array or data table formula's range of more than
    /// one cell.
    /// </summary>
    public bool IsProtected(int row, int column) =>
        protectedRanges.Exists(r => row >= r.FirstRow && row <= r.LastRow && column >= r.FirstColumn && column <= r.LastColumn);

    private static long Key(int row, int column) => ((long)row << 16) | (uint)column;
}
