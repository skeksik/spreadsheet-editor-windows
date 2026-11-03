using SpreadsheetEditor.Core.Xlsx;

namespace SpreadsheetEditor.Core.Formulas;

/// <summary>
/// Worksheet rows inserted into or deleted from one sheet, 1-based. Inserting <see cref="Count"/> rows at <see cref="FirstRow"/> puts
/// the new rows at <c>FirstRow … FirstRow + Count − 1</c> and moves the old row <c>FirstRow</c> and everything below it down by
/// <see cref="Count"/>. Deleting removes rows <c>FirstRow … FirstRow + Count − 1</c> and moves everything below them up.
/// </summary>
public readonly record struct RowShift
{
    /// <summary>The last worksheet row. Microsoft, "Excel specifications and limits" (read 2026-11-03): "Total number of rows and
    /// columns on a worksheet | 1,048,576 rows by 16,384 columns".</summary>
    public const int MaxRow = CellReference.MaxRow;

    private RowShift(string sheet, int firstRow, int count, RowShiftKind kind)
    {
        ArgumentNullException.ThrowIfNull(sheet);
        ArgumentOutOfRangeException.ThrowIfLessThan(firstRow, 1);
        ArgumentOutOfRangeException.ThrowIfGreaterThan(firstRow, MaxRow);
        ArgumentOutOfRangeException.ThrowIfLessThan(count, 1);
        ArgumentOutOfRangeException.ThrowIfGreaterThan(count, MaxRow - firstRow + 1);
        Sheet = sheet;
        FirstRow = firstRow;
        Count = count;
        Kind = kind;
    }

    /// <summary>The name of the worksheet whose rows change.</summary>
    public string Sheet { get; }

    /// <summary>The first inserted or deleted row.</summary>
    public int FirstRow { get; }

    /// <summary>How many rows are inserted or deleted, at least 1.</summary>
    public int Count { get; }

    public RowShiftKind Kind { get; }

    /// <summary>The last deleted row (or the last new row, for an insertion).</summary>
    public int LastRow => FirstRow + Count - 1;

    /// <summary>Inserts <paramref name="count"/> rows at <paramref name="firstRow"/> on <paramref name="sheet"/>.</summary>
    public static RowShift Insert(string sheet, int firstRow, int count) => new(sheet, firstRow, count, RowShiftKind.Insert);

    /// <summary>Deletes <paramref name="count"/> rows from <paramref name="firstRow"/> on <paramref name="sheet"/>.</summary>
    public static RowShift Delete(string sheet, int firstRow, int count) => new(sheet, firstRow, count, RowShiftKind.Delete);
}
