using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>
/// The worksheet a sheet was read from and the region its table took there (1-based, inclusive), for the writer: it finds the worksheet
/// by <see cref="RelationshipId"/> and clears the old region before writing the table (docs/spec.md §9.19).
/// </summary>
/// <remarks>
/// A duplicated sheet carries its original's source marked as a copy (<c>Source = original.Source with { IsCopy = true }</c>): the writer
/// then copies the whole worksheet (other cells, column widths, page setup, drawings) into a new worksheet before writing the copy's
/// table (§9.30), and the original keeps its worksheet. A sheet whose source is not a copy is written into that worksheet in place.
/// </remarks>
/// <param name="RelationshipId">The workbook part's relationship id of the worksheet part.</param>
/// <param name="SheetId">The worksheet's <c>sheetId</c> in the workbook.</param>
/// <param name="FirstRow">The table's first row: its name row, else its header row, else its first card.</param>
/// <param name="LastRow">The last row of the last card.</param>
/// <param name="FirstColumn">The table's first column.</param>
/// <param name="LastColumn">The table's last column.</param>
/// <param name="IsCopy">True for a duplicated sheet: its worksheet is a copy of this one, made when the workbook is written.</param>
internal sealed record SheetSource(
    string RelationshipId, uint SheetId, int FirstRow, int LastRow, int FirstColumn, int LastColumn, bool IsCopy = false)
{
    /// <summary>
    /// The worksheet's shared formula groups and their cells, by worksheet position as read, over the whole sheet: the writer leaves an
    /// unedited cell of a group as it is, takes an edited dependent out of its group, and gives the other cells of an edited master's
    /// group their own formula.
    /// </summary>
    public SharedFormulaMap SharedFormulas { get; init; } = SharedFormulaMap.Empty;

    /// <summary>
    /// For a worksheet that had no table (no occupied cell outside the metadata column), the picture layout it was read with; null for a
    /// worksheet with a table. Such a worksheet opens as a new sheet's defaults (docs/spec.md §2.3, §9.10).
    /// </summary>
    public ImageLayout? EmptyLayout { get; init; }

    /// <summary>True when the worksheet had no table and opened as a new sheet's defaults (§9.10).</summary>
    public bool IsEmpty => EmptyLayout is not null;

    /// <summary>
    /// True when the worksheet had no table and <paramref name="sheet"/> is still exactly the defaults it opened as (§2.3), its picture
    /// layout included: the writer then leaves the worksheet as it was read (§9.10). Edits undone back to the defaults count as untouched.
    /// </summary>
    public bool IsUntouchedEmpty(Sheet sheet) =>
        EmptyLayout is { } layout
        && sheet.ImageLayout == layout
        && string.IsNullOrEmpty(sheet.TableName)
        && sheet.TableNameColumn == 1
        && sheet.FirstColumn == 1
        && sheet.HeaderRow == 1
        && sheet.FirstDataRow == 2
        && sheet.Columns is [{ Name: "", IsImageColumn: false, SourceIndex: null }]
        && sheet.Cards is [{ SourceFirstRow: null, CopiedFromRow: null, Entries: [{ Text: "", Original: null, Style: null, Images.Count: 0 }] }];
}
