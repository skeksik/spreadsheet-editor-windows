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
    string RelationshipId, uint SheetId, int FirstRow, int LastRow, int FirstColumn, int LastColumn, bool IsCopy = false);
