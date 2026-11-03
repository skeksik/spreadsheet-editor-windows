namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>
/// The worksheet a sheet was read from and the region its table took there (1-based, inclusive), for the writer: it finds the worksheet
/// by <see cref="RelationshipId"/> and clears the old region before writing the table (docs/spec.md §9.19).
/// </summary>
/// <param name="RelationshipId">The workbook part's relationship id of the worksheet part.</param>
/// <param name="SheetId">The worksheet's <c>sheetId</c> in the workbook.</param>
/// <param name="FirstRow">The table's first row: its name row, else its header row, else its first card.</param>
/// <param name="LastRow">The last row of the last card.</param>
/// <param name="FirstColumn">The table's first column.</param>
/// <param name="LastColumn">The table's last column.</param>
internal sealed record SheetSource(string RelationshipId, uint SheetId, int FirstRow, int LastRow, int FirstColumn, int LastColumn);
