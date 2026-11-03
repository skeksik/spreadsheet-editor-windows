namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>Where a worksheet's table is (docs/spec.md §4.1). Rows and columns are 1-based and inclusive.</summary>
/// <param name="FirstRow">The first row of the row block: the name row, else the header row.</param>
/// <param name="LastRow">The last row of the row block.</param>
/// <param name="FirstColumn">The first column of the column block.</param>
/// <param name="LastColumn">The last column of the column block.</param>
/// <param name="Name">The table name cell, or null when the table has none.</param>
/// <param name="HeaderRow">The row of the header cells (read as data when <paramref name="HeadersHidden"/>).</param>
/// <param name="HeadersHidden">True when every non-blank header cell is a number (§9.9).</param>
internal sealed record DetectedTable(
    int FirstRow, int LastRow, int FirstColumn, int LastColumn, ReadCell? Name, int HeaderRow, bool HeadersHidden)
{
    /// <summary>The first card's row: the header row itself when the headers are hidden, else the row below it.</summary>
    public int FirstDataRow => HeadersHidden ? HeaderRow : HeaderRow + 1;
}
