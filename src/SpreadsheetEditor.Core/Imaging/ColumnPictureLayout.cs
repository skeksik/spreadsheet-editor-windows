namespace SpreadsheetEditor.Core.Imaging;

/// <summary>
/// One table column's picture layout for a sheet (docs/spec.md §5.2 steps 6–8).
/// </summary>
/// <param name="MaxUnits">The largest total of picture units in any one card's value in this column; 0 when it has no pictures.</param>
/// <param name="Capacity">How many picture units fit on one picture row of this column (colP).</param>
/// <param name="WidthTargetPx">The width the column needs, Capacity × picture width in pixels, or null when it has no pictures.</param>
public sealed record ColumnPictureLayout(int MaxUnits, int Capacity, int? WidthTargetPx);
