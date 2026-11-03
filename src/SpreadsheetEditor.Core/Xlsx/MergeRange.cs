namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>A merged cell range on a worksheet, 1-based and inclusive.</summary>
internal readonly record struct MergeRange(int FirstRow, int FirstColumn, int LastRow, int LastColumn);
