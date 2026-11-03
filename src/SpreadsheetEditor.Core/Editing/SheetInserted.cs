using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Editing;

/// <summary>A sheet was inserted into the workbook's sheets at <paramref name="Index"/> (added, duplicated, or put back by an undo).</summary>
public sealed record SheetInserted(Sheet Sheet, int Index) : WorkbookChange(Sheet);
