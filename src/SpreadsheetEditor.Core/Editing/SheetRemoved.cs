using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Editing;

/// <summary>A sheet was removed from the workbook's sheets; it was at <paramref name="Index"/>.</summary>
public sealed record SheetRemoved(Sheet Sheet, int Index) : WorkbookChange(Sheet);
