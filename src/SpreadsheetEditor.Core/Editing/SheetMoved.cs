using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Editing;

/// <summary>A sheet moved from <paramref name="OldIndex"/> to <paramref name="NewIndex"/> in the workbook's sheets.</summary>
public sealed record SheetMoved(Sheet Sheet, int OldIndex, int NewIndex) : WorkbookChange(Sheet);
