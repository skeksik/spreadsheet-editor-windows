using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Editing;

/// <summary>A sheet's <see cref="Sheet.Name"/> changed.</summary>
public sealed record SheetRenamed(Sheet Sheet) : WorkbookChange(Sheet);
