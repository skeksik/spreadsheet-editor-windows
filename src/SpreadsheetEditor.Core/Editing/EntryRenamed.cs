using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Editing;

/// <summary>A column's name changed; every card shows the new <see cref="Column.Name"/>.</summary>
public sealed record EntryRenamed(Sheet Sheet, Column Column) : WorkbookChange(Sheet);
