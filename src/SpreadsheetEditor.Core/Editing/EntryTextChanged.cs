using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Editing;

/// <summary>An entry's text changed: <see cref="Entry.Text"/> of <paramref name="Card"/>'s entry in <paramref name="Column"/>.</summary>
public sealed record EntryTextChanged(Sheet Sheet, Card Card, Column Column) : WorkbookChange(Sheet);
