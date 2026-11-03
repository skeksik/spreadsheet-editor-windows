using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Editing;

/// <summary>An entry's fill or text colour changed: <see cref="Entry.Style"/> of <paramref name="Card"/>'s entry in <paramref name="Column"/>.</summary>
public sealed record EntryStyleChanged(Sheet Sheet, Card Card, Column Column) : WorkbookChange(Sheet);
