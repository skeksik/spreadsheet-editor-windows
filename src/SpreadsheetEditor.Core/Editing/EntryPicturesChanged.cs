using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Editing;

/// <summary>Pictures were added to, removed from or reordered in <paramref name="Card"/>'s entry in <paramref name="Column"/>; <see cref="Column.IsImageColumn"/> may have changed with them.</summary>
public sealed record EntryPicturesChanged(Sheet Sheet, Card Card, Column Column) : WorkbookChange(Sheet);
