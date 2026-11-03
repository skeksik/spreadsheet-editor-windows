using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Editing;

/// <summary>A card was removed from <see cref="WorkbookChange.Sheet"/>'s cards; it was at <paramref name="Index"/>.</summary>
public sealed record CardRemoved(Sheet Sheet, int Index, Card Card) : WorkbookChange(Sheet);
