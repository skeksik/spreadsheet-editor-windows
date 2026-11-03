using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Editing;

/// <summary>A card was inserted into <see cref="WorkbookChange.Sheet"/>'s cards at <paramref name="Index"/>.</summary>
public sealed record CardInserted(Sheet Sheet, int Index, Card Card) : WorkbookChange(Sheet);
