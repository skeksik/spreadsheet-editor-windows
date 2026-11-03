using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Editing;

/// <summary>A card moved from <paramref name="OldIndex"/> to <paramref name="NewIndex"/> in <see cref="WorkbookChange.Sheet"/>'s cards.</summary>
public sealed record CardMoved(Sheet Sheet, Card Card, int OldIndex, int NewIndex) : WorkbookChange(Sheet);
