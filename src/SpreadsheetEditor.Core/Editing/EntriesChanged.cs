using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Editing;

/// <summary>
/// The sheet's columns were added, removed or reordered, and every card's entries with them: view models rebuild their entry rows from
/// <see cref="Sheet.Columns"/> and each card's <see cref="Card.Entries"/>.
/// </summary>
public sealed record EntriesChanged(Sheet Sheet) : WorkbookChange(Sheet);
