using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Editing;

/// <summary>
/// One change an edit made to a sheet, raised by <see cref="WorkbookEditor.Changed"/> after the model has changed, for a do, an undo
/// and a redo alike (an undo raises the inverse change). View models update from it.
/// </summary>
/// <param name="Sheet">The sheet that changed.</param>
public abstract record WorkbookChange(Sheet Sheet);
