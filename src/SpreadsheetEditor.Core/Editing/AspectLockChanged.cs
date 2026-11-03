using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Editing;

/// <summary>The sheet's aspect lock was turned on or off (<see cref="Sheet.LockedAspectRatio"/>). Not an undoable step.</summary>
public sealed record AspectLockChanged(Sheet Sheet) : WorkbookChange(Sheet);
