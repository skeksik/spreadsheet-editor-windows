namespace SpreadsheetEditor.Core.Editing;

/// <summary>The changes one edit, undo or redo made, in the order they were made.</summary>
public sealed class WorkbookChangedEventArgs : EventArgs
{
    public WorkbookChangedEventArgs(IReadOnlyList<WorkbookChange> changes) => Changes = changes;

    public IReadOnlyList<WorkbookChange> Changes { get; }
}
