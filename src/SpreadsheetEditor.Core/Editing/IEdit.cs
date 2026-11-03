namespace SpreadsheetEditor.Core.Editing;

/// <summary>One undoable step: applies its change to the model, or reverts it, and says what changed.</summary>
internal interface IEdit
{
    IReadOnlyList<WorkbookChange> Apply();

    IReadOnlyList<WorkbookChange> Revert();

    /// <summary>
    /// Folds <paramref name="next"/>, already applied, into this step so that one undo reverts both (typing). Returns false when the two
    /// are separate steps.
    /// </summary>
    bool TryMerge(IEdit next);
}
