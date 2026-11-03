namespace SpreadsheetEditor.Core.Editing;

/// <summary>
/// The undo and redo stacks and the saved point. The document is unmodified exactly when the undo depth equals the depth at the last
/// load or save; once a new edit drops the redo steps that led back to that point, it can no longer be reached.
/// </summary>
internal sealed class UndoHistory
{
    private const int Unreachable = -1;

    private readonly List<IEdit> undo = [];
    private readonly List<IEdit> redo = [];
    private int savedDepth;
    private bool mergeOpen;

    public bool CanUndo => undo.Count > 0;

    public bool CanRedo => redo.Count > 0;

    public bool IsModified => undo.Count != savedDepth;

    /// <summary>
    /// Records an edit that has just been applied: merges it into the last step while typing continues, drops the redo steps. A merged
    /// step that now changes nothing is dropped, and the next edit starts a new step.
    /// </summary>
    public void Record(IEdit edit)
    {
        if (redo.Count > 0)
        {
            if (savedDepth > undo.Count)
            {
                savedDepth = Unreachable;
            }

            redo.Clear();
        }

        if (mergeOpen && undo.Count > 0 && undo.Count != savedDepth && undo[^1].TryMerge(edit))
        {
            if (undo[^1].IsEmpty)
            {
                undo.RemoveAt(undo.Count - 1);
                mergeOpen = false;
            }

            return;
        }

        undo.Add(edit);
        mergeOpen = true;
    }

    /// <summary>Ends the current typing step: the next edit starts a new one.</summary>
    public void BreakMerge() => mergeOpen = false;

    public IEdit? PopUndo()
    {
        mergeOpen = false;
        if (undo.Count == 0)
        {
            return null;
        }

        var edit = undo[^1];
        undo.RemoveAt(undo.Count - 1);
        redo.Add(edit);
        return edit;
    }

    public IEdit? PopRedo()
    {
        mergeOpen = false;
        if (redo.Count == 0)
        {
            return null;
        }

        var edit = redo[^1];
        redo.RemoveAt(redo.Count - 1);
        undo.Add(edit);
        return edit;
    }

    public void MarkSaved()
    {
        savedDepth = undo.Count;
        mergeOpen = false;
    }

    public void Clear()
    {
        undo.Clear();
        redo.Clear();
        savedDepth = 0;
        mergeOpen = false;
    }
}
