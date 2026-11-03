namespace SpreadsheetEditor.Core.Editing;

/// <summary>An edit made of an apply and a revert action; never merges.</summary>
internal sealed class DelegateEdit(Func<WorkbookChange[]> apply, Func<WorkbookChange[]> revert) : IEdit
{
    public IReadOnlyList<WorkbookChange> Apply() => apply();

    public IReadOnlyList<WorkbookChange> Revert() => revert();

    public bool TryMerge(IEdit next) => false;
}
