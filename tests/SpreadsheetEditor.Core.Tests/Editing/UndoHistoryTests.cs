using SpreadsheetEditor.Core.Editing;

namespace SpreadsheetEditor.Core.Tests.Editing;

[TestClass]
public sealed class UndoHistoryTests
{
    [TestMethod]
    public void StartsEmptyAndUnmodified()
    {
        var history = new UndoHistory();

        Assert.IsFalse(history.CanUndo);
        Assert.IsFalse(history.CanRedo);
        Assert.IsFalse(history.IsModified);
        Assert.IsNull(history.PopUndo());
        Assert.IsNull(history.PopRedo());
    }

    [TestMethod]
    public void UndoesAndRedoesInStackOrder()
    {
        var history = new UndoHistory();
        var first = new TestEdit();
        var second = new TestEdit();
        history.Record(first);
        history.Record(second);

        Assert.AreSame(second, history.PopUndo());
        Assert.AreSame(first, history.PopUndo());
        Assert.IsFalse(history.CanUndo);
        Assert.AreSame(first, history.PopRedo());
        Assert.AreSame(second, history.PopRedo());
        Assert.IsFalse(history.CanRedo);
    }

    [TestMethod]
    public void ANewEditDropsRedo()
    {
        var history = new UndoHistory();
        history.Record(new TestEdit());
        history.PopUndo();

        history.Record(new TestEdit());

        Assert.IsFalse(history.CanRedo);
        Assert.IsTrue(history.CanUndo);
    }

    [TestMethod]
    public void ModifiedFollowsTheSavedPoint()
    {
        var history = new UndoHistory();
        history.Record(new TestEdit());
        Assert.IsTrue(history.IsModified);

        history.MarkSaved();
        Assert.IsFalse(history.IsModified);

        history.PopUndo();
        Assert.IsTrue(history.IsModified);

        history.PopRedo();
        Assert.IsFalse(history.IsModified);
    }

    [TestMethod]
    public void ASavedPointDroppedWithRedoCannotBeReached()
    {
        var history = new UndoHistory();
        history.Record(new TestEdit());
        history.MarkSaved();
        history.PopUndo();

        history.Record(new TestEdit());
        Assert.IsTrue(history.IsModified);
        history.PopUndo();
        Assert.IsTrue(history.IsModified, "the empty state was never saved");
        history.PopRedo();
        Assert.IsTrue(history.IsModified);
    }

    [TestMethod]
    public void ASavedPointBelowTheRedoStepsStaysReachable()
    {
        var history = new UndoHistory();
        history.Record(new TestEdit());
        history.MarkSaved();
        history.Record(new TestEdit());
        history.PopUndo();

        history.Record(new TestEdit());
        history.PopUndo();

        Assert.IsFalse(history.IsModified);
    }

    [TestMethod]
    public void MergesUntilTheMergeIsBroken()
    {
        var history = new UndoHistory();
        history.Record(new TestEdit(mergeable: true));
        history.Record(new TestEdit(mergeable: true));
        history.BreakMerge();
        history.Record(new TestEdit(mergeable: true));

        Assert.IsNotNull(history.PopUndo());
        Assert.IsNotNull(history.PopUndo());
        Assert.IsFalse(history.CanUndo);
    }

    [TestMethod]
    public void NeverMergesAcrossASave()
    {
        var history = new UndoHistory();
        history.Record(new TestEdit(mergeable: true));
        history.MarkSaved();
        history.Record(new TestEdit(mergeable: true));

        history.PopUndo();

        Assert.IsFalse(history.IsModified);
    }

    [TestMethod]
    public void NeverMergesAfterUndoOrRedo()
    {
        var history = new UndoHistory();
        history.Record(new TestEdit(mergeable: true));
        history.Record(new TestEdit(mergeable: true));
        history.PopUndo();
        history.Record(new TestEdit(mergeable: true));
        history.PopUndo();
        history.PopRedo();
        history.Record(new TestEdit(mergeable: true));

        Assert.IsNotNull(history.PopUndo());
        Assert.IsNotNull(history.PopUndo());
        Assert.IsFalse(history.CanUndo);
    }

    [TestMethod]
    public void ClearForgetsEverything()
    {
        var history = new UndoHistory();
        history.Record(new TestEdit());
        history.Record(new TestEdit());
        history.PopUndo();

        history.Clear();

        Assert.IsFalse(history.CanUndo);
        Assert.IsFalse(history.CanRedo);
        Assert.IsFalse(history.IsModified);
    }

    private sealed class TestEdit(bool mergeable = false) : IEdit
    {
        public IReadOnlyList<WorkbookChange> Apply() => [];

        public IReadOnlyList<WorkbookChange> Revert() => [];

        public bool TryMerge(IEdit next) => mergeable && next is TestEdit { IsMergeable: true };

        private bool IsMergeable => mergeable;
    }
}
