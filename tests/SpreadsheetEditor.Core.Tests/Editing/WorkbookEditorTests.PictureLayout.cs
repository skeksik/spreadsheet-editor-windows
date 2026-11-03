using SpreadsheetEditor.Core.Editing;
using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Tests.Editing;

// Picture size settings (§2.8, §9.50) and the aspect lock (§9.51).
public sealed partial class WorkbookEditorTests
{
    [TestMethod]
    public void WidthReducesPicturesPerRowToTheNewMaximumInOneStep()
    {
        var (editor, sheet) = Open(cards: 1);

        Assert.IsTrue(editor.SetPictureWidth(sheet, 300));

        Assert.AreEqual(new ImageLayout(300, 180, 5), sheet.ImageLayout, "max(1, floor(1762 / 300)) = 5");
        Assert.IsTrue(editor.IsModified);
        editor.Undo();
        Assert.AreEqual(ImageLayout.Default, sheet.ImageLayout, "one undo restores the width and pictures per row");
        Assert.IsFalse(editor.IsModified);
        editor.Redo();
        Assert.AreEqual(new ImageLayout(300, 180, 5), sheet.ImageLayout);
    }

    [TestMethod]
    public void WidthAndHeightAreClampedToTheirRanges()
    {
        var (editor, sheet) = Open(cards: 1);

        editor.SetPictureWidth(sheet, 2000);
        Assert.AreEqual(new ImageLayout(1000, 180, 1), sheet.ImageLayout);
        editor.SetPictureWidth(sheet, 10);
        Assert.AreEqual(new ImageLayout(50, 180, 1), sheet.ImageLayout, "a narrower width does not raise pictures per row");
        editor.SetPictureHeight(sheet, 600);
        Assert.AreEqual(new ImageLayout(50, 500, 1), sheet.ImageLayout);
        editor.SetPictureHeight(sheet, -1);
        Assert.AreEqual(new ImageLayout(50, 50, 1), sheet.ImageLayout);
    }

    [TestMethod]
    public void TheValueItAlreadyHasDoesNothing()
    {
        var (editor, sheet) = Open(cards: 1);
        sheet.ImageLayout = new ImageLayout(1000, 500, 1);

        Assert.IsFalse(editor.SetPictureWidth(sheet, 1000));
        Assert.IsFalse(editor.SetPictureWidth(sheet, 5000), "clamped to the current width");
        Assert.IsFalse(editor.SetPictureHeight(sheet, 500));
        Assert.IsFalse(editor.SetPicturesPerRow(sheet, 1));
        Assert.IsFalse(editor.SetPicturesPerRow(sheet, 3), "clamped to the maximum, 1");

        Assert.IsFalse(editor.CanUndo);
        Assert.IsFalse(editor.IsModified);
    }

    [TestMethod]
    public void PicturesPerRowIsClampedToTheMaximumForTheWidth()
    {
        var (editor, sheet) = Open(cards: 1);

        Assert.IsTrue(editor.SetPicturesPerRow(sheet, 20));
        Assert.AreEqual(12, sheet.ImageLayout.ImagesPerRow, "floor(1762 / 136) = 12");
        Assert.IsTrue(editor.SetPicturesPerRow(sheet, 0));
        Assert.AreEqual(1, sheet.ImageLayout.ImagesPerRow);

        editor.Undo();
        Assert.AreEqual(12, sheet.ImageLayout.ImagesPerRow);
        editor.Undo();
        Assert.AreEqual(ImageLayout.Default, sheet.ImageLayout);
    }

    [TestMethod]
    public void TheAspectLockKeepsTheRatioCapturedWhenTurnedOn()
    {
        var (editor, sheet) = Open(cards: 1);
        sheet.ImageLayout = new ImageLayout(100, 200, 8);

        Assert.IsTrue(editor.SetAspectLock(sheet, true));
        Assert.AreEqual(0.5, sheet.LockedAspectRatio);
        Assert.IsFalse(editor.CanUndo, "the lock is not an undoable step");
        Assert.IsFalse(editor.IsModified, "nor does it mark the document modified");

        editor.SetPictureWidth(sheet, 150);
        Assert.AreEqual(new ImageLayout(150, 300, 8), sheet.ImageLayout);
        editor.SetPictureHeight(sheet, 101);
        Assert.AreEqual(new ImageLayout(51, 101, 8), sheet.ImageLayout, "round(101 × 0.5) = 51, half away from zero");

        editor.SetPictureWidth(sheet, 400);
        Assert.AreEqual(new ImageLayout(400, 500, 4), sheet.ImageLayout, "height 800 clamped to 500; pictures per row reduced");
        Assert.AreEqual(0.5, sheet.LockedAspectRatio, "the ratio is not recaptured after clamping");
        editor.SetPictureHeight(sheet, 50);
        Assert.AreEqual(new ImageLayout(50, 50, 4), sheet.ImageLayout, "width 25 clamped to 50");

        editor.Undo();
        Assert.AreEqual(new ImageLayout(400, 500, 4), sheet.ImageLayout);
        editor.Undo();
        Assert.AreEqual(new ImageLayout(51, 101, 8), sheet.ImageLayout, "one undo restores both sides and pictures per row");
    }

    [TestMethod]
    public void HeightWithTheLockOnReducesPicturesPerRow()
    {
        var (editor, sheet) = Open(cards: 1);
        editor.SetAspectLock(sheet, true);

        editor.SetPictureHeight(sheet, 500);

        Assert.AreEqual(new ImageLayout(378, 500, 4), sheet.ImageLayout, "width round(500 × 136 / 180) = 378, floor(1762 / 378) = 4");
    }

    [TestMethod]
    public void TurningTheLockOnAgainKeepsItsRatioAndOffStopsTheCoupling()
    {
        var (editor, sheet) = Open(cards: 1);
        sheet.ImageLayout = new ImageLayout(100, 100, 8);
        editor.SetAspectLock(sheet, true);
        editor.SetPictureHeight(sheet, 200);

        Assert.IsFalse(editor.SetAspectLock(sheet, true));
        Assert.AreEqual(1.0, sheet.LockedAspectRatio);

        Assert.IsTrue(editor.SetAspectLock(sheet, false));
        Assert.IsNull(sheet.LockedAspectRatio);
        Assert.IsFalse(editor.SetAspectLock(sheet, false));
        editor.SetPictureWidth(sheet, 300);
        Assert.AreEqual(new ImageLayout(300, 200, 5), sheet.ImageLayout);
    }

    [TestMethod]
    public void RaisesChangesForThePictureLayoutAndTheLock()
    {
        var (editor, sheet) = Open(cards: 1);
        var changes = RecordChanges(editor);
        var states = 0;
        editor.StateChanged += (_, _) => states++;

        editor.SetAspectLock(sheet, true);
        editor.SetPictureWidth(sheet, 200);
        editor.SetPictureHeight(sheet, 100);
        editor.SetPicturesPerRow(sheet, 2);
        editor.Undo();
        editor.SetAspectLock(sheet, false);

        CollectionAssert.AreEqual(
            new WorkbookChange[]
            {
                new AspectLockChanged(sheet),
                new PictureLayoutChanged(sheet),
                new PictureLayoutChanged(sheet),
                new PictureLayoutChanged(sheet),
                new PictureLayoutChanged(sheet),
                new AspectLockChanged(sheet),
            },
            changes);
        Assert.AreEqual(4, states, "the lock does not change the undo state");
    }
}
