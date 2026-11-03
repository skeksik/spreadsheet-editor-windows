using SpreadsheetEditor.Core.Editing;
using SpreadsheetEditor.Core.Model;
using SpreadsheetEditor.Core.Tests.Imaging;

namespace SpreadsheetEditor.Core.Tests.Editing;

// Values (§2.6), the table name (§2.7) and the typing merge shared with entry names.
public sealed partial class WorkbookEditorTests
{
    private static readonly RgbColor Red = new(0xFF, 0, 0);
    private static readonly RgbColor Blue = new(0, 0, 0xFF);

    [TestMethod]
    public void SetsAnEntrysTextAndKeepsTheValueAsRead()
    {
        var (editor, sheet) = Open(cards: 2);
        var entry = sheet.Cards[1].Entries[0];
        var original = new CellValue(CellValueKind.Number, "r1a", Number: 1);
        entry.Original = original;

        Assert.IsTrue(editor.SetEntryText(sheet, sheet.Cards[1], sheet.Columns[0], "x"));

        Assert.AreEqual("r0a r0b|x r1b", Describe(sheet));
        Assert.AreSame(original, entry.Original);
        Assert.IsTrue(editor.IsModified);
        AssertUndoRedo(editor, sheet, "r0a r0b|r1a r1b", "r0a r0b|x r1b");
    }

    [TestMethod]
    public void UnchangedTextRecordsNothing()
    {
        var (editor, sheet) = Open(cards: 1);

        Assert.IsFalse(editor.SetEntryText(sheet, sheet.Cards[0], sheet.Columns[0], "r0a"));

        Assert.IsFalse(editor.CanUndo);
        Assert.IsFalse(editor.IsModified);
    }

    [TestMethod]
    public void TypingInOneEntryIsOneStepPerEntry()
    {
        var (editor, sheet) = Open(cards: 1);
        var card = sheet.Cards[0];
        var (a, b) = (sheet.Columns[0], sheet.Columns[1]);

        editor.SetEntryText(sheet, card, a, "1");
        editor.SetEntryText(sheet, card, a, "11");
        editor.SetEntryText(sheet, card, b, "2");
        editor.SetEntryText(sheet, card, b, "22");
        editor.SetEntryText(sheet, card, b, "222");
        editor.EndTyping();
        editor.SetEntryText(sheet, card, b, "3");

        Assert.AreEqual("11 3", Describe(sheet));
        editor.Undo();
        Assert.AreEqual("11 222", Describe(sheet));
        editor.Undo();
        Assert.AreEqual("11 r0b", Describe(sheet));
        editor.Undo();
        Assert.AreEqual("r0a r0b", Describe(sheet));
        Assert.IsFalse(editor.CanUndo);
    }

    [TestMethod]
    public void TypingInTheSameEntryOfAnotherCardIsAnotherStep()
    {
        var (editor, sheet) = Open(cards: 2);
        var column = sheet.Columns[0];

        editor.SetEntryText(sheet, sheet.Cards[0], column, "x");
        editor.SetEntryText(sheet, sheet.Cards[1], column, "y");

        editor.Undo();
        Assert.AreEqual("x r0b|r1a r1b", Describe(sheet));
    }

    [TestMethod]
    public void TypingBackToTheOldTextDropsTheStep()
    {
        var (editor, sheet) = Open(cards: 1);
        editor.AddCard(sheet);
        editor.MarkSaved();
        var (card, column) = (sheet.Cards[0], sheet.Columns[0]);

        editor.SetEntryText(sheet, card, column, "r0ax");
        Assert.IsTrue(editor.IsModified);
        editor.SetEntryText(sheet, card, column, "r0a");

        Assert.IsFalse(editor.IsModified);
        Assert.IsFalse(editor.CanRedo);
        editor.Undo();
        Assert.AreEqual("r0a r0b", Describe(sheet), "the undo reverts the step before, not the dropped one");

        editor.Redo();
        editor.SetEntryText(sheet, card, column, "a");
        editor.SetEntryText(sheet, card, column, "r0a");
        editor.SetEntryText(sheet, card, column, "b");
        editor.Undo();
        Assert.AreEqual("r0a r0b|_ _", Describe(sheet), "typing on after a dropped step starts a new step");
        Assert.IsFalse(editor.IsModified);
    }

    [TestMethod]
    public void RenamingBackToTheOldNameDropsTheStep()
    {
        var (editor, sheet) = Open(cards: 1);
        var column = sheet.Columns[0];

        editor.RenameEntry(sheet, column, "ax");
        editor.RenameEntry(sheet, column, "a");

        Assert.IsFalse(editor.IsModified);
        Assert.IsFalse(editor.CanUndo);
        Assert.AreEqual("a b", Names(sheet));
    }

    [TestMethod]
    public void SetsAndClearsTheFill()
    {
        var (editor, sheet) = Open(cards: 1);
        var (card, column) = (sheet.Cards[0], sheet.Columns[0]);
        var entry = card.Entries[0];
        var font = new FontStyle { Bold = true };
        entry.Style = new CellStyle { TextColor = Blue, Font = font };

        Assert.IsTrue(editor.SetFill(sheet, card, column, Red));
        Assert.AreEqual(new CellStyle { Fill = Red, TextColor = Blue, Font = font }, entry.Style);
        Assert.IsFalse(editor.SetFill(sheet, card, column, Red));

        Assert.IsTrue(editor.SetFill(sheet, card, column, null));
        Assert.AreEqual(new CellStyle { NoFill = true, TextColor = Blue, Font = font }, entry.Style);
        Assert.IsFalse(editor.SetFill(sheet, card, column, null));

        editor.Undo();
        Assert.AreEqual(Red, entry.Style!.Fill);
        editor.Undo();
        Assert.AreEqual(new CellStyle { TextColor = Blue, Font = font }, entry.Style);
        editor.Redo();
        editor.Redo();
        Assert.IsTrue(entry.Style!.NoFill);
        Assert.IsNull(entry.Style.Fill);
    }

    [TestMethod]
    public void FillOnAnEntryWithoutAStyleStartsFromAnEmptyStyle()
    {
        var (editor, sheet) = Open(cards: 1);
        var entry = sheet.Cards[0].Entries[1];

        editor.SetFill(sheet, sheet.Cards[0], sheet.Columns[1], Red);

        Assert.AreEqual(new CellStyle { Fill = Red }, entry.Style);
        editor.Undo();
        Assert.IsNull(entry.Style);
    }

    [TestMethod]
    public void SetsTheTextColourOrAutomatic()
    {
        var (editor, sheet) = Open(cards: 1);
        var (card, column) = (sheet.Cards[0], sheet.Columns[0]);
        var entry = card.Entries[0];
        entry.Style = new CellStyle { Fill = Red, AutomaticTextColor = true };

        Assert.IsFalse(editor.SetTextColor(sheet, card, column, null), "already automatic");
        Assert.IsTrue(editor.SetTextColor(sheet, card, column, Blue));
        Assert.AreEqual(new CellStyle { Fill = Red, TextColor = Blue }, entry.Style);
        Assert.IsTrue(editor.SetTextColor(sheet, card, column, null));
        Assert.AreEqual(new CellStyle { Fill = Red, AutomaticTextColor = true }, entry.Style);

        editor.Undo();
        Assert.AreEqual(Blue, entry.Style!.TextColor);
        editor.Undo();
        Assert.AreEqual(new CellStyle { Fill = Red, AutomaticTextColor = true }, entry.Style);
        Assert.IsFalse(editor.IsModified);
    }

    [TestMethod]
    public void AddsPicturesFromBytesAndMarksTheColumn()
    {
        var (editor, sheet) = Open(cards: 1);
        var (card, column) = (sheet.Cards[0], sheet.Columns[1]);
        var entry = card.Entries[1];
        var existing = TestImages.Tall();
        entry.Images.Add(existing);
        var png = PictureBytes.Png(30, 20);
        byte[] junk = [1, 2, 3];

        var added = editor.AddPictures(sheet, card, column, [png, junk]);

        Assert.AreEqual(2, added.Count);
        CollectionAssert.AreEqual(new[] { existing, added[0], added[1] }, entry.Images);
        Assert.AreSame(png, added[0].Data, "the bytes are kept unchanged");
        Assert.AreEqual(ImageFormat.Png, added[0].Format);
        Assert.AreEqual(30, added[0].PixelWidth);
        Assert.IsFalse(added[1].IsDecodable, "unreadable bytes are kept as an undecodable picture");
        Assert.IsTrue(column.IsImageColumn);

        editor.Undo();
        CollectionAssert.AreEqual(new[] { existing }, entry.Images);
        Assert.IsFalse(column.IsImageColumn);
        editor.Redo();
        CollectionAssert.AreEqual(new[] { existing, added[0], added[1] }, entry.Images);
        Assert.IsTrue(column.IsImageColumn);
    }

    [TestMethod]
    public void AddingNoPicturesRecordsNothing()
    {
        var (editor, sheet) = Open(cards: 1);

        Assert.AreEqual(0, editor.AddPictures(sheet, sheet.Cards[0], sheet.Columns[0], []).Count);

        Assert.IsFalse(editor.CanUndo);
        Assert.IsFalse(sheet.Columns[0].IsImageColumn);
    }

    [TestMethod]
    public void RemovesOnePictureByIdentityOrAll()
    {
        var (editor, sheet) = Open(cards: 1);
        var (card, column) = (sheet.Cards[0], sheet.Columns[0]);
        column.IsImageColumn = true;
        var entry = card.Entries[0];
        var (p0, p1, p2) = (TestImages.Tall(), TestImages.Wide(), TestImages.Square());
        entry.Images.AddRange([p0, p1, p2]);

        editor.RemovePicture(sheet, card, column, p1);
        CollectionAssert.AreEqual(new[] { p0, p2 }, entry.Images);
        Assert.ThrowsExactly<ArgumentException>(() => editor.RemovePicture(sheet, card, column, p1));

        Assert.IsTrue(editor.RemoveAllPictures(sheet, card, column));
        Assert.AreEqual(0, entry.Images.Count);
        Assert.IsTrue(column.IsImageColumn, "the column stays a picture column");
        Assert.IsFalse(editor.RemoveAllPictures(sheet, card, column));

        editor.Undo();
        CollectionAssert.AreEqual(new[] { p0, p2 }, entry.Images);
        editor.Undo();
        CollectionAssert.AreEqual(new[] { p0, p1, p2 }, entry.Images);
        editor.Redo();
        CollectionAssert.AreEqual(new[] { p0, p2 }, entry.Images);
    }

    [TestMethod]
    public void MovesAPictureWithinItsEntry()
    {
        var (editor, sheet) = Open(cards: 1);
        var (card, column) = (sheet.Cards[0], sheet.Columns[0]);
        var entry = card.Entries[0];
        var (p0, p1, p2) = (TestImages.Tall(), TestImages.Wide(), TestImages.Square());
        entry.Images.AddRange([p0, p1, p2]);

        Assert.IsFalse(editor.MovePicture(sheet, card, column, p1, 1));
        Assert.IsFalse(editor.CanUndo);
        Assert.IsTrue(editor.MovePicture(sheet, card, column, p0, 2));
        CollectionAssert.AreEqual(new[] { p1, p2, p0 }, entry.Images);
        Assert.ThrowsExactly<ArgumentOutOfRangeException>(() => editor.MovePicture(sheet, card, column, p0, 3));
        Assert.ThrowsExactly<ArgumentException>(() => editor.MovePicture(sheet, card, column, TestImages.Tall(), 0));

        editor.Undo();
        CollectionAssert.AreEqual(new[] { p0, p1, p2 }, entry.Images);
        editor.Redo();
        CollectionAssert.AreEqual(new[] { p1, p2, p0 }, entry.Images);
    }

    [TestMethod]
    public void ValueOperationsRejectACardOrEntryOfAnotherSheet()
    {
        var (editor, sheet) = Open(cards: 1);
        var (_, other) = Open(cards: 1);

        Assert.ThrowsExactly<ArgumentException>(() => editor.SetEntryText(sheet, other.Cards[0], sheet.Columns[0], "x"));
        Assert.ThrowsExactly<ArgumentException>(() => editor.SetFill(sheet, sheet.Cards[0], other.Columns[0], Red));
        Assert.ThrowsExactly<ArgumentException>(() => editor.AddPictures(other, other.Cards[0], other.Columns[0], []));
    }

    [TestMethod]
    public void SetsAndClearsTheTableName()
    {
        var (editor, sheet) = Open(cards: 1);
        sheet.TableName = "Old";

        Assert.IsFalse(editor.SetTableName(sheet, "Old"));
        Assert.IsTrue(editor.SetTableName(sheet, ""));
        Assert.IsNull(sheet.TableName, "empty text clears the name");
        Assert.IsFalse(editor.SetTableName(sheet, null));
        editor.EndTyping();
        Assert.IsTrue(editor.SetTableName(sheet, "N"));
        editor.SetTableName(sheet, "Ne");
        editor.SetTableName(sheet, "New");

        Assert.AreEqual("New", sheet.TableName);
        editor.Undo();
        Assert.IsNull(sheet.TableName, "typing the new name was one step");
        editor.Undo();
        Assert.AreEqual("Old", sheet.TableName);
        Assert.IsFalse(editor.IsModified);
        editor.Redo();
        editor.Redo();
        Assert.AreEqual("New", sheet.TableName);
    }

    [TestMethod]
    public void TypingTheTableNameBackDropsTheStep()
    {
        var (editor, sheet) = Open(cards: 1);

        editor.SetTableName(sheet, "T");
        editor.SetTableName(sheet, "");

        Assert.IsNull(sheet.TableName);
        Assert.IsFalse(editor.CanUndo);
        Assert.IsFalse(editor.IsModified);
    }

    [TestMethod]
    public void RaisesChangesForValuesAndTheTableName()
    {
        var (editor, sheet) = Open(cards: 1);
        var (card, column) = (sheet.Cards[0], sheet.Columns[0]);
        var changes = RecordChanges(editor);

        editor.SetEntryText(sheet, card, column, "x");
        editor.Undo();
        editor.SetFill(sheet, card, column, Red);
        editor.SetTextColor(sheet, card, column, Blue);
        var added = editor.AddPictures(sheet, card, column, [PictureBytes.Png(1, 1), PictureBytes.Png(2, 2)]);
        editor.MovePicture(sheet, card, column, added[0], 1);
        editor.RemovePicture(sheet, card, column, added[0]);
        editor.RemoveAllPictures(sheet, card, column);
        editor.SetTableName(sheet, "T");

        CollectionAssert.AreEqual(
            new WorkbookChange[]
            {
                new EntryTextChanged(sheet, card, column),
                new EntryTextChanged(sheet, card, column),
                new EntryStyleChanged(sheet, card, column),
                new EntryStyleChanged(sheet, card, column),
                new EntryPicturesChanged(sheet, card, column),
                new EntryPicturesChanged(sheet, card, column),
                new EntryPicturesChanged(sheet, card, column),
                new EntryPicturesChanged(sheet, card, column),
                new TableNameChanged(sheet),
            },
            changes);
    }

    private static List<WorkbookChange> RecordChanges(WorkbookEditor editor)
    {
        var changes = new List<WorkbookChange>();
        editor.Changed += (_, e) => changes.AddRange(e.Changes);
        return changes;
    }
}
