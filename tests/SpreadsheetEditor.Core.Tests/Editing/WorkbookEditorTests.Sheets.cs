using SpreadsheetEditor.Core.Editing;
using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Tests.Editing;

// Sheets (§2.3, §9.31, §9.39, §9.41).
public sealed partial class WorkbookEditorTests
{
    [TestMethod]
    public void AddSheetAppendsANewSheetWithTheDefaults()
    {
        var (editor, first) = Open(cards: 1);
        first.Name = "Sheet1";

        var sheet = editor.AddSheet("Sheet");

        Assert.AreEqual("Sheet1 Sheet2", SheetNamesOf(editor));
        Assert.AreSame(sheet, editor.Workbook.Sheets[1]);
        Assert.AreEqual(1, sheet.Columns.Count);
        Assert.AreEqual("", sheet.Columns[0].Name);
        Assert.AreEqual("_", Describe(sheet));
        Assert.AreEqual(1, sheet.HeaderRow, "headers shown");
        Assert.AreEqual(2, sheet.FirstDataRow);
        Assert.AreEqual(1, sheet.FirstColumn);
        Assert.AreEqual(1, sheet.TableNameColumn);
        Assert.IsNull(sheet.TableName);
        Assert.AreEqual(ImageLayout.Default, sheet.ImageLayout);
        Assert.IsNull(sheet.LockedAspectRatio);
        Assert.IsNull(sheet.Source);
        Assert.IsNull(sheet.Cards[0].SourceFirstRow);
        Assert.IsTrue(editor.IsModified);

        editor.Undo();
        Assert.AreEqual("Sheet1", SheetNamesOf(editor));
        Assert.IsFalse(editor.IsModified);
        editor.Redo();
        Assert.AreSame(sheet, editor.Workbook.Sheets[1]);
    }

    [TestMethod]
    public void NewSheetNamesAvoidSkippedSheetsIgnoringCase()
    {
        var (editor, first) = Open(cards: 1);
        first.Name = "sheet1";
        editor.Workbook.SkippedSheetNames.AddRange(["SHEET2", "Chart"]);

        Assert.AreEqual("Sheet3", editor.AddSheet("Sheet").Name);
    }

    [TestMethod]
    public void RenamesASheetInPlace()
    {
        var (editor, sheet) = Open(cards: 1);
        var columns = sheet.Columns.ToArray();

        Assert.IsTrue(editor.RenameSheet(sheet, "Prices"));
        Assert.AreEqual("Prices", sheet.Name);
        Assert.AreSame(sheet, editor.Workbook.Sheets[0]);
        CollectionAssert.AreEqual(columns, sheet.Columns);
        Assert.IsFalse(editor.RenameSheet(sheet, "Prices"));
        Assert.IsTrue(editor.RenameSheet(sheet, "PRICES"), "a change of case only is allowed");

        editor.Undo();
        Assert.AreEqual("Prices", sheet.Name);
        editor.Undo();
        Assert.AreEqual("S", sheet.Name);
        Assert.IsFalse(editor.IsModified);
        editor.Redo();
        Assert.AreEqual("Prices", sheet.Name);
    }

    [TestMethod]
    public void RenameThrowsWithTheReasonAndChangesNothing()
    {
        var (editor, sheet) = Open(cards: 1);
        var other = editor.AddSheet("Sheet");
        editor.MarkSaved();
        editor.Workbook.SkippedSheetNames.Add("Chart");

        var cases = new (string Name, SheetNameError Error, char? Character)[]
        {
            ("", SheetNameError.Empty, null),
            (new string('x', 32), SheetNameError.TooLong, null),
            ("a/b", SheetNameError.InvalidCharacter, '/'),
            ("'a", SheetNameError.LeadingApostrophe, null),
            ("a'", SheetNameError.TrailingApostrophe, null),
            ("history", SheetNameError.Reserved, null),
            (other.Name.ToUpperInvariant(), SheetNameError.Duplicate, null),
            ("chart", SheetNameError.Duplicate, null),
        };
        foreach (var (name, error, character) in cases)
        {
            var thrown = Assert.ThrowsExactly<SheetNameException>(() => editor.RenameSheet(sheet, name), name);
            Assert.AreEqual(new SheetNameCheck(error, character), thrown.Check, name);
            Assert.AreEqual(new SheetNameCheck(error, character), editor.CheckSheetName(name, sheet), name);
        }

        Assert.AreEqual("S", sheet.Name);
        Assert.IsFalse(editor.IsModified);
        Assert.IsTrue(editor.CheckSheetName("s", sheet).IsValid);
        Assert.AreEqual(SheetNameError.Duplicate, editor.CheckSheetName("s").Error, "a new sheet may not take an existing name");
    }

    [TestMethod]
    public void RemovesASheetButNeverTheLast()
    {
        var (editor, first) = Open(cards: 1);
        var second = editor.AddSheet("Sheet");
        editor.MarkSaved();

        Assert.IsTrue(editor.CanRemoveSheet(first));
        editor.RemoveSheet(first);
        Assert.AreEqual(second.Name, SheetNamesOf(editor));
        Assert.IsTrue(editor.IsModified);
        Assert.IsFalse(editor.CanRemoveSheet(second));
        var error = Assert.ThrowsExactly<InvalidOperationException>(() => editor.RemoveSheet(second));
        Assert.AreEqual("Cannot remove the last sheet.", error.Message);

        editor.Undo();
        Assert.AreSame(first, editor.Workbook.Sheets[0]);
        Assert.AreSame(second, editor.Workbook.Sheets[1]);
        Assert.IsFalse(editor.IsModified, "the modified flag follows the history");
        editor.Redo();
        Assert.AreSame(second, editor.Workbook.Sheets.Single());
    }

    [TestMethod]
    public void MovesASheetAndMarksModified()
    {
        var (editor, first) = Open(cards: 1);
        first.Name = "A";
        editor.AddSheet("B");
        editor.AddSheet("C");
        editor.MarkSaved();

        Assert.IsFalse(editor.MoveSheet(first, 0));
        Assert.IsFalse(editor.IsModified);
        Assert.IsTrue(editor.MoveSheet(first, 2));
        Assert.AreEqual("B1 C1 A", SheetNamesOf(editor));
        Assert.IsTrue(editor.IsModified, "reordering marks the document modified (§9.31)");
        Assert.ThrowsExactly<ArgumentOutOfRangeException>(() => editor.MoveSheet(first, 3));

        editor.Undo();
        Assert.AreEqual("A B1 C1", SheetNamesOf(editor));
        Assert.IsFalse(editor.IsModified);
        editor.Redo();
        Assert.AreEqual("B1 C1 A", SheetNamesOf(editor));
    }

    [TestMethod]
    public void SheetOperationsRejectASheetOfAnotherWorkbook()
    {
        var (editor, _) = Open(cards: 1);
        var (_, other) = Open(cards: 1);

        Assert.ThrowsExactly<ArgumentException>(() => editor.RenameSheet(other, "x"));
        Assert.ThrowsExactly<ArgumentException>(() => editor.RemoveSheet(other));
        Assert.ThrowsExactly<ArgumentException>(() => editor.MoveSheet(other, 0));
        Assert.ThrowsExactly<ArgumentException>(() => editor.CanRemoveSheet(other));
    }

    [TestMethod]
    public void RaisesChangesForSheets()
    {
        var (editor, first) = Open(cards: 1);
        var changes = RecordChanges(editor);

        var added = editor.AddSheet("Sheet");
        editor.RenameSheet(added, "N");
        editor.MoveSheet(added, 0);
        editor.RemoveSheet(first);
        editor.Undo();

        CollectionAssert.AreEqual(
            new WorkbookChange[]
            {
                new SheetInserted(added, 1),
                new SheetRenamed(added),
                new SheetMoved(added, 1, 0),
                new SheetRemoved(first, 1),
                new SheetInserted(first, 1),
            },
            changes);
    }

    private static string SheetNamesOf(WorkbookEditor editor) => string.Join(" ", editor.Workbook.Sheets.Select(s => s.Name));
}
