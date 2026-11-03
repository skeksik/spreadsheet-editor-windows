using System.Globalization;
using SpreadsheetEditor.Core.Editing;
using SpreadsheetEditor.Core.Model;
using SpreadsheetEditor.Core.Tests.Imaging;
using SpreadsheetEditor.Core.Tests.Xlsx;
using SpreadsheetEditor.Core.Xlsx;

namespace SpreadsheetEditor.Core.Tests.Editing;

// Sheets (§2.3, §9.30-9.32, §9.39, §9.41).
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
    public void DuplicateInsertsADeepCopyAfterTheSource()
    {
        var (editor, sheet) = Open(cards: 2);
        var last = editor.AddSheet("Sheet");
        sheet.TableName = "Edited name";
        sheet.TableNameColumn = 3;
        sheet.FirstColumn = 2;
        sheet.HeaderRow = 4;
        sheet.FirstDataRow = 5;
        sheet.ImageLayout = new ImageLayout(200, 100, 3);
        sheet.Columns[1].IsImageColumn = true;
        sheet.Columns[1].SourceIndex = 0;
        sheet.Cards[0].SourceRowSpan = 3;
        sheet.Cards[0].SourceFirstRow = 5;
        var entry = sheet.Cards[0].Entries[1];
        entry.Original = new CellValue(CellValueKind.Number, "r0b", Number: 7);
        entry.Style = new CellStyle { Fill = Red };
        entry.Images.Add(TestImages.Wide());
        editor.SetAspectLock(sheet, true);
        editor.MarkSaved();

        var copy = editor.DuplicateSheet(sheet);

        Assert.AreEqual("S S (Copy) Sheet1", SheetNamesOf(editor));
        Assert.AreSame(copy, editor.Workbook.Sheets[1]);
        Assert.AreSame(last, editor.Workbook.Sheets[2]);
        Assert.AreEqual(Describe(sheet), Describe(copy));
        Assert.AreEqual(Names(sheet), Names(copy));
        Assert.AreEqual(Spans(sheet), Spans(copy), "read cards keep their source rows: the writer copies the whole worksheet");
        Assert.AreEqual("Edited name", copy.TableName, "the current table name (§9.32)");
        Assert.AreEqual((3, 2, 4, 5), (copy.TableNameColumn, copy.FirstColumn, copy.HeaderRow, copy.FirstDataRow));
        Assert.AreEqual(sheet.ImageLayout, copy.ImageLayout);
        Assert.IsNull(copy.LockedAspectRatio, "the aspect lock is not copied (§2.3)");
        Assert.IsNull(copy.Source, "a sheet not read from a file has no source to copy");
        for (var c = 0; c < sheet.Columns.Count; c++)
        {
            Assert.AreNotSame(sheet.Columns[c], copy.Columns[c]);
            Assert.AreEqual(sheet.Columns[c].IsImageColumn, copy.Columns[c].IsImageColumn);
            Assert.AreEqual(sheet.Columns[c].SourceIndex, copy.Columns[c].SourceIndex);
        }

        var copied = copy.Cards[0].Entries[1];
        Assert.AreNotSame(sheet.Cards[0], copy.Cards[0]);
        Assert.AreNotSame(entry, copied);
        Assert.AreSame(entry.Original, copied.Original);
        Assert.AreEqual(entry.Style, copied.Style);
        Assert.AreNotSame(entry.Images, copied.Images);
        CollectionAssert.AreEqual(entry.Images, copied.Images);
        Assert.IsTrue(editor.IsModified);

        editor.SetEntryText(copy, copy.Cards[0], copy.Columns[0], "changed");
        editor.RenameEntry(copy, copy.Columns[0], "renamed");
        Assert.AreEqual("r0a", sheet.Cards[0].Entries[0].Text, "editing the copy leaves the source alone");
        Assert.AreEqual("a", sheet.Columns[0].Name);

        editor.Undo();
        editor.Undo();
        editor.Undo();
        Assert.AreEqual("S Sheet1", SheetNamesOf(editor));
        Assert.IsFalse(editor.IsModified);
        editor.Redo();
        Assert.AreSame(copy, editor.Workbook.Sheets[1]);
    }

    [TestMethod]
    public void DuplicateNamesAvoidUsedAndSkippedNames()
    {
        var (editor, sheet) = Open(cards: 1);
        editor.Workbook.SkippedSheetNames.Add("s (copy)");

        Assert.AreEqual("S (Copy 1)", editor.DuplicateSheet(sheet).Name);
        Assert.AreEqual("S (Copy 2)", editor.DuplicateSheet(sheet).Name);
        sheet.Name = new string('x', 31);
        Assert.AreEqual(new string('x', 24) + " (Copy)", editor.DuplicateSheet(sheet).Name);
    }

    [TestMethod]
    public void DuplicateOfAReadSheetIsWrittenAsACopyOfItsWorksheet()
    {
        var book = new TestWorkbook();
        book.AddSheet("Plants").Row("A1", "Garden").Row("A2", "Name", "Colour").Row("A3", "Fern", "Green").Row("A4", "Rose", "Red")
            .Merge("A4:A5");
        using var input = new MemoryStream(book.ToBytes());
        var workbook = XlsxReader.Read(input, CultureInfo.InvariantCulture);
        var editor = new WorkbookEditor(workbook);
        var source = workbook.Sheets[0];

        var copy = editor.DuplicateSheet(source);
        Assert.AreEqual("Garden", copy.TableName);
        editor.SetEntryText(copy, copy.Cards[0], copy.Columns[1], "Blue");
        editor.SetTableName(copy, "Copied plants");

        Assert.IsNotNull(copy.Source);
        Assert.IsTrue(copy.Source.IsCopy);
        Assert.AreEqual(source.Source! with { IsCopy = true }, copy.Source);
        Assert.IsFalse(source.Source.IsCopy);
        Assert.AreEqual(Spans(source), Spans(copy));
        using var output = new MemoryStream();
        XlsxWriter.Write(workbook, output, CultureInfo.InvariantCulture);
        output.Position = 0;
        var back = XlsxReader.Read(output, CultureInfo.InvariantCulture);

        Assert.AreEqual("Plants|Plants (Copy)", string.Join("|", back.Sheets.Select(s => s.Name)));
        Assert.AreEqual("Fern Green|Rose Red", Describe(back.Sheets[0]));
        Assert.AreEqual("Name Colour", Names(back.Sheets[0]));
        Assert.AreEqual("Garden", back.Sheets[0].TableName);
        Assert.AreEqual("Fern Blue|Rose Red", Describe(back.Sheets[1]));
        Assert.AreEqual("Name Colour", Names(back.Sheets[1]));
        Assert.AreEqual("Copied plants", back.Sheets[1].TableName);
        Assert.AreEqual(2, back.Sheets[1].Cards[1].SourceRowSpan, "the copy keeps the merged card");
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
