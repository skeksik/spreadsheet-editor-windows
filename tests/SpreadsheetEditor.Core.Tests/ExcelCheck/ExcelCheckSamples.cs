using System.Globalization;
using SpreadsheetEditor.Core.Editing;
using SpreadsheetEditor.Core.Model;
using SpreadsheetEditor.Core.Tests.Xlsx;
using SpreadsheetEditor.Core.Xlsx;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetEditor.Core.Tests.ExcelCheck;

/// <summary>
/// The sample workbooks of the Excel check (docs/notes/excel-check.md). Each builds a source workbook in code, reads it with
/// <see cref="XlsxReader"/>, edits the model with <see cref="WorkbookEditor"/> as the app would, and returns the edited workbook for
/// <see cref="XlsxWriter"/> to save. Every source workbook ends with a sheet "Check" that says in a few lines what the file tests; it is
/// also a sheet the edits do not touch. Every value in them is made up.
/// </summary>
internal static class ExcelCheckSamples
{
    /// <summary>The UI language the edits are typed in; Excel shows the values in the PC's own formats.</summary>
    public static readonly CultureInfo Culture = CultureInfo.GetCultureInfo("en-US");

    public const int NumberFormatText = 49;
    public const int NumberFormatShortDate = 14;

    /// <summary>01: values typed into cards (numbers, dates, times, booleans, formulas, Text cells) and values left as read.</summary>
    public static Workbook Values()
    {
        var book = new TestWorkbook();
        var date = book.AddNumberFormat(NumberFormatShortDate);
        var text = book.AddNumberFormat(NumberFormatText);
        book.AddSheet("Values")
            .Row("A1", "Item", "Value")
            .Row("A2", "Decimal typed", "x")
            .Row("A3", "Whole number typed", "x")
            .Row("A4", "Date typed", "x")
            .Row("A5", "Date and time typed", "x")
            .Row("A6", "Time typed", "x")
            .Row("A7", "TRUE typed", "x")
            .Row("A8", "false typed", "x")
            .Row("A9", "Formula typed", "x")
            .Row("A10", "Formula kept").Formula("B10", "B3+1", 11)
            .Row("A11", "Digits typed in a Text cell").Set("B11", "x", text)
            .Row("A12", "Leading zeros typed", "x")
            .Row("A13", "Words with a number typed", "x")
            .Row("A14", "Date kept").Set("B14", 45000, date)
            .Row("A15", "Number kept", 1234.5)
            .Row("A16", "Boolean kept", true)
            .Row("A17", "Text kept", "plain words")
            .Row("A18", "Cleared", "old text");
        AddCheckSheet(book, "01 values: typed values become numbers, dates, booleans and formulas; untouched values keep their type.");

        var (editor, sheet) = Open(book, "Values");
        string[] typed = ["3.5", "42", "3/15/2024", "3/15/2024 2:30 PM", "2:30 PM", "TRUE", "false", "=B2*2"];
        for (var i = 0; i < typed.Length; i++)
        {
            Type(editor, sheet, i, "Value", typed[i]);
        }

        Type(editor, sheet, 9, "Value", "007");
        Type(editor, sheet, 10, "Value", "007");
        Type(editor, sheet, 11, "Value", "12 apples");
        Type(editor, sheet, 16, "Value", "");
        return editor.Workbook;
    }

    /// <summary>02: fills and text colours set, removed and set back to automatic; fonts and alignment kept; multi-line text.</summary>
    public static Workbook Styles()
    {
        var book = new TestWorkbook();
        var green = book.AddStyle(fill: TestWorkbook.Solid(new X.ForegroundColor { Rgb = "FF00B050" }));
        var redText = book.AddStyle(font: Calibri(color: "FFFF0000"));
        var boldItalic = book.AddStyle(font: Calibri(bold: true, italic: true));
        var orangeBold = book.AddStyle(
            font: Calibri(bold: true),
            fill: TestWorkbook.Solid(new X.ForegroundColor { Rgb = "FFFFC000" }));
        var centred = book.AddStyle(alignment: new X.Alignment { Horizontal = X.HorizontalAlignmentValues.Center });
        book.AddSheet("Styles")
            .Row("A1", "Item", "Value")
            .Row("A2", "Fill set to yellow", "yellow fill")
            .Row("A3", "Green fill removed").Set("B3", "no fill now", green)
            .Row("A4", "Text colour set to blue", "blue text")
            .Row("A5", "Red text set to Automatic").Set("B5", "automatic text", redText)
            .Row("A6", "Bold italic, text edited").Set("B6", "old", boldItalic)
            .Row("A7", "Orange bold, untouched").Set("B7", "untouched", orangeBold)
            .Row("A8", "Multi-line text typed", "x")
            .Row("A9", "Centred, text edited").Set("B9", "old", centred);
        AddCheckSheet(book, "02 styles: fills and text colours set, removed and automatic; fonts and alignment kept; multi-line text.");

        var (editor, sheet) = Open(book, "Styles");
        var value = Column(sheet, "Value");
        editor.SetFill(sheet, sheet.Cards[0], value, new RgbColor(0xFF, 0xFF, 0x00));
        editor.SetFill(sheet, sheet.Cards[1], value, null);
        editor.SetTextColor(sheet, sheet.Cards[2], value, new RgbColor(0x00, 0x46, 0xC8));
        editor.SetTextColor(sheet, sheet.Cards[3], value, null);
        Type(editor, sheet, 4, "Value", "still bold italic");
        Type(editor, sheet, 6, "Value", "line one\nline two\nline three");
        Type(editor, sheet, 7, "Value", "still centred");
        return editor.Workbook;
    }

    /// <summary>
    /// 03: cards added, removed and reordered on a sheet with content below the table (totals, a merge, a styled cell, a defined name, a
    /// print area) and formulas on another sheet pointing into the table and below it.
    /// </summary>
    public static Workbook Cards()
    {
        var book = CardsSource("03 cards: Banana removed, Elder moved to the top, Fig added at the end, Grape inserted above Cherry.");
        var (editor, sheet) = Open(book, "Cards");
        var banana = sheet.Cards[1];
        var cherry = sheet.Cards[2];
        var elder = sheet.Cards[4];
        editor.RemoveCard(sheet, banana);
        editor.MoveCard(sheet, elder, 0);
        var fig = editor.AddCard(sheet);
        Type(editor, sheet, fig, "Name", "Fig");
        Type(editor, sheet, fig, "Qty", "6");
        var grape = editor.InsertCardAbove(sheet, cherry);
        Type(editor, sheet, grape, "Name", "Grape");
        Type(editor, sheet, grape, "Qty", "7");
        return editor.Workbook;
    }

    /// <summary>
    /// The source workbook of sample 03, whose last sheet "Check" says <paramref name="checkLine"/>; with <paramref name="picture"/>, a
    /// picture anchored at G2 too.
    /// </summary>
    private static TestWorkbook CardsSource(string checkLine, bool picture = false)
    {
        var book = new TestWorkbook();
        var yellow = book.AddStyle(fill: TestWorkbook.Solid(new X.ForegroundColor { Rgb = "FFFFFF00" }));
        var cards = book.AddSheet("Cards");
        if (picture)
        {
            cards.Picture("G2", CheckPictures.Png(40, 60, 2, 3, 0), PictureAnchor.OneCell);
        }

        cards
            .Row("A1", "Name", "Qty")
            .Row("A2", "Apple", 1)
            .Row("A3", "Banana", 2)
            .Row("A4", "Cherry", 3)
            .Row("A5", "Damson", 4)
            .Row("A6", "Elder", 5)
            .Set("A20", "Total").Formula("B20", "SUM(B2:B6)", 15)
            .Set("A21", "Count").Formula("B21", "COUNTA(A2:A6)", 5)
            .Set("A22", "Cherry's qty").Formula("B22", "B4", 3)
            .Set("A23", "Banana's qty").Formula("B23", "B3", 2)
            .Set("D20", "merged note below the table").Merge("D20:E20")
            .Set("A24", "yellow cell below the table", yellow);
        book.AddSheet("Report")
            .Row("A1", "Item", "Value")
            .Row("A2", "Total").Formula("B2", "Cards!B20", 15)
            .Row("A3", "Row 2 qty").Formula("B3", "Cards!B2", 1)
            .Row("A4", "Named total").Formula("B4", "CardTotal", 15);
        book.AddDefinedName("CardTotal", "Cards!$B$20");
        book.AddDefinedName("_xlnm.Print_Area", "Cards!$A$1:$E$24", 0);
        AddCheckSheet(book, checkLine);
        return book;
    }

    /// <summary>
    /// 00: the control: a source workbook as the test builders make it (sample 03's, with a picture added), not read or written by the
    /// app. If Excel offers to repair this file too, the builders are at fault, not the writer.
    /// </summary>
    public static byte[] Control()
    {
        var book = CardsSource("00 control: made by the test builders alone; the app did not read or write this file.", picture: true);
        var bytes = book.ToBytes();
        var errors = PackageValidation.Errors(bytes);
        return errors.Count == 0 ? bytes : throw new InvalidOperationException("The control is not valid:\n" + string.Join("\n", errors));
    }

    /// <summary>04: a table name given to a table read without one: a row is inserted above the header and everything below shifts.</summary>
    public static Workbook TableName()
    {
        var book = new TestWorkbook();
        book.AddSheet("Named")
            .Row("A1", "Name", "Price")
            .Row("A2", "Pen", 2)
            .Row("A3", "Ink", 5)
            .Row("A4", "Pad", 3)
            .Set("A15", "Total").Formula("B15", "SUM(B2:B4)", 10);
        book.AddSheet("Report").Row("A1", "Item", "Value").Row("A2", "Total").Formula("B2", "Named!B15", 10);
        AddCheckSheet(book, "04 table name: the name Stationery is given to the table; a row is inserted above the header.");

        var (editor, sheet) = Open(book, "Named");
        editor.SetTableName(sheet, "Stationery");
        return editor.Workbook;
    }

    /// <summary>05: a shared formula over the table's rows, split by a removed card and an edited cell.</summary>
    public static Workbook SharedFormulas()
    {
        var book = new TestWorkbook();
        book.AddSheet("Shared")
            .Row("A1", "Name", "Qty", "Double")
            .Row("A2", "One", 1).SharedMaster("C2", 0, "C2:C6", "B2*2", 2)
            .Row("A3", "Two", 2).Shared("C3", 0, 4)
            .Row("A4", "Three", 3).Shared("C4", 0, 6)
            .Row("A5", "Four", 4).Shared("C5", 0, 8)
            .Row("A6", "Five", 5).Shared("C6", 0, 10);
        AddCheckSheet(book, "05 shared formulas: card Three removed; Two's Double typed as =B3*10.");

        var (editor, sheet) = Open(book, "Shared");
        editor.RemoveCard(sheet, sheet.Cards[2]);
        Type(editor, sheet, 1, "Double", "=B3*10");
        return editor.Workbook;
    }

    /// <summary>
    /// 06: a sheet renamed (to a name that needs quotes) that formulas and a defined name refer to; that sheet duplicated, with its content
    /// outside the table, column widths and page setup; a sheet removed; the sheet order changed.
    /// </summary>
    public static Workbook Sheets()
    {
        var book = new TestWorkbook();
        var plants = book.AddSheet("Plants")
            .Row("A1", "Name", "Height")
            .Row("A2", "Fern", 30)
            .Row("A3", "Rose", 45)
            .Set("G1", "Note beside the table")
            .Set("G2", "kept in the copy too")
            .Width(7, 7, 30);
        plants.PageSetup = true;
        book.AddSheet("Summary")
            .Row("A1", "Item", "Value")
            .Row("A2", "Fern height").Formula("B2", "Plants!B2", 30)
            .Row("A3", "Rose height").Formula("B3", "'Plants'!B3", 45)
            .Row("A4", "Named").Formula("B4", "FernHeight", 30)
            .Row("A5", "Sum").Formula("B5", "SUM(Plants!B2:B3)", 75);
        book.AddSheet("Old").Row("A1", "Name").Row("A2", "to be removed");
        book.AddDefinedName("FernHeight", "Plants!$B$2");
        AddCheckSheet(book, "06 sheets: Plants renamed Garden plants and duplicated; Old removed; Summary moved to the front.");

        var (editor, sheet) = Open(book, "Plants");
        var workbook = editor.Workbook;
        editor.RenameSheet(sheet, "Garden plants");
        var copy = editor.DuplicateSheet(sheet);
        Type(editor, copy, 0, "Height", "99");
        editor.RemoveSheet(workbook.Sheets.Single(s => s.Name == "Old"));
        editor.MoveSheet(workbook.Sheets.Single(s => s.Name == "Summary"), 0);
        return workbook;
    }

    /// <summary>The picture layout of the sheet "Pictures" in sample 07: 100 × 150 px units, 4 per row, 3 picture rows per worksheet row.</summary>
    public static readonly ImageLayout PicturesLayout = new(100, 150, 4);

    /// <summary>The picture layout of the sheet "Big pictures" in sample 07: 200 × 300 px, 1 per row, so 1 picture row per worksheet row.</summary>
    public static readonly ImageLayout BigLayout = new(200, 300, 1);

    /// <summary>
    /// 07: pictures laid out in their cells: tall, wide and cropped ones, several per cell, a card growing worksheet rows with content
    /// below it, columns widened; a picture read from the file removed and one outside the table kept; a sheet with tall units.
    /// </summary>
    public static Workbook Pictures()
    {
        var book = new TestWorkbook();
        book.AddSheet("Pictures")
            .Row("A1", "Name", "Photo", "Gallery")
            .Row("A2", "Tall, fits")
            .Row("A3", "Tall, cropped top and bottom")
            .Row("A4", "Wide, fits")
            .Row("A5", "Wide, cropped at the sides")
            .Row("A6", "Square, cropped top and bottom")
            .Row("A7", "Several in one cell")
            .Row("A8", "Thirteen, the card grows")
            .Row("A9", "Picture from the file removed")
            .Picture("B9", CheckPictures.Png(40, 60, 40, 60, 99), PictureAnchor.OneCell)
            .Set("A21", "Below the table").Formula("B21", "COUNTA(A2:A9)", 8)
            .Picture("H2", CheckPictures.Png(40, 60, 40, 60, 0), PictureAnchor.TwoCell);
        book.AddSheet("Big pictures").Row("A1", "Name", "Photo").Row("A2", "Two tall pictures").Row("A3", "After");
        AddCheckSheet(book, "07 pictures: every picture shows a white box with a round black ring and a number, no red.");

        var (editor, sheet) = Open(book, "Pictures");
        editor.SetPicturesPerRow(sheet, PicturesLayout.ImagesPerRow);
        editor.SetAspectLock(sheet, false);
        editor.SetPictureWidth(sheet, PicturesLayout.Width);
        editor.SetPictureHeight(sheet, PicturesLayout.Height);
        var photo = Column(sheet, "Photo");
        var gallery = Column(sheet, "Gallery");

        // Tall pictures are shown 100 × 150 (aspect 2:3), wide ones 200 × 150 (4:3).
        editor.AddPictures(sheet, sheet.Cards[0], photo, [CheckPictures.Png(40, 60, 2, 3, 1)]);
        editor.AddPictures(sheet, sheet.Cards[1], photo, [CheckPictures.Png(40, 90, 2, 3, 2)]);
        editor.AddPictures(sheet, sheet.Cards[2], photo, [CheckPictures.Png(80, 60, 4, 3, 3)]);
        editor.AddPictures(sheet, sheet.Cards[3], photo, [CheckPictures.Png(120, 60, 4, 3, 4)]);
        editor.AddPictures(sheet, sheet.Cards[4], photo, [CheckPictures.Png(64, 64, 4, 3, 5)]);
        editor.AddPictures(sheet, sheet.Cards[5], gallery,
        [
            CheckPictures.Png(40, 60, 2, 3, 1),
            CheckPictures.Png(80, 60, 4, 3, 2),
            CheckPictures.Png(40, 60, 2, 3, 3),
            CheckPictures.Png(40, 60, 2, 3, 4),
            CheckPictures.Png(80, 60, 4, 3, 5),
        ]);
        editor.AddPictures(sheet, sheet.Cards[6], gallery, [.. Enumerable.Range(1, 13).Select(n => CheckPictures.Png(40, 60, 2, 3, n))]);
        var removed = sheet.Cards[7].Entries[sheet.Columns.IndexOf(photo)].Images.Single();
        editor.RemovePicture(sheet, sheet.Cards[7], photo, removed);

        var big = editor.Workbook.Sheets.Single(s => s.Name == "Big pictures");
        editor.SetPicturesPerRow(big, BigLayout.ImagesPerRow);
        editor.SetAspectLock(big, false);
        editor.SetPictureWidth(big, BigLayout.Width);
        editor.SetPictureHeight(big, BigLayout.Height);
        editor.AddPictures(big, big.Cards[0], Column(big, "Photo"),
            [CheckPictures.Png(40, 60, 2, 3, 1), CheckPictures.Png(40, 60, 2, 3, 2)]);
        return editor.Workbook;
    }

    /// <summary>08: a workbook made from scratch, as File, New does: two sheets with entries, cards, values, a table name and a picture.</summary>
    public static Workbook NewWorkbook()
    {
        var editor = new WorkbookEditor(new Workbook());
        var sheet = editor.AddSheet("Sheet");
        editor.RenameEntry(sheet, sheet.Columns[0], "Name");
        editor.RenameEntry(sheet, editor.AddEntry(sheet), "Count");
        editor.RenameEntry(sheet, editor.AddEntry(sheet), "Picture");
        editor.SetTableName(sheet, "Made from scratch");
        Type(editor, sheet, 0, "Name", "First");
        Type(editor, sheet, 0, "Count", "12");
        editor.AddPictures(sheet, sheet.Cards[0], Column(sheet, "Picture"), [CheckPictures.Png(68, 90, 136, 180, 1)]);
        var second = editor.AddCard(sheet);
        Type(editor, sheet, second, "Name", "Second");
        Type(editor, sheet, second, "Count", "=B3+1");

        var check = editor.AddSheet("Sheet");
        editor.RenameSheet(check, "Check");
        editor.RenameEntry(check, check.Columns[0], "Check");
        Type(editor, check, 0, "Check", "08 new workbook: made in the app's model with no source file.");
        return editor.Workbook;
    }

    /// <summary>The size of the JPEGs in sample 09: wide, 2:1, the aspect of a wide picture's box at the sheet's 100 × 100 units.</summary>
    public const int OrientationWidth = 96;

    public const int OrientationHeight = 48;

    /// <summary>09: the same made-up JPEG twice, without EXIF and with EXIF orientation 6, to see whether Excel turns the second.</summary>
    public static Workbook JpegOrientation()
    {
        var book = new TestWorkbook();
        book.AddSheet("Orientation").Row("A1", "Name", "Picture").Row("A2", "No EXIF").Row("A3", "EXIF orientation 6");
        AddCheckSheet(book, "09 JPEG orientation: does Excel turn the picture whose EXIF orientation is 6?");

        var (editor, sheet) = Open(book, "Orientation");
        editor.SetAspectLock(sheet, false);
        editor.SetPictureWidth(sheet, 100);
        editor.SetPictureHeight(sheet, 100);
        var picture = Column(sheet, "Picture");
        editor.AddPictures(sheet, sheet.Cards[0], picture, [CheckPictures.OrientationJpeg(OrientationWidth, OrientationHeight, null)]);
        editor.AddPictures(sheet, sheet.Cards[1], picture, [CheckPictures.OrientationJpeg(OrientationWidth, OrientationHeight, 6)]);
        return editor.Workbook;
    }

    /// <summary>
    /// 10: cards added at both ends of a table with a total over every card and a reference to one card below it: the total covers both
    /// new cards, as an Excel Table's references do, and the reference follows its card.
    /// </summary>
    public static Workbook TableEnds()
    {
        var book = new TestWorkbook();
        book.AddSheet("Ends")
            .Row("A1", "Name", "Qty")
            .Row("A2", "Apple", 1)
            .Row("A3", "Banana", 2)
            .Row("A4", "Cherry", 3)
            .Set("A20", "Total").Formula("B20", "SUM(B2:B4)", 6)
            .Set("A21", "Apple's qty").Formula("B21", "B2", 1);
        AddCheckSheet(book, "10 table ends: Moss added at the top, Fig added at the end; the total counts both.");

        var (editor, sheet) = Open(book, "Ends");
        var moss = editor.InsertCardAbove(sheet, sheet.Cards[0]);
        Type(editor, sheet, moss, "Name", "Moss");
        Type(editor, sheet, moss, "Qty", "9");
        var fig = editor.AddCard(sheet);
        Type(editor, sheet, fig, "Name", "Fig");
        Type(editor, sheet, fig, "Qty", "6");
        return editor.Workbook;
    }

    /// <summary>Every sample, by file name, in the order of the notes.</summary>
    public static IReadOnlyList<(string FileName, Func<Workbook> Build)> All { get; } =
    [
        ("01-values.xlsx", Values),
        ("02-styles.xlsx", Styles),
        ("03-cards.xlsx", Cards),
        ("04-table-name.xlsx", TableName),
        ("05-shared-formulas.xlsx", SharedFormulas),
        ("06-sheets.xlsx", Sheets),
        ("07-pictures.xlsx", Pictures),
        ("08-new-workbook.xlsx", NewWorkbook),
        ("09-jpeg-orientation.xlsx", JpegOrientation),
        ("10-table-ends.xlsx", TableEnds),
    ];

    /// <summary>A Calibri 11 font, its children in the schema's order (b, i, sz, color, name).</summary>
    private static X.Font Calibri(bool bold = false, bool italic = false, string? color = null)
    {
        var font = new X.Font();
        if (bold)
        {
            font.AppendChild(new X.Bold());
        }

        if (italic)
        {
            font.AppendChild(new X.Italic());
        }

        font.AppendChild(new X.FontSize { Val = 11 });
        if (color is not null)
        {
            font.AppendChild(new X.Color { Rgb = color });
        }

        font.AppendChild(new X.FontName { Val = "Calibri" });
        return font;
    }

    private static void AddCheckSheet(TestWorkbook book, string line) =>
        book.AddSheet("Check").Row("A1", "Check").Row("A2", line).Row("A3", "What to look at: docs/notes/excel-check.md in the repository.");

    private static (WorkbookEditor Editor, Sheet Sheet) Open(TestWorkbook book, string sheetName)
    {
        var bytes = book.ToBytes();
        var errors = PackageValidation.Errors(bytes);
        if (errors.Count > 0)
        {
            throw new InvalidOperationException("The source workbook is not valid:\n" + string.Join("\n", errors));
        }

        using var stream = new MemoryStream(bytes);
        var editor = new WorkbookEditor(XlsxReader.Read(stream, Culture));
        return (editor, editor.Workbook.Sheets.Single(s => s.Name == sheetName));
    }

    private static Column Column(Sheet sheet, string name) => sheet.Columns.Single(c => c.Name == name);

    private static void Type(WorkbookEditor editor, Sheet sheet, int card, string column, string text) =>
        Type(editor, sheet, sheet.Cards[card], column, text);

    private static void Type(WorkbookEditor editor, Sheet sheet, Card card, string column, string text)
    {
        editor.SetEntryText(sheet, card, Column(sheet, column), text);
        editor.EndTyping();
    }
}
