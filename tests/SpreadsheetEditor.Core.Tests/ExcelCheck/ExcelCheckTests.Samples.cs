using DocumentFormat.OpenXml.Packaging;
using SpreadsheetEditor.Core.Model;
using SpreadsheetEditor.Core.Tests.Xlsx;
using X = DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace SpreadsheetEditor.Core.Tests.ExcelCheck;

/// <summary>One test per sample: what docs/notes/excel-check.md tells the owner to look for, checked on the saved file.</summary>
public sealed partial class ExcelCheckTests
{
    [TestMethod]
    public void SampleValues()
    {
        var saved = Save("01-values.xlsx");
        var sheet = saved.Sheet("Values");

        AssertNumber(Value(sheet, "Decimal typed"), 3.5);
        AssertNumber(Value(sheet, "Whole number typed"), 42);
        AssertDate(Value(sheet, "Date typed"), new DateTime(2024, 3, 15));
        AssertDate(Value(sheet, "Date and time typed"), new DateTime(2024, 3, 15, 14, 30, 0));
        var time = Value(sheet, "Time typed").Original!;
        Assert.AreEqual(CellValueKind.Date, time.Kind);
        Assert.AreEqual(new TimeSpan(14, 30, 0), time.Date!.Value.TimeOfDay);
        Assert.IsTrue(Value(sheet, "TRUE typed").Original!.Boolean);
        Assert.IsFalse(Value(sheet, "false typed").Original!.Boolean);
        Assert.AreEqual("B2*2", Value(sheet, "Formula typed").Original!.Formula);
        Assert.AreEqual("B3+1", Value(sheet, "Formula kept").Original!.Formula);
        AssertText(Value(sheet, "Digits typed in a Text cell"), "007");
        AssertText(Value(sheet, "Leading zeros typed"), "007");
        AssertText(Value(sheet, "Words with a number typed"), "12 apples");
        var kept = Value(sheet, "Date kept").Original!;
        Assert.AreEqual(CellValueKind.Date, kept.Kind);
        Assert.AreEqual(45000, kept.Number);
        AssertNumber(Value(sheet, "Number kept"), 1234.5);
        Assert.IsTrue(Value(sheet, "Boolean kept").Original!.Boolean);
        AssertText(Value(sheet, "Text kept"), "plain words");
        Assert.AreEqual("", Value(sheet, "Cleared").Text);

        using var package = new XlsxPackage(saved.Bytes);
        Assert.IsTrue(package.Workbook.CalculationProperties?.FullCalculationOnLoad?.Value, "Excel is asked to recalculate on open.");
        Assert.IsNull(package.Cell("Values", "B18")?.CellValue, "The cleared cell has no value.");
    }

    [TestMethod]
    public void SampleStyles()
    {
        var saved = Save("02-styles.xlsx");
        var sheet = saved.Sheet("Styles");

        Assert.AreEqual(new RgbColor(0xFF, 0xFF, 0x00), Value(sheet, "Fill set to yellow").Style!.Fill);
        Assert.IsNull(Value(sheet, "Green fill removed").Style?.Fill);
        Assert.AreEqual(new RgbColor(0x00, 0x46, 0xC8), Value(sheet, "Text colour set to blue").Style!.TextColor);
        var automatic = Value(sheet, "Red text set to Automatic");
        Assert.IsNull(automatic.Style?.TextColor);
        Assert.AreEqual("automatic text", automatic.Text);
        var boldItalic = Value(sheet, "Bold italic, text edited");
        Assert.AreEqual("still bold italic", boldItalic.Text);
        Assert.IsTrue(boldItalic.Style!.Font!.Bold);
        Assert.IsTrue(boldItalic.Style.Font.Italic);
        var untouched = Value(sheet, "Orange bold, untouched").Style!;
        Assert.AreEqual(new RgbColor(0xFF, 0xC0, 0x00), untouched.Fill);
        Assert.IsTrue(untouched.Font!.Bold);
        var multiLine = Value(sheet, "Multi-line text typed");
        Assert.AreEqual("line one\nline two\nline three", multiLine.Text);
        Assert.IsTrue(multiLine.Style!.WrapText);
        Assert.AreEqual("top", multiLine.Style.VerticalAlignment);
        var centred = Value(sheet, "Centred, text edited");
        Assert.AreEqual("still centred", centred.Text);
        Assert.AreEqual("center", centred.Style!.HorizontalAlignment);
    }

    [TestMethod]
    public void SampleCards()
    {
        var saved = Save("03-cards.xlsx");
        var sheet = saved.Sheet("Cards");

        AssertSequence(
            sheet.Cards.Select(c => $"{c.Entries[0].Text} {c.Entries[1].Text}").ToArray(), "Elder 5", "Apple 1", "Grape 7", "Cherry 3", "Damson 4",
            "Fig 6");

        // One card removed and two added: everything below the table moved down one row, from row 20 to 21.
        using var package = new XlsxPackage(saved.Bytes);
        Assert.AreEqual("Total", package.Text("Cards", "A21"));
        Assert.AreEqual("SUM(B2:B6)", Formula(package, "Cards", "B21"));
        Assert.AreEqual("COUNTA(A2:A6)", Formula(package, "Cards", "B22"));
        Assert.AreEqual("B3", Formula(package, "Cards", "B23"));
        Assert.AreEqual("#REF!", Formula(package, "Cards", "B24"));
        Assert.AreEqual("merged note below the table", package.Text("Cards", "D21"));
        CollectionAssert.Contains(Merges(package, "Cards"), "D21:E21");
        Assert.AreEqual("yellow cell below the table", package.Text("Cards", "A25"));
        Assert.AreEqual("Cards!B21", Formula(package, "Report", "B2"));
        Assert.AreEqual("Cards!B2", Formula(package, "Report", "B3"));
        Assert.AreEqual("Cards!$B$21", DefinedName(package, "CardTotal"));
        Assert.AreEqual("Cards!$A$1:$E$25", DefinedName(package, "_xlnm.Print_Area", 0));
    }

    [TestMethod]
    public void SampleTableName()
    {
        var saved = Save("04-table-name.xlsx");
        var sheet = saved.Sheet("Named");

        Assert.AreEqual("Stationery", sheet.TableName);
        Assert.AreEqual(2, sheet.HeaderRow, "The header moved from row 1 to row 2.");
        AssertSequence(sheet.Cards.Select(c => c.Entries[0].Text).ToArray(), "Pen", "Ink", "Pad");

        using var package = new XlsxPackage(saved.Bytes);
        Assert.AreEqual("Stationery", package.Text("Named", "A1"));
        Assert.AreEqual("Name", package.Text("Named", "A2"));
        Assert.AreEqual("Total", package.Text("Named", "A16"));
        Assert.AreEqual("SUM(B3:B5)", Formula(package, "Named", "B16"));
        Assert.AreEqual("Named!B16", Formula(package, "Report", "B2"));
    }

    [TestMethod]
    public void SampleSharedFormulas()
    {
        var saved = Save("05-shared-formulas.xlsx");
        var sheet = saved.Sheet("Shared");

        AssertSequence(sheet.Cards.Select(c => c.Entries[0].Text).ToArray(), "One", "Two", "Four", "Five");
        AssertSequence(sheet.Cards.Select(c => c.Entries[2].Original?.Formula).ToArray(), "B2*2", "B3*10", "B4*2", "B5*2");

        // Every formula cell can be read on its own: a shared formula's other cells name a master that is still there.
        using var package = new XlsxPackage(saved.Bytes);
        var cells = package.Worksheet("Shared").Descendants<X.Cell>().Where(c => c.CellFormula is not null).ToList();
        var masters = cells.Where(c => c.CellFormula!.FormulaType?.Value == X.CellFormulaValues.Shared && c.CellFormula.Reference is not null)
            .Select(c => c.CellFormula!.SharedIndex!.Value)
            .ToHashSet();
        foreach (var cell in cells.Where(c => c.CellFormula!.FormulaType?.Value == X.CellFormulaValues.Shared))
        {
            Assert.Contains(cell.CellFormula!.SharedIndex!.Value, masters, $"{cell.CellReference} names a shared formula without a master.");
        }
    }

    [TestMethod]
    public void SampleSheets()
    {
        var saved = Save("06-sheets.xlsx");

        using var package = new XlsxPackage(saved.Bytes);
        AssertSequence(package.SheetNames, "Summary", "Garden plants", "Garden plants (Copy)", "Check");
        Assert.AreEqual("'Garden plants'!B2", Formula(package, "Summary", "B2"));
        Assert.AreEqual("'Garden plants'!B3", Formula(package, "Summary", "B3"));
        Assert.AreEqual("FernHeight", Formula(package, "Summary", "B4"));
        Assert.AreEqual("SUM('Garden plants'!B2:B3)", Formula(package, "Summary", "B5"));
        Assert.AreEqual("'Garden plants'!$B$2", DefinedName(package, "FernHeight"));
        foreach (var name in new[] { "Garden plants", "Garden plants (Copy)" })
        {
            Assert.AreEqual("Note beside the table", package.Text(name, "G1"), name);
            Assert.AreEqual("kept in the copy too", package.Text(name, "G2"), name);
            var worksheet = package.Worksheet(name);
            Assert.AreEqual(30, worksheet.GetFirstChild<X.Columns>()!.Elements<X.Column>().Single(c => c.Min! == 7u).Width!.Value, name);
            Assert.AreEqual(X.OrientationValues.Landscape, worksheet.GetFirstChild<X.PageSetup>()!.Orientation!.Value, name);
        }

        Assert.AreEqual("30", package.Text("Garden plants", "B2"));
        Assert.AreEqual("99", package.Text("Garden plants (Copy)", "B2"));
    }

    [TestMethod]
    public void SamplePictures()
    {
        var saved = Save("07-pictures.xlsx");
        var sheet = saved.Sheet("Pictures");
        Assert.AreEqual(ExcelCheckSamples.PicturesLayout, sheet.ImageLayout);
        int[] counts = [1, 1, 1, 1, 1, 0, 0, 0];
        int[] gallery = [0, 0, 0, 0, 0, 5, 13, 0];
        CollectionAssert.AreEqual(counts, sheet.Cards.Select(c => c.Entries[1].Images.Count).ToArray());
        CollectionAssert.AreEqual(gallery, sheet.Cards.Select(c => c.Entries[2].Images.Count).ToArray());
        Assert.AreEqual(2, sheet.Cards[6].SourceRowSpan, "The card of thirteen pictures spans two worksheet rows.");

        using var package = new XlsxPackage(saved.Bytes);
        // Every picture in the table: (row, column, x, y, width, height) in px, in the order written.
        var anchors = OneCellAnchors(package, "Pictures");
        (int, int, long, long, long, long)[] expected =
        [
            (2, 2, 0, 0, 100, 150), (3, 2, 0, 0, 100, 150), (4, 2, 0, 0, 200, 150), (5, 2, 0, 0, 200, 150), (6, 2, 0, 0, 200, 150),
            (7, 3, 0, 0, 100, 150), (7, 3, 100, 0, 200, 150), (7, 3, 300, 0, 100, 150), (7, 3, 0, 150, 100, 150), (7, 3, 100, 150, 200, 150),
            .. Enumerable.Range(0, 13).Select(n => (8 + (n / 12), 3, (long)(n % 4 * 100), (long)(n / 4 % 3 * 150), 100L, 150L)),
        ];
        CollectionAssert.AreEqual(expected, anchors.Select(a => (a.Row, a.Column, a.X, a.Y, a.Width, a.Height)).ToArray());
        AssertSequence(
            anchors.Take(5).Select(a => a.Crop).ToArray(), "0,0,0,0", "0,16666,0,16666", "0,0,0,0", "16666,0,16666,0",
            "0,12500,0,12500");
        Assert.IsTrue(anchors.Skip(5).All(a => a.Crop == "0,0,0,0"), "Pictures that fit their box are not cropped.");

        // The picture outside the table is kept as it was read; the one read in B9 was removed.
        var drawing = package.Part("Pictures").DrawingsPart!.WorksheetDrawing!;
        var outside = drawing.Elements<Xdr.TwoCellAnchor>().Single();
        Assert.AreEqual("7", outside.FromMarker!.ColumnId!.Text);
        Assert.AreEqual("1", outside.FromMarker.RowId!.Text);
        Assert.AreEqual(anchors.Count + 1, drawing.ChildElements.Count);

        AssertSequence(
            RowHeights(package, "Pictures"), ("2", 112.5), ("3", 112.5), ("4", 112.5), ("5", 112.5), ("6", 112.5), ("7", 225.0), ("8", 342.0),
            ("9", 112.5));
        AssertSet(Merges(package, "Pictures"), "A8:A9", "B8:B9", "C8:C9");
        Assert.AreEqual(Math.Round(195 / 7.0, 6), Math.Round(ColumnWidth(package, "Pictures", 2), 6));
        Assert.AreEqual(Math.Round(395 / 7.0, 6), Math.Round(ColumnWidth(package, "Pictures", 3), 6));
        Assert.AreEqual("Below the table", package.Text("Pictures", "A22"));
        Assert.AreEqual("COUNTA(A2:A10)", Formula(package, "Pictures", "B22"), "The row inserted inside the range widens it.");

        var big = OneCellAnchors(package, "Big pictures");
        AssertSequence(
            big.Select(a => (a.Row, a.Column, a.X, a.Y, a.Width, a.Height)).ToArray(), (2, 2, 0L, 0L, 200L, 300L),
            (3, 2, 0L, 0L, 200L, 300L));
        AssertSequence(RowHeights(package, "Big pictures"), ("2", 229.5), ("3", 229.5));
        Assert.AreEqual("After", package.Text("Big pictures", "A4"));
    }

    [TestMethod]
    public void SampleNewWorkbook()
    {
        var saved = Save("08-new-workbook.xlsx");
        var sheet = saved.Sheet("Sheet1");

        Assert.AreEqual("Made from scratch", sheet.TableName);
        AssertSequence(sheet.Columns.Select(c => c.Name).ToArray(), "Name", "Count", "Picture");
        AssertSequence(sheet.Cards.Select(c => c.Entries[0].Text).ToArray(), "First", "Second");
        AssertNumber(sheet.Cards[0].Entries[1], 12);
        Assert.AreEqual("B3+1", sheet.Cards[1].Entries[1].Original!.Formula);
        Assert.HasCount(1, sheet.Cards[0].Entries[2].Images);

        using var package = new XlsxPackage(saved.Bytes);
        AssertSequence(package.SheetNames, "Sheet1", "Check");
        AssertSequence(
            OneCellAnchors(package, "Sheet1").Select(a => (a.Row, a.Column, a.X, a.Y, a.Width, a.Height)).ToArray(),
            (3, 3, 0L, 0L, 136L, 180L));
    }

    [TestMethod]
    public void SampleJpegOrientation()
    {
        var saved = Save("09-jpeg-orientation.xlsx");
        var sheet = saved.Sheet("Orientation");
        Assert.IsTrue(sheet.Cards.All(c => c.Entries[1].Images.Single().Format == ImageFormat.Jpeg));

        using var package = new XlsxPackage(saved.Bytes);
        var anchors = OneCellAnchors(package, "Orientation");
        AssertSequence(
            anchors.Select(a => (a.Row, a.Column, a.X, a.Y, a.Width, a.Height)).ToArray(), (2, 2, 0L, 0L, 200L, 100L),
            (3, 2, 0L, 0L, 200L, 100L));
        int w = ExcelCheckSamples.OrientationWidth, h = ExcelCheckSamples.OrientationHeight;
        CollectionAssert.AreEqual(CheckPictures.OrientationJpeg(w, h, null), anchors[0].Bytes, "The plain JPEG is embedded unchanged.");
        CollectionAssert.AreEqual(CheckPictures.OrientationJpeg(w, h, 6), anchors[1].Bytes, "The EXIF JPEG is embedded unchanged.");
    }

    private static void AssertSequence<T>(IEnumerable<T> actual, params T[] expected) =>
        CollectionAssert.AreEqual(expected, actual.ToArray());

    private static void AssertSet<T>(IEnumerable<T> actual, params T[] expected) =>
        CollectionAssert.AreEquivalent(expected, actual.ToArray());

    private static Entry Value(Sheet sheet, string item) =>
        sheet.Cards.Single(c => c.Entries[0].Text == item).Entries[1];

    private static void AssertNumber(Entry entry, double expected)
    {
        Assert.AreEqual(CellValueKind.Number, entry.Original?.Kind, entry.Text);
        Assert.AreEqual(expected, entry.Original!.Number);
    }

    private static void AssertDate(Entry entry, DateTime expected)
    {
        Assert.AreEqual(CellValueKind.Date, entry.Original?.Kind, entry.Text);
        Assert.AreEqual(expected, entry.Original!.Date);
    }

    private static void AssertText(Entry entry, string expected)
    {
        Assert.AreEqual(CellValueKind.Text, entry.Original?.Kind, entry.Text);
        Assert.AreEqual(expected, entry.Text);
    }

    private static string? Formula(XlsxPackage package, string sheet, string reference) => package.Cell(sheet, reference)?.CellFormula?.Text;

    private static List<string> Merges(XlsxPackage package, string sheet) =>
        package.Worksheet(sheet).Descendants<X.MergeCell>().Select(m => m.Reference!.Value!).ToList();

    private static string DefinedName(XlsxPackage package, string name, uint? localSheetId = null) =>
        package.Workbook.DefinedNames!.Elements<X.DefinedName>()
            .Single(n => n.Name == name && n.LocalSheetId?.Value == localSheetId).Text;

    private static (string, double)[] RowHeights(XlsxPackage package, string sheet) =>
        package.Worksheet(sheet).Descendants<X.Row>().Where(r => r.Height is not null)
            .Select(r => (r.RowIndex!.Value.ToString(System.Globalization.CultureInfo.InvariantCulture), r.Height!.Value))
            .ToArray();

    private static double ColumnWidth(XlsxPackage package, string sheet, uint column) =>
        package.Worksheet(sheet).GetFirstChild<X.Columns>()!.Elements<X.Column>().Single(c => c.Min! <= column && c.Max! >= column).Width!.Value;

    /// <summary>A one-cell anchor: its cell (1-based), offsets and size in px, source rectangle "l,t,r,b" and picture bytes.</summary>
    private sealed record Anchor(int Row, int Column, long X, long Y, long Width, long Height, string Crop, byte[] Bytes);

    private static List<Anchor> OneCellAnchors(XlsxPackage package, string sheet)
    {
        const long EmuPerPixel = 9525;
        var drawing = package.Part(sheet).DrawingsPart!;
        var result = new List<Anchor>();
        foreach (var anchor in drawing.WorksheetDrawing!.Elements<Xdr.OneCellAnchor>())
        {
            var from = anchor.FromMarker!;
            var fill = anchor.GetFirstChild<Xdr.Picture>()!.BlipFill!;
            var source = fill.SourceRectangle;
            var crop = string.Join(",", new[] { source?.Left, source?.Top, source?.Right, source?.Bottom }.Select(v => v?.Value ?? 0));
            using var stream = ((ImagePart)drawing.GetPartById(fill.Blip!.Embed!.Value!)).GetStream();
            using var copy = new MemoryStream();
            stream.CopyTo(copy);
            result.Add(new Anchor(
                int.Parse(from.RowId!.Text, System.Globalization.CultureInfo.InvariantCulture) + 1,
                int.Parse(from.ColumnId!.Text, System.Globalization.CultureInfo.InvariantCulture) + 1,
                long.Parse(from.ColumnOffset!.Text, System.Globalization.CultureInfo.InvariantCulture) / EmuPerPixel,
                long.Parse(from.RowOffset!.Text, System.Globalization.CultureInfo.InvariantCulture) / EmuPerPixel,
                anchor.Extent!.Cx!.Value / EmuPerPixel,
                anchor.Extent.Cy!.Value / EmuPerPixel,
                crop,
                copy.ToArray()));
        }

        return result;
    }
}
