using System.Globalization;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using SpreadsheetEditor.Core.Model;
using SpreadsheetEditor.Core.Tests.Imaging;
using SpreadsheetEditor.Core.Xlsx;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetEditor.Core.Tests.Xlsx;

[TestClass]
public sealed class XlsxWriterTests
{
    private static readonly CultureInfo Invariant = CultureInfo.InvariantCulture;

    private static Workbook Read(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        return XlsxReader.Read(stream, Invariant);
    }

    private static byte[] Write(Workbook workbook)
    {
        using var stream = new MemoryStream();
        XlsxWriter.Write(workbook, stream, Invariant);
        return stream.ToArray();
    }

    private static byte[] Change(byte[] original, Action<Workbook> change)
    {
        var workbook = Read(original);
        change(workbook);
        return Write(workbook);
    }

    private static SpreadsheetWriteError WriteError(byte[] original, Action<Workbook> change)
    {
        var workbook = Read(original);
        change(workbook);
        return Assert.ThrowsExactly<SpreadsheetWriteException>(() => Write(workbook)).Reason;
    }

    private static void AssertSequence(IEnumerable<string> actual, params string[] expected) =>
        CollectionAssert.AreEqual(expected, actual.ToArray());

    private static string[] Texts(Card card) => card.Entries.Select(e => e.Text).ToArray();

    private static void AssertSameXml(OpenXmlPart expected, OpenXmlPart actual) =>
        Assert.IsTrue(XNode.DeepEquals(XlsxPackage.Xml(expected), XlsxPackage.Xml(actual)), $"{actual.Uri} changed.");

    private static TestWorkbook Plants()
    {
        var book = new TestWorkbook();
        book.AddSheet("Plants").Row("A1", "Name", "Colour", "Height").Row("A2", "Fern", "Green", 30).Row("A3", "Rose", "Red", 45);
        return book;
    }

    [TestMethod]
    public void UnchangedWorkbookReadsBackTheSame()
    {
        var book = new TestWorkbook();
        var styled = book.AddNumberFormat(10);
        book.AddSheet("Plants").Row("A1", "Name", "Count", "Ok").Row("A2", "Fern", 3, true).Set("B3", 0.5, styled).Set("A3", "Rose")
            .Formula("D2", "B2*2", 6).Set("J20", "beside").PageSetup = true;
        book.AddSheet("Empty");
        book.AddChartSheet("Chart");
        book.AddSheet("Tools").Row("A1", "Tool").Row("A2", "Spade");
        book.AddDefinedName("Counts", "Plants!$B$2:$B$3");
        book.AddDefinedName("_xlnm.Print_Area", "Plants!$A$1:$C$3", 0);
        var original = book.ToBytes();

        var bytes = Change(original, _ => { });

        var back = Read(bytes);
        Assert.AreEqual(2, back.Sheets.Count);
        AssertSequence(Texts(back.Sheets[0].Cards[0]), "Fern", "3", "TRUE", "=B2*2");
        using var before = new XlsxPackage(original);
        using var after = new XlsxPackage(bytes);
        AssertSequence(after.SheetNames, "Plants", "Empty", "Chart", "Tools");
        AssertSameXml(before.Part("Empty"), after.Part("Empty"));
        AssertSameXml(before.WorkbookPart.WorkbookStylesPart!, after.WorkbookPart.WorkbookStylesPart!);
        Assert.AreEqual(before.Workbook.DefinedNames!.OuterXml, after.Workbook.DefinedNames!.OuterXml);
        // Every cell of the table, the formula and the cell beside it are as they were; only the metadata cells were added.
        var cellsBefore = before.Worksheet("Plants").Descendants<X.Cell>().Select(c => c.OuterXml).ToList();
        var cellsAfter = after.Worksheet("Plants").Descendants<X.Cell>().Where(c => !c.CellReference!.Value!.StartsWith("XFD", StringComparison.Ordinal))
            .Select(c => c.OuterXml).ToList();
        CollectionAssert.AreEqual(cellsBefore, cellsAfter);
        Assert.AreEqual(before.Worksheet("Plants").GetFirstChild<X.PageSetup>()!.OuterXml, after.Worksheet("Plants").GetFirstChild<X.PageSetup>()!.OuterXml);
        // Nothing but metadata was written: no value changed, so no recalculation is asked for.
        Assert.IsNull(after.Workbook.CalculationProperties);
    }

    [TestMethod]
    public void TwoSavesGiveTheSameFile()
    {
        var workbook = Read(Plants().ToBytes());
        workbook.Sheets[0].Cards[0].Entries[2].Text = "31";
        workbook.Sheets[0].Cards[1].Entries[0].Text = "Tulip";

        var first = Write(workbook);
        var second = Write(workbook);

        CollectionAssert.AreEqual(first, second);
        // Zip entries carry a fixed time, not the time of the save, so saves a second apart are the same too.
        using var zip = new System.IO.Compression.ZipArchive(new MemoryStream(first));
        Assert.IsTrue(zip.Entries.All(e => e.LastWriteTime.Year == 1980));
    }

    [TestMethod]
    public void TwoSavesOfANewWorkbookGiveTheSameFile()
    {
        var workbook = NewWorkbook();

        CollectionAssert.AreEqual(Write(workbook), Write(workbook));
    }

    [TestMethod]
    public void WritesMetadataAndHidesTheLastColumn()
    {
        var original = Plants().ToBytes();

        var bytes = Change(original, w => w.Sheets[0].ImageLayout = new ImageLayout(200, 150, 12));

        using var after = new XlsxPackage(bytes);
        Assert.AreEqual("ImageHeight:150", after.Text("Plants", "XFD1"));
        Assert.AreEqual("ExcelImagesPerRow:8", after.Text("Plants", "XFD2"));
        Assert.AreEqual("ImageWidth:200", after.Text("Plants", "XFD3"));
        var column = after.Worksheet("Plants").Descendants<X.Column>().Single(c => c.Min!.Value == 16384);
        Assert.IsTrue(column.Hidden!.Value);
        Assert.AreEqual(new ImageLayout(200, 150, 8), Read(bytes).Sheets[0].ImageLayout);
    }

    [TestMethod]
    public void HidingTheLastColumnSplitsAColumnRangeThatCoversIt()
    {
        var book = Plants();
        book.AddSheet("Wide").Width(5, 16384, 20).Row("A1", "Name", "Kind").Row("A2", "Oak", "Tree");

        var bytes = Change(book.ToBytes(), _ => { });

        using var after = new XlsxPackage(bytes);
        var columns = after.Worksheet("Wide").Descendants<X.Column>().ToList();
        Assert.AreEqual(2, columns.Count);
        Assert.AreEqual((5u, 16383u, 20.0, (bool?)null), (columns[0].Min!.Value, columns[0].Max!.Value, columns[0].Width!.Value, columns[0].Hidden?.Value));
        Assert.AreEqual((16384u, 16384u, 20.0, (bool?)true), (columns[1].Min!.Value, columns[1].Max!.Value, columns[1].Width!.Value, columns[1].Hidden?.Value));
    }

    [TestMethod]
    public void EditedValuesAreWrittenTyped()
    {
        var original = Plants().ToBytes();

        var bytes = Change(original, w =>
        {
            var cards = w.Sheets[0].Cards;
            cards[0].Entries[0].Text = "Moss";
            cards[0].Entries[1].Text = "TRUE";
            cards[0].Entries[2].Text = "12.5";
            cards[1].Entries[1].Text = "=A3&\"!\"";
            cards[1].Entries[2].Text = "01/15/2024";
        });

        using var after = new XlsxPackage(bytes);
        Assert.AreEqual("Moss", after.Text("Plants", "A2"));
        Assert.AreEqual(X.CellValues.Boolean, after.Cell("Plants", "B2")!.DataType!.Value);
        Assert.AreEqual("1", after.Text("Plants", "B2"));
        Assert.IsNull(after.Cell("Plants", "C2")!.DataType);
        Assert.AreEqual("12.5", after.Text("Plants", "C2"));
        var formula = after.Cell("Plants", "B3")!;
        Assert.AreEqual("A3&\"!\"", formula.CellFormula!.Text);
        Assert.IsNull(formula.CellValue);
        // 15 January 2024 is serial 45306 in the 1900 date system, shown with the default short date format (id 14).
        var date = after.Cell("Plants", "C3")!;
        Assert.AreEqual("45306", date.CellValue!.Text);
        var xf = after.WorkbookPart.WorkbookStylesPart!.Stylesheet!.CellFormats!.Elements<X.CellFormat>().ElementAt((int)date.StyleIndex!.Value);
        Assert.AreEqual(14u, xf.NumberFormatId!.Value);
        Assert.IsTrue(after.Workbook.CalculationProperties!.FullCalculationOnLoad!.Value);

        var back = Read(bytes).Sheets[0];
        AssertSequence(Texts(back.Cards[0]), "Moss", "TRUE", "12.5");
        AssertSequence(Texts(back.Cards[1]), "Rose", "=A3&\"!\"", "01/15/2024");
        Assert.AreEqual(CellValueKind.Date, back.Cards[1].Entries[2].Original!.Kind);
    }

    [TestMethod]
    public void UnchangedFormulasAndTheirCachedValuesStay()
    {
        var book = new TestWorkbook();
        book.AddSheet("Sums").Row("A1", "Name", "Value").Row("A2", "One", 1).Row("A3", "Two", 2).Formula("B4", "SUM(B2:B3)", 3).Set("A4", "Total")
            .Row("C1", "Double").SharedFormula("C2", "B2*2", "C2:C4", 2).SharedDependent("C3", 4).SharedDependent("C4", 6);
        var original = book.ToBytes();

        var bytes = Change(original, w => w.Sheets[0].Cards[0].Entries[0].Text = "First");

        using var before = new XlsxPackage(original);
        using var after = new XlsxPackage(bytes);
        foreach (var reference in new[] { "B4", "C2", "C3", "C4" })
        {
            Assert.AreEqual(before.Cell("Sums", reference)!.OuterXml, after.Cell("Sums", reference)!.OuterXml, reference);
        }
    }

    [TestMethod]
    public void EditingASharedFormulaCellThatOthersUseIsNotSupported()
    {
        var book = new TestWorkbook();
        book.AddSheet("Sums").Row("A1", "Name", "Double").Row("A2", "One").Row("A3", "Two")
            .SharedFormula("B2", "ROW()*2", "B2:B3", 4).SharedDependent("B3", 6);
        var original = book.ToBytes();

        Assert.AreEqual(SpreadsheetWriteError.NotSupported, WriteError(original, w => w.Sheets[0].Cards[0].Entries[1].Text = "5"));

        // A cell that only uses the shared formula may be changed: it leaves the group.
        var bytes = Change(original, w => w.Sheets[0].Cards[1].Entries[1].Text = "5");
        using var after = new XlsxPackage(bytes);
        Assert.IsNull(after.Cell("Sums", "B3")!.CellFormula);
        Assert.AreEqual("5", after.Text("Sums", "B3"));
    }

    [TestMethod]
    public void ChangingAFormulaCellDropsTheCalculationChain()
    {
        var book = new TestWorkbook { CalculationChain = true };
        book.AddSheet("Sums").Row("A1", "Name", "Value").Row("A2", "One", 1).Formula("B3", "B2+1", 2).Set("A3", "Two");
        var original = book.ToBytes();

        var unchanged = Change(original, w => w.Sheets[0].Cards[0].Entries[0].Text = "Uno");
        var changed = Change(original, w => w.Sheets[0].Cards[1].Entries[1].Text = "7");

        using (var after = new XlsxPackage(unchanged))
        {
            Assert.IsNotNull(after.WorkbookPart.CalculationChainPart);
        }

        using (var after = new XlsxPackage(changed))
        {
            Assert.IsNull(after.WorkbookPart.CalculationChainPart);
            Assert.AreEqual("7", after.Text("Sums", "B3"));
        }
    }

    [TestMethod]
    public void ADateBeforeThe1904SystemStartIsWrittenAsText()
    {
        var book = new TestWorkbook { Date1904 = true };
        book.AddSheet("Dates").Row("A1", "Name", "When").Row("A2", "Old", "x");

        var bytes = Change(book.ToBytes(), w => w.Sheets[0].Cards[0].Entries[1].Text = "06/15/1902");

        using var after = new XlsxPackage(bytes);
        Assert.AreEqual(X.CellValues.SharedString, after.Cell("Dates", "B2")!.DataType!.Value);
        Assert.AreEqual("06/15/1902", after.Text("Dates", "B2"));
    }

    [TestMethod]
    public void TextFormattedCellsKeepTypedNumbersAsText()
    {
        var book = new TestWorkbook();
        var text = book.AddNumberFormat(49);
        book.AddSheet("Codes").Row("A1", "Name", "Code").Set("A2", "Gate").Set("B2", "x", text);

        var bytes = Change(book.ToBytes(), w => w.Sheets[0].Cards[0].Entries[1].Text = "007");

        using var after = new XlsxPackage(bytes);
        var cell = after.Cell("Codes", "B2")!;
        Assert.AreEqual(X.CellValues.SharedString, cell.DataType!.Value);
        Assert.AreEqual("007", after.Text("Codes", "B2"));
        Assert.AreEqual(text, cell.StyleIndex!.Value);
    }

    [TestMethod]
    public void AnEmptyValueClearsTheCellAndKeepsItsStyle()
    {
        var book = new TestWorkbook();
        var style = book.AddNumberFormat(10);
        book.AddSheet("Plants").Row("A1", "Name", "Share").Set("A2", "Fern").Set("B2", 0.25, style).Row("A3", "Rose", 0.5);

        var bytes = Change(book.ToBytes(), w => w.Sheets[0].Cards[0].Entries[1].Text = "");

        using var after = new XlsxPackage(bytes);
        var cell = after.Cell("Plants", "B2")!;
        Assert.IsNull(cell.CellValue);
        Assert.AreEqual(style, cell.StyleIndex!.Value);
        Assert.AreEqual("", Read(bytes).Sheets[0].Cards[0].Entries[1].Text);
    }

    [TestMethod]
    public void HeadersAreWrittenOnlyWhenRenamed()
    {
        var book = new TestWorkbook();
        book.AddSheet("Years").Row("A1", "Name", 2024, "Price").Row("A2", "Plan", 3, 9);
        var original = book.ToBytes();

        var bytes = Change(original, w => w.Sheets[0].Columns[2].Name = "Cost");

        using var before = new XlsxPackage(original);
        using var after = new XlsxPackage(bytes);
        // The numeric header was not renamed: it stays a number.
        Assert.AreEqual(before.Cell("Years", "B1")!.OuterXml, after.Cell("Years", "B1")!.OuterXml);
        Assert.AreEqual("Cost", after.Text("Years", "C1"));
        AssertSequence(Read(bytes).Sheets[0].Columns.Select(c => c.Name).ToArray(), "Name", "2024", "Cost");
    }

    [TestMethod]
    public void AClearedTableNameClearsItsCell()
    {
        var book = new TestWorkbook();
        book.AddSheet("Plants").Row("A1", "Garden").Row("A2", "Name", "Colour").Row("A3", "Fern", "Green");
        var original = book.ToBytes();

        var renamed = Change(original, w => w.Sheets[0].TableName = "Yard");
        var cleared = Change(original, w => w.Sheets[0].TableName = null);

        using (var after = new XlsxPackage(renamed))
        {
            Assert.AreEqual("Yard", after.Text("Plants", "A1"));
        }

        using (var after = new XlsxPackage(cleared))
        {
            Assert.IsNull(after.Cell("Plants", "A1")?.CellValue);
        }
    }

    [TestMethod]
    public void NamingATableWithoutANameRowIsNotSupported()
    {
        Assert.AreEqual(SpreadsheetWriteError.NotSupported, WriteError(Plants().ToBytes(), w => w.Sheets[0].TableName = "Garden"));
    }

    [TestMethod]
    public void ARemovedEntryClearsItsColumnAndMerges()
    {
        var book = new TestWorkbook();
        book.AddSheet("Plants").Row("A1", "Name", "Colour", "Note").Row("A2", "Fern", "Green", "shade").Row("A3", "Rose", "Red", "sun")
            .Row("A4", "Lily", "White").Merge("C3:C4").Set("G2", "beside");
        var original = book.ToBytes();

        var bytes = Change(original, w =>
        {
            var sheet = w.Sheets[0];
            sheet.Columns.RemoveAt(1);
            foreach (var card in sheet.Cards)
            {
                card.Entries.RemoveAt(1);
            }
        });

        using var after = new XlsxPackage(bytes);
        // The note moved into the colour's column; the last column is empty and its merge is gone.
        Assert.AreEqual("Note", after.Text("Plants", "B1"));
        Assert.AreEqual("shade", after.Text("Plants", "B2"));
        foreach (var reference in new[] { "C1", "C2", "C3" })
        {
            Assert.IsNull(after.Cell("Plants", reference)?.CellValue, reference);
        }

        Assert.IsNull(after.Worksheet("Plants").GetFirstChild<X.MergeCells>());
        Assert.AreEqual("beside", after.Text("Plants", "G2"));
        var back = Read(bytes).Sheets[0];
        AssertSequence(back.Columns.Select(c => c.Name).ToArray(), "Name", "Note");
    }

    [TestMethod]
    public void ReorderedEntriesMoveTheirCellsWithTheirTypes()
    {
        var book = new TestWorkbook();
        var percent = book.AddNumberFormat(10);
        book.AddSheet("Plants").Row("A1", "Name", "Share").Set("A2", "Fern").Set("B2", 0.25, percent);

        var bytes = Change(book.ToBytes(), w =>
        {
            var sheet = w.Sheets[0];
            sheet.Columns.Reverse();
            sheet.Cards[0].Entries.Reverse();
        });

        using var after = new XlsxPackage(bytes);
        Assert.AreEqual("Share", after.Text("Plants", "A1"));
        var moved = after.Cell("Plants", "A2")!;
        Assert.IsNull(moved.DataType);
        Assert.AreEqual("0.25", moved.CellValue!.Text);
        Assert.AreEqual(percent, moved.StyleIndex!.Value);
        Assert.AreEqual("Fern", after.Text("Plants", "B2"));
        Assert.IsNull(after.Cell("Plants", "B2")!.StyleIndex);
    }

    [TestMethod]
    public void AMovedFormulaIsNotSupported()
    {
        var book = new TestWorkbook();
        book.AddSheet("Sums").Row("A1", "Name", "Value").Row("A2", "One", 1).Formula("B3", "B2+1", 2).Set("A3", "Two");

        Assert.AreEqual(SpreadsheetWriteError.NotSupported, WriteError(book.ToBytes(), w => w.Sheets[0].Cards.Reverse()));
    }

    [TestMethod]
    public void CardsOfEqualSpansCanBeReordered()
    {
        var original = Plants().ToBytes();

        var bytes = Change(original, w => w.Sheets[0].Cards.Reverse());

        var back = Read(bytes).Sheets[0];
        AssertSequence(Texts(back.Cards[0]), "Rose", "Red", "45");
        AssertSequence(Texts(back.Cards[1]), "Fern", "Green", "30");
        Assert.AreEqual(CellValueKind.Number, back.Cards[0].Entries[2].Original!.Kind);
    }

    [TestMethod]
    public void ChangesThatNeedRowsInsertedOrDeletedAreNotSupported()
    {
        var original = Plants().ToBytes();

        Assert.AreEqual(SpreadsheetWriteError.NotSupported, WriteError(original, w => w.Sheets[0].Cards.RemoveAt(0)));
        Assert.AreEqual(SpreadsheetWriteError.NotSupported, WriteError(original, w =>
        {
            var card = new Card();
            card.Entries.AddRange([new Entry(), new Entry(), new Entry()]);
            w.Sheets[0].Cards.Add(card);
        }));
        Assert.AreEqual(SpreadsheetWriteError.NotSupported, WriteError(original, w => w.Sheets[0].Cards[0].SourceRowSpan = 2));
    }

    [TestMethod]
    public void ANewCardInTheSlotOfARemovedOneIsWritten()
    {
        var original = Plants().ToBytes();

        var bytes = Change(original, w =>
        {
            var cards = w.Sheets[0].Cards;
            cards.RemoveAt(1);
            var card = new Card();
            card.Entries.AddRange([new Entry { Text = "Tulip" }, new Entry(), new Entry { Text = "20" }]);
            cards.Add(card);
        });

        AssertSequence(Texts(Read(bytes).Sheets[0].Cards[1]), "Tulip", "", "20");
    }

    [TestMethod]
    public void AMovedMultiRowCardWithLaterRowValuesIsNotSupported()
    {
        var book = new TestWorkbook();
        book.AddSheet("Plants").Row("A1", "Name", "Note").Row("A2", "Fern").Row("B3", "later").Merge("A2:A3").Row("A4", "Rose").Merge("A4:A5");
        var original = book.ToBytes();

        Assert.AreEqual(SpreadsheetWriteError.NotSupported, WriteError(original, w => w.Sheets[0].Cards.Reverse()));
    }

    [TestMethod]
    public void MovingACardWithPicturesIsNotSupported()
    {
        var book = new TestWorkbook();
        book.AddSheet("Plants").Row("A1", "Name", "Photo").Row("A2", "Fern").Row("A3", "Rose").Picture("B2", PictureBytes.Png(4, 3));

        Assert.AreEqual(SpreadsheetWriteError.NotSupported, WriteError(book.ToBytes(), w => w.Sheets[0].Cards.Reverse()));
    }

    [TestMethod]
    public void LaterRowValuesOfAMultiRowCardStay()
    {
        var book = new TestWorkbook();
        book.AddSheet("Plants").Row("A1", "Name", "Note").Row("A2", "Fern").Row("B3", "later").Merge("A2:A3");

        var bytes = Change(book.ToBytes(), w => w.Sheets[0].Cards[0].Entries[1].Text = "first");

        using var after = new XlsxPackage(bytes);
        Assert.AreEqual("first", after.Text("Plants", "B2"));
        Assert.AreEqual("later", after.Text("Plants", "B3"));
        Assert.AreEqual("A2:A3", after.Worksheet("Plants").Descendants<X.MergeCell>().Single().Reference!.Value);
    }

    [TestMethod]
    public void StylesAreWrittenWithoutChangingOtherCells()
    {
        var book = new TestWorkbook();
        var shared = book.AddNumberFormat(10);
        book.AddSheet("Plants").Row("A1", "Name", "Share").Set("A2", "Fern", shared).Set("B2", 0.5, shared).Set("A3", "Rose", shared);
        var original = book.ToBytes();

        var bytes = Change(original, w => w.Sheets[0].Cards[0].Entries[0].Style = new CellStyle
        {
            Fill = new RgbColor(0x11, 0x22, 0x33),
            TextColor = new RgbColor(0xAA, 0xBB, 0xCC),
            HorizontalAlignment = "center",
            Font = new FontStyle { Bold = true, Size = 14, Name = "Arial" },
        });

        using var before = new XlsxPackage(original);
        using var after = new XlsxPackage(bytes);
        var stylesBefore = before.WorkbookPart.WorkbookStylesPart!.Stylesheet!;
        var styles = after.WorkbookPart.WorkbookStylesPart!.Stylesheet!;
        var xfsBefore = stylesBefore.CellFormats!.Elements<X.CellFormat>().Select(x => x.OuterXml).ToList();
        var xfs = styles.CellFormats!.Elements<X.CellFormat>().ToList();
        // Existing entries are unchanged; the edited cell got a new one; the cells that shared its old one keep it.
        CollectionAssert.AreEqual(xfsBefore, xfs.Take(xfsBefore.Count).Select(x => x.OuterXml).ToList());
        var index = (int)after.Cell("Plants", "A2")!.StyleIndex!.Value;
        Assert.IsTrue(index >= xfsBefore.Count);
        Assert.AreEqual(shared, after.Cell("Plants", "B2")!.StyleIndex!.Value);
        Assert.AreEqual(shared, after.Cell("Plants", "A3")!.StyleIndex!.Value);
        var xf = xfs[index];
        Assert.AreEqual(10u, xf.NumberFormatId!.Value);
        Assert.AreEqual("center", xf.Alignment!.Horizontal!.InnerText);
        var font = styles.Fonts!.Elements<X.Font>().ElementAt((int)xf.FontId!.Value);
        Assert.IsNotNull(font.GetFirstChild<X.Bold>());
        Assert.AreEqual(14.0, font.GetFirstChild<X.FontSize>()!.Val!.Value);
        Assert.AreEqual("Arial", font.GetFirstChild<X.FontName>()!.Val!.Value);
        Assert.AreEqual("FFAABBCC", font.GetFirstChild<X.Color>()!.Rgb!.Value);
        var fill = styles.Fills!.Elements<X.Fill>().ElementAt((int)xf.FillId!.Value);
        Assert.AreEqual("FF112233", fill.PatternFill!.ForegroundColor!.Rgb!.Value);
    }

    [TestMethod]
    public void StylesAsReadChangeNothing()
    {
        var book = new TestWorkbook();
        var red = book.AddStyle(
            font: new X.Font(new X.Bold(), new X.Color { Rgb = "FFFF0000" }, new X.FontName { Val = "Calibri" }),
            fill: TestWorkbook.Solid(new X.ForegroundColor { Theme = 4 }),
            alignment: new X.Alignment { Horizontal = X.HorizontalAlignmentValues.Center });
        var rowStyle = book.AddNumberFormat(10);
        book.AddSheet("Plants").Row("A1", "Name", "Colour").Set("A2", "Fern", red).Set("B2", "Green").Row("A3", "Rose").RowStyle(3, rowStyle);
        var original = book.ToBytes();

        var bytes = Change(original, w => w.Sheets[0].Cards[1].Entries[0].Text = "Tulip");

        using var before = new XlsxPackage(original);
        using var after = new XlsxPackage(bytes);
        AssertSameXml(before.WorkbookPart.WorkbookStylesPart!, after.WorkbookPart.WorkbookStylesPart!);
        Assert.AreEqual(before.Cell("Plants", "A2")!.OuterXml, after.Cell("Plants", "A2")!.OuterXml);
        // The empty cell B3 still has no element; A3 still has no style of its own and shows the row's, as it did when read.
        Assert.IsNull(after.Cell("Plants", "B3"));
        Assert.IsNull(after.Cell("Plants", "A3")!.StyleIndex);
        Assert.AreEqual(rowStyle, after.Cell("Plants", "A3")!.Parent is X.Row row ? row.StyleIndex!.Value : 0u);
        Assert.AreEqual("Tulip", after.Text("Plants", "A3"));
    }

    [TestMethod]
    public void AValueInAStyledRowGetsTheRowsStyle()
    {
        var book = new TestWorkbook();
        var rowStyle = book.AddNumberFormat(10);
        book.AddSheet("Plants").Row("A1", "Name", "Share").Row("A2", "Fern").RowStyle(2, rowStyle);

        var bytes = Change(book.ToBytes(), w => w.Sheets[0].Cards[0].Entries[1].Text = "0.5");

        using var after = new XlsxPackage(bytes);
        Assert.AreEqual(rowStyle, after.Cell("Plants", "B2")!.StyleIndex!.Value);
    }

    [TestMethod]
    public void AutomaticTextColourIsAFontWithoutColour()
    {
        var book = new TestWorkbook();
        var red = book.AddStyle(font: new X.Font(new X.Bold(), new X.Color { Rgb = "FFFF0000" }, new X.FontName { Val = "Calibri" }));
        book.AddSheet("Plants").Row("A1", "Name", "Colour").Set("A2", "Fern", red).Set("B2", "Green");

        var bytes = Change(book.ToBytes(), w =>
        {
            var entry = w.Sheets[0].Cards[0].Entries[0];
            Assert.AreEqual(new RgbColor(0xFF, 0, 0), entry.Style!.TextColor);
            entry.Style = entry.Style with { TextColor = null, AutomaticTextColor = true };
        });

        using var after = new XlsxPackage(bytes);
        var styles = after.WorkbookPart.WorkbookStylesPart!.Stylesheet!;
        var xf = styles.CellFormats!.Elements<X.CellFormat>().ElementAt((int)after.Cell("Plants", "A2")!.StyleIndex!.Value);
        var font = styles.Fonts!.Elements<X.Font>().ElementAt((int)xf.FontId!.Value);
        Assert.IsNull(font.GetFirstChild<X.Color>());
        Assert.IsNotNull(font.GetFirstChild<X.Bold>());
        var back = Read(bytes).Sheets[0].Cards[0].Entries[0].Style!;
        Assert.IsTrue(back.AutomaticTextColor);
        Assert.IsTrue(back.Font!.Bold);
    }

    [TestMethod]
    public void MultiLineTextIsWrappedAndTopAligned()
    {
        var bytes = Change(Plants().ToBytes(), w => w.Sheets[0].Cards[0].Entries[1].Text = "Green\nand gold");

        using var after = new XlsxPackage(bytes);
        var xf = after.WorkbookPart.WorkbookStylesPart!.Stylesheet!.CellFormats!.Elements<X.CellFormat>()
            .ElementAt((int)after.Cell("Plants", "B2")!.StyleIndex!.Value);
        Assert.IsTrue(xf.Alignment!.WrapText!.Value);
        Assert.AreEqual("top", xf.Alignment.Vertical!.InnerText);
    }

    [TestMethod]
    public void ARenamedSheetIsRenamedInPlace()
    {
        var book = Plants();
        book.AddSheet("Tools").Row("A1", "Tool", "Use").Row("A2", "Spade", "dig");
        var original = book.ToBytes();

        var bytes = Change(original, w => w.Sheets[0].Name = "Garden");

        using var before = new XlsxPackage(original);
        using var after = new XlsxPackage(bytes);
        AssertSequence(after.SheetNames, "Garden", "Tools");
        Assert.AreEqual(before.Entry("Plants").SheetId!.Value, after.Entry("Garden").SheetId!.Value);
        Assert.AreEqual(before.Part("Plants").Uri, after.Part("Garden").Uri);
    }

    [TestMethod]
    public void RenamingASheetThatFormulasReferToIsNotSupported()
    {
        var book = Plants();
        book.AddSheet("Sums").Row("A1", "Name", "Total").Row("A2", "All").Formula("B2", "SUM(Plants!C2:C3)", 75);
        var original = book.ToBytes();
        var named = Plants();
        named.AddDefinedName("Heights", "'Plants'!$C$2:$C$3");

        Assert.AreEqual(SpreadsheetWriteError.NotSupported, WriteError(original, w => w.Sheets[0].Name = "Garden"));
        Assert.AreEqual(SpreadsheetWriteError.NotSupported, WriteError(named.ToBytes(), w => w.Sheets[0].Name = "Garden"));
        // The other sheet's formula does not refer to itself: renaming that one is fine.
        Change(original, w => w.Sheets[1].Name = "Totals");
    }

    [TestMethod]
    public void ARemovedSheetIsDeletedWithItsNames()
    {
        var book = Plants();
        book.AddSheet("Tools").Row("A1", "Tool", "Use").Row("A2", "Spade", "dig");
        book.AddSheet("Notes").Row("A1", "Note", "By").Row("A2", "Water", "me");
        book.AddDefinedName("_xlnm.Print_Area", "Tools!$A$1:$B$2", 1);
        book.AddDefinedName("_xlnm.Print_Area", "Notes!$A$1:$B$2", 2);
        book.AddDefinedName("ToolList", "Tools!$A$2:$A$2");
        var original = book.ToBytes();

        var bytes = Change(original, w => w.Sheets.RemoveAt(1));

        using var after = new XlsxPackage(bytes);
        AssertSequence(after.SheetNames, "Plants", "Notes");
        Assert.AreEqual(2, after.WorkbookPart.WorksheetParts.Count());
        var name = after.Workbook.DefinedNames!.Elements<X.DefinedName>().Single();
        Assert.AreEqual("Notes!$A$1:$B$2", name.Text);
        Assert.AreEqual(1u, name.LocalSheetId!.Value);
    }

    [TestMethod]
    public void RemovingASheetThatOtherSheetsReferToIsNotSupported()
    {
        var book = Plants();
        book.AddSheet("Sums").Row("A1", "Name", "Total").Row("A2", "All").Formula("B2", "SUM(Plants!C2:C3)", 75);

        Assert.AreEqual(SpreadsheetWriteError.NotSupported, WriteError(book.ToBytes(), w => w.Sheets.RemoveAt(0)));
    }

    [TestMethod]
    public void SheetsTheAppDoesNotShowKeepTheirPlaces()
    {
        var book = new TestWorkbook();
        book.AddSheet("A").Row("A1", "Name", "Kind").Row("A2", "x", "y");
        book.AddSheet("Blank");
        book.AddSheet("B").Row("A1", "Name", "Kind").Row("A2", "x", "y");
        book.AddChartSheet("Chart");
        book.AddSheet("C").Row("A1", "Name", "Kind").Row("A2", "x", "y");
        book.AddDefinedName("_xlnm.Print_Area", "C!$A$1:$B$2", 4);
        var original = book.ToBytes();

        var bytes = Change(original, w => w.Sheets.Reverse());

        using var before = new XlsxPackage(original);
        using var after = new XlsxPackage(bytes);
        AssertSequence(after.SheetNames, "C", "Blank", "B", "Chart", "A");
        Assert.AreEqual(0u, after.Workbook.DefinedNames!.Elements<X.DefinedName>().Single().LocalSheetId!.Value);
        AssertSameXml(before.Part("Blank"), after.Part("Blank"));
        AssertSequence(Read(bytes).Sheets.Select(s => s.Name).ToArray(), "C", "B", "A");
    }

    [TestMethod]
    public void ANewSheetIsAppended()
    {
        var bytes = Change(Plants().ToBytes(), w => w.Sheets.Add(NewSheet("Tools")));

        using var after = new XlsxPackage(bytes);
        AssertSequence(after.SheetNames, "Plants", "Tools");
        Assert.AreEqual(2u, after.Entry("Tools").SheetId!.Value);
        var back = Read(bytes).Sheets[1];
        Assert.AreEqual("Garden tools", back.TableName);
        AssertSequence(back.Columns.Select(c => c.Name).ToArray(), "Tool", "Count");
        AssertSequence(Texts(back.Cards[0]), "Spade", "2");
        Assert.AreEqual(CellValueKind.Number, back.Cards[0].Entries[1].Original!.Kind);
    }

    [TestMethod]
    public void ANameUsedByAWorksheetTheAppDoesNotShowIsInvalid()
    {
        var book = Plants();
        book.AddSheet("Blank");

        Assert.AreEqual(SpreadsheetWriteError.InvalidWorkbook, WriteError(book.ToBytes(), w => w.Sheets.Add(NewSheet("blank"))));
    }

    [TestMethod]
    public void ADuplicatedSheetCopiesTheWholeWorksheet()
    {
        var book = new TestWorkbook();
        var sheet = book.AddSheet("Plants").Row("A1", "Name", "Colour").Row("A2", "Fern", "Green").Set("F1", "beside").Width(1, 2, 25)
            .Picture("P40", PictureBytes.Png(4, 3), PictureAnchor.OneCell);
        sheet.PageSetup = true;
        book.AddDefinedName("_xlnm.Print_Area", "Plants!$A$1:$B$2", 0);
        var original = book.ToBytes();

        var bytes = Change(original, w =>
        {
            var source = w.Sheets[0];
            var copy = new Sheet("Plants (2)") { Source = source.Source! with { IsCopy = true }, FirstDataRow = source.FirstDataRow };
            copy.Columns.AddRange(source.Columns.Select(c => new Column(c.Name) { SourceIndex = c.SourceIndex }));
            var card = new Card { SourceFirstRow = 2 };
            card.Entries.AddRange([new Entry { Text = "Fern", Original = source.Cards[0].Entries[0].Original }, new Entry { Text = "Blue" }]);
            copy.Cards.Add(card);
            w.Sheets.Add(copy);
        });

        using var before = new XlsxPackage(original);
        using var after = new XlsxPackage(bytes);
        AssertSequence(after.SheetNames, "Plants", "Plants (2)");
        Assert.AreEqual("Green", after.Text("Plants", "B2"));
        Assert.AreEqual("Blue", after.Text("Plants (2)", "B2"));
        Assert.AreEqual("beside", after.Text("Plants (2)", "F1"));
        var copy = after.Worksheet("Plants (2)");
        Assert.AreEqual(before.Worksheet("Plants").GetFirstChild<X.PageSetup>()!.OuterXml, copy.GetFirstChild<X.PageSetup>()!.OuterXml);
        Assert.AreEqual(25.0, copy.Descendants<X.Column>().First().Width!.Value);
        var drawing = after.Part("Plants (2)").DrawingsPart!;
        Assert.AreNotEqual(after.Part("Plants").DrawingsPart!.Uri, drawing.Uri);
        Assert.AreEqual(after.Part("Plants").DrawingsPart!.ImageParts.Single().Uri, drawing.ImageParts.Single().Uri);
        var names = after.Workbook.DefinedNames!.Elements<X.DefinedName>().ToList();
        Assert.AreEqual(2, names.Count);
        Assert.AreEqual("'Plants (2)'!$A$1:$B$2", names[1].Text);
        Assert.AreEqual(1u, names[1].LocalSheetId!.Value);
    }

    [TestMethod]
    public void ANewWorkbookIsWrittenFromAnEmptyPackage()
    {
        var bytes = Write(NewWorkbook());

        var back = Read(bytes);
        Assert.AreEqual(1, back.Sheets.Count);
        AssertSequence(Texts(back.Sheets[0].Cards[0]), "Spade", "2");
        using var after = new XlsxPackage(bytes);
        AssertSequence(after.SheetNames, "Tools");
    }

    [TestMethod]
    public void AWorkbookWithoutSheetsIsInvalid()
    {
        var error = Assert.ThrowsExactly<SpreadsheetWriteException>(() => Write(new Workbook()));

        Assert.AreEqual(SpreadsheetWriteError.InvalidWorkbook, error.Reason);
    }

    [TestMethod]
    public void WritesToAPathAndLeavesNoTemporaryFile()
    {
        var folder = Directory.CreateTempSubdirectory("writer-tests-");
        try
        {
            var path = Path.Combine(folder.FullName, "plants.xlsx");
            var workbook = Read(Plants().ToBytes());
            workbook.Sheets[0].Cards[0].Entries[0].Text = "Moss";

            XlsxWriter.Write(workbook, path, Invariant);

            Assert.AreEqual("Moss", XlsxReader.Read(path, Invariant).Sheets[0].Cards[0].Entries[0].Text);
            AssertSequence(folder.GetFiles().Select(f => f.Name).ToArray(), "plants.xlsx");
        }
        finally
        {
            folder.Delete(recursive: true);
        }
    }

    [TestMethod]
    public void AFailedWriteLeavesTheTargetUnchanged()
    {
        var folder = Directory.CreateTempSubdirectory("writer-tests-");
        try
        {
            var path = Path.Combine(folder.FullName, "plants.xlsx");
            var original = Plants().ToBytes();
            File.WriteAllBytes(path, original);
            var workbook = Read(original);
            workbook.Sheets[0].Cards.RemoveAt(0);

            Assert.ThrowsExactly<SpreadsheetWriteException>(() => XlsxWriter.Write(workbook, path, Invariant));

            CollectionAssert.AreEqual(original, File.ReadAllBytes(path));
            Assert.AreEqual(1, folder.GetFiles().Length);
        }
        finally
        {
            folder.Delete(recursive: true);
        }
    }

    [TestMethod]
    public async Task WriteAsyncWritesTheFile()
    {
        var folder = Directory.CreateTempSubdirectory("writer-tests-");
        try
        {
            var path = Path.Combine(folder.FullName, "new.xlsx");

            await XlsxWriter.WriteAsync(NewWorkbook(), path, Invariant);

            Assert.AreEqual("Tools", XlsxReader.Read(path, Invariant).Sheets[0].Name);
        }
        finally
        {
            folder.Delete(recursive: true);
        }
    }

    private static Workbook NewWorkbook()
    {
        var workbook = new Workbook();
        workbook.Sheets.Add(NewSheet("Tools"));
        return workbook;
    }

    private static Sheet NewSheet(string name)
    {
        var sheet = new Sheet(name) { TableName = "Garden tools" };
        sheet.Columns.AddRange([new Column("Tool"), new Column("Count")]);
        var card = new Card();
        card.Entries.AddRange([new Entry { Text = "Spade" }, new Entry { Text = "2" }]);
        sheet.Cards.Add(card);
        return sheet;
    }
}
