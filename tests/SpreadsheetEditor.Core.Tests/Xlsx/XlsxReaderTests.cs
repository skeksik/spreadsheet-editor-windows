using System.Globalization;
using System.IO.Compression;
using SpreadsheetEditor.Core.Model;
using SpreadsheetEditor.Core.Xlsx;

namespace SpreadsheetEditor.Core.Tests.Xlsx;

[TestClass]
public sealed class XlsxReaderTests
{
    private static readonly CultureInfo Invariant = CultureInfo.InvariantCulture;

    private static Workbook Read(TestWorkbook workbook)
    {
        using var stream = workbook.ToStream();
        return XlsxReader.Read(stream, Invariant);
    }

    private static Sheet ReadSingle(TestWorkbook workbook)
    {
        var sheets = Read(workbook).Sheets;
        Assert.AreEqual(1, sheets.Count);
        return sheets[0];
    }

    private static readonly string[] ExpectedSheetNames = ["First", "Last"];
    private static readonly string[] ExpectedSkippedSheetNames = ["Empty", "Chart", "Only metadata"];
    private static readonly byte[] Garbage = [1, 2, 3, 4, 5];
    private static readonly byte[] CompoundFileHeader = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, 0, 0, 0, 0];

    private static void AssertTexts(Card card, params string[] expected) =>
        CollectionAssert.AreEqual(expected, card.Entries.Select(e => e.Text).ToArray());

    private static void AssertNames(Sheet sheet, params string[] expected) =>
        CollectionAssert.AreEqual(expected, sheet.Columns.Select(c => c.Name).ToArray());

    [TestMethod]
    public void ReadsHeaderAndCards()
    {
        var book = new TestWorkbook();
        book.AddSheet("Plants").Row("A1", "Name", "Colour").Row("A2", "Fern", "Green").Row("A3", "Rose", "Red");

        var sheet = ReadSingle(book);

        Assert.AreEqual("Plants", sheet.Name);
        AssertNames(sheet, "Name", "Colour");
        Assert.IsNull(sheet.TableName);
        Assert.AreEqual(1, sheet.TableNameColumn);
        Assert.AreEqual(1, sheet.FirstColumn);
        Assert.AreEqual(1, sheet.HeaderRow);
        Assert.AreEqual(2, sheet.FirstDataRow);
        Assert.AreEqual(2, sheet.Cards.Count);
        AssertTexts(sheet.Cards[0], "Fern", "Green");
        AssertTexts(sheet.Cards[1], "Rose", "Red");
        Assert.AreEqual(2, sheet.Cards[0].SourceFirstRow);
        Assert.AreEqual(1, sheet.Cards[0].SourceRowSpan);
        Assert.AreEqual(1u, sheet.Source!.SheetId);
        Assert.AreEqual((1, 3, 1, 2), (sheet.Source.FirstRow, sheet.Source.LastRow, sheet.Source.FirstColumn, sheet.Source.LastColumn));
    }

    [TestMethod]
    public void FindsTableAwayFromTheCorner()
    {
        var book = new TestWorkbook();
        book.AddSheet("S").Row("C4", "Name", "Size").Row("C5", "Box", 3);

        var sheet = ReadSingle(book);

        Assert.AreEqual(3, sheet.FirstColumn);
        Assert.AreEqual(3, sheet.TableNameColumn);
        Assert.AreEqual(4, sheet.HeaderRow);
        Assert.AreEqual(5, sheet.FirstDataRow);
        AssertTexts(sheet.Cards[0], "Box", "3");
    }

    [TestMethod]
    public void ReadsTableName()
    {
        var book = new TestWorkbook();
        book.AddSheet("S").Row("B1", "Inventory").Row("A2", "Name", "Count").Row("A3", "Pen", 4);

        var sheet = ReadSingle(book);

        Assert.AreEqual("Inventory", sheet.TableName);
        Assert.AreEqual(2, sheet.TableNameColumn);
        Assert.AreEqual(2, sheet.HeaderRow);
        Assert.AreEqual(3, sheet.FirstDataRow);
        AssertNames(sheet, "Name", "Count");
        Assert.AreEqual(1, sheet.Source!.FirstRow);
    }

    [TestMethod]
    public void WhiteSpaceCellIsBlankInTheNameTest()
    {
        var book = new TestWorkbook();
        book.AddSheet("S").Row("A1", "Title", "   ").Row("A2", "Name", "Count").Row("A3", "Pen", 4);

        Assert.AreEqual("Title", ReadSingle(book).TableName);
    }

    [TestMethod]
    public void NoNameWhenTheNextRowHasOneCell()
    {
        var book = new TestWorkbook();
        book.AddSheet("S").Row("A1", "Title").Row("A2", "Pen").Row("A3", "Ink");

        var sheet = ReadSingle(book);

        Assert.IsNull(sheet.TableName);
        Assert.AreEqual(1, sheet.HeaderRow);
        AssertNames(sheet, "Title");
        Assert.AreEqual(2, sheet.Cards.Count);
    }

    [TestMethod]
    public void NoNameWhenTheTableHasOneRow()
    {
        var book = new TestWorkbook();
        book.AddSheet("S").Row("A1", "Title");

        var sheet = ReadSingle(book);

        Assert.IsNull(sheet.TableName);
        Assert.AreEqual(1, sheet.HeaderRow);
    }

    [TestMethod]
    public void RowGapOfTenIsOneBlockAndElevenEndsIt()
    {
        var book = new TestWorkbook();
        // Rows 2 and 12 are 10 apart (9 empty rows between): one block. Row 23 is 11 below row 12: outside.
        book.AddSheet("S").Row("A1", "Name", "Value").Row("A2", "a", 1).Row("A12", "b", 2).Row("A23", "c", 3);

        var sheet = ReadSingle(book);

        Assert.AreEqual(11, sheet.Cards.Count);
        AssertTexts(sheet.Cards[10], "b", "2");
        Assert.AreEqual(12, sheet.Source!.LastRow);
    }

    [TestMethod]
    public void ColumnGapOfThreeIsOneBlockAndFourEndsIt()
    {
        var book = new TestWorkbook();
        // A and D are 3 apart (2 empty columns): one block. H is 4 right of D: outside.
        book.AddSheet("S").Row("A1", "One", null, null, "Four", null, null, null, "Eight").Row("A2", "x", null, null, "y", null, null, null, "z");

        var sheet = ReadSingle(book);

        AssertNames(sheet, "One", "", "", "Four");
        AssertTexts(sheet.Cards[0], "x", "", "", "y");
    }

    [TestMethod]
    public void IgnoresLastColumnAndMetadataStrings()
    {
        var book = new TestWorkbook();
        // A saved file's metadata in XFD1:XFD3 must not pull the table up to row 1 (§9.1); a stray metadata string elsewhere neither.
        book.AddSheet("S")
            .Set("XFD1", "ImageHeight:200").Set("XFD2", "ExcelImagesPerRow:4").Set("XFD3", "ImageWidth:150")
            .Set("XFD20", "anything")
            .Set("J2", "ImageWidth:300")
            .Row("B14", "Name", "Kind").Row("B15", "Oak", "Tree");

        var sheet = ReadSingle(book);

        Assert.AreEqual(14, sheet.HeaderRow);
        Assert.AreEqual(2, sheet.FirstColumn);
        AssertNames(sheet, "Name", "Kind");
        Assert.AreEqual(new ImageLayout(150, 200, 4), sheet.ImageLayout);
    }

    [TestMethod]
    public void TextThatOnlyLooksLikeMetadataIsOccupied()
    {
        var book = new TestWorkbook();
        book.AddSheet("S").Row("A1", "ImageWidth: wide", "Other").Row("A2", "x", "y");

        AssertNames(ReadSingle(book), "ImageWidth: wide", "Other");
    }

    [TestMethod]
    public void HeadersAreHiddenWhenEveryHeaderCellIsANumber()
    {
        var book = new TestWorkbook();
        book.AddSheet("S").Row("A1", 1, 2.5, 3).Row("A2", 4, 5, 6);

        var sheet = ReadSingle(book);

        Assert.IsNull(sheet.HeaderRow);
        Assert.AreEqual(1, sheet.FirstDataRow);
        AssertNames(sheet, "", "", "");
        Assert.AreEqual(2, sheet.Cards.Count);
        AssertTexts(sheet.Cards[0], "1", "2.5", "3");
        Assert.AreEqual(1, sheet.Cards[0].SourceFirstRow);
    }

    [TestMethod]
    public void HeaderWithTextAndNumbersIsShown()
    {
        var book = new TestWorkbook();
        book.AddSheet("S").Row("A1", "Name", 2024, "Price").Row("A2", "Tea", 1, 2);

        var sheet = ReadSingle(book);

        Assert.AreEqual(1, sheet.HeaderRow);
        AssertNames(sheet, "Name", "2024", "Price");
    }

    [TestMethod]
    public void BooleanHeaderIsNotANumber()
    {
        var book = new TestWorkbook();
        book.AddSheet("S").Row("A1", true, 2).Row("A2", 1, 2);

        Assert.AreEqual(1, ReadSingle(book).HeaderRow);
    }

    [TestMethod]
    public void DateHeaderIsNotANumber()
    {
        var book = new TestWorkbook();
        var date = book.AddNumberFormat(14);
        book.AddSheet("S").Set("A1", 45000, date).Set("B1", 7).Row("A2", 1, 2);

        Assert.AreEqual(1, ReadSingle(book).HeaderRow);
    }

    [TestMethod]
    public void FormulaHeaderCountsByItsCachedValue()
    {
        var book = new TestWorkbook();
        book.AddSheet("S").Formula("A1", "1+1", 2).Set("B1", 3).Row("A2", 1, 2);

        Assert.IsNull(ReadSingle(book).HeaderRow);

        var text = new TestWorkbook();
        text.AddSheet("S").Formula("A1", "\"a\"&\"b\"", "ab").Set("B1", 3).Row("A2", 1, 2);

        Assert.AreEqual(1, ReadSingle(text).HeaderRow);
    }

    [TestMethod]
    public void ReadsValueKinds()
    {
        var book = new TestWorkbook();
        book.AddSheet("S")
            .Row("A1", "Text", "Number", "Flag", "Inline", "Rich", "Error", "Formula", "Text formula")
            .Set("A2", "plain").Set("B2", 3.5).Set("C2", false).Inline("D2", "inline text").RichText("E2", "Bo", "ld")
            .Error("F2", "#N/A").Formula("G2", "SUM(B2:B2)", 3.5).Formula("H2", "UPPER(A2)", "PLAIN");

        var entries = ReadSingle(book).Cards[0].Entries;

        Assert.AreEqual(new CellValue(CellValueKind.Text, "plain", Format: NumberFormat.General), entries[0].Original);
        Assert.AreEqual(new CellValue(CellValueKind.Number, "3.5", Number: 3.5, Format: NumberFormat.General), entries[1].Original);
        Assert.AreEqual(new CellValue(CellValueKind.Boolean, "FALSE", Boolean: false, Format: NumberFormat.General), entries[2].Original);
        Assert.AreEqual("inline text", entries[3].Text);
        Assert.AreEqual("Bold", entries[4].Text);
        Assert.AreEqual(new CellValue(CellValueKind.Error, "#N/A", Format: NumberFormat.General), entries[5].Original);
        Assert.AreEqual("=SUM(B2:B2)", entries[6].Text);
        Assert.AreEqual(
            new CellValue(CellValueKind.Number, "=SUM(B2:B2)", Number: 3.5, Formula: "SUM(B2:B2)", Format: NumberFormat.General),
            entries[6].Original);
        Assert.AreEqual("=UPPER(A2)", entries[7].Text);
        Assert.AreEqual(CellValueKind.Text, entries[7].Original!.Kind);
        Assert.AreEqual("UPPER(A2)", entries[7].Original!.Formula);
    }

    [TestMethod]
    public void KeepsTheNumberFormatOnEveryValue()
    {
        var book = new TestWorkbook();
        var text = book.AddNumberFormat(49);
        var custom = book.AddNumberFormat(170, "0.000");
        book.AddSheet("S").Row("A1", "Code", "Weight").Set("A2", "007", text).Set("B2", 1.5, custom);

        var entries = ReadSingle(book).Cards[0].Entries;

        Assert.AreEqual(new NumberFormat(49), entries[0].Original!.Format);
        Assert.AreEqual("007", entries[0].Text);
        Assert.AreEqual(new NumberFormat(170, "0.000"), entries[1].Original!.Format);
        // The number format is kept, not applied to the text (CellText.Format).
        Assert.AreEqual("1.5", entries[1].Text);
    }

    [TestMethod]
    public void FormulaWithoutCachedValueIsShownAsItsSource()
    {
        var book = new TestWorkbook();
        book.AddSheet("S").Row("A1", "A").Formula("A2", "NOW()", "");

        var original = ReadSingle(book).Cards[0].Entries[0].Original!;

        Assert.AreEqual("NOW()", original.Formula);
        Assert.AreEqual("=NOW()", original.DisplayText);
    }

    [TestMethod]
    public void EmptyCellsGiveEmptyEntries()
    {
        var book = new TestWorkbook();
        book.AddSheet("S").Row("A1", "A", "B").Set("A2", "x").Empty("B2");

        var entry = ReadSingle(book).Cards[0].Entries[1];

        Assert.AreEqual("", entry.Text);
        Assert.IsNull(entry.Original);
    }

    [TestMethod]
    public void ReadsDatesInBothDateSystems()
    {
        foreach (var date1904 in new[] { false, true })
        {
            var book = new TestWorkbook { Date1904 = date1904 };
            var builtIn = book.AddNumberFormat(14);
            var custom = book.AddNumberFormat(170, "dd\\.mm\\.yyyy");
            var notDate = book.AddNumberFormat(171, "0.00\" days\"");
            // 2024-01-15 is serial 45306 in the 1900 system and 43844 in the 1904 system.
            var serial = date1904 ? 43844 : 45306;
            book.AddSheet("S").Row("A1", "Built in", "Custom", "Number")
                .Set("A2", serial, builtIn).Set("B2", serial + 0.5, custom).Set("C2", serial, notDate);

            var entries = ReadSingle(book).Cards[0].Entries;

            Assert.AreEqual(CellValueKind.Date, entries[0].Original!.Kind);
            Assert.AreEqual(new DateTime(2024, 1, 15), entries[0].Original!.Date);
            Assert.AreEqual((double)serial, entries[0].Original!.Number);
            Assert.AreEqual(new DateTime(2024, 1, 15, 12, 0, 0), entries[1].Original!.Date);
            Assert.AreEqual(CellValueKind.Number, entries[2].Original!.Kind);
        }
    }

    [TestMethod]
    public void MergesInsideTheTableGroupRowsIntoCards()
    {
        var book = new TestWorkbook();
        book.AddSheet("S")
            .Row("A1", "Name", "Note")
            .Row("A2", "Fern", "first").Row("A3", null, "second row, not read")
            .Row("A4", "Rose", "x")
            .Merge("A2:A3");

        var sheet = ReadSingle(book);

        Assert.AreEqual(2, sheet.Cards.Count);
        Assert.AreEqual(2, sheet.Cards[0].SourceRowSpan);
        Assert.AreEqual(2, sheet.Cards[0].SourceFirstRow);
        AssertTexts(sheet.Cards[0], "Fern", "first");
        Assert.AreEqual(4, sheet.Cards[1].SourceFirstRow);
        AssertTexts(sheet.Cards[1], "Rose", "x");
    }

    [TestMethod]
    public void ChainedMergesExtendACard()
    {
        var book = new TestWorkbook();
        book.AddSheet("S").Row("A1", "A", "B").Row("A2", "a", "b").Row("A5", "c", "d").Merge("A2:A3").Merge("B3:B4");

        var sheet = ReadSingle(book);

        Assert.AreEqual(3, sheet.Cards[0].SourceRowSpan);
        Assert.AreEqual(5, sheet.Cards[1].SourceFirstRow);
    }

    [TestMethod]
    public void MergesOutsideTheTableColumnsDoNotGroup()
    {
        var book = new TestWorkbook();
        // The table is A:B. D2:D3 lies outside it and B2:C3 straddles its edge: neither groups rows (§9.7; a straddling range is not
        // wholly inside the table's columns).
        book.AddSheet("S").Row("A1", "A", "B").Row("A2", "a", "b").Row("A3", "c", "d").Merge("D2:D3").Merge("B2:C3");

        var sheet = ReadSingle(book);

        Assert.AreEqual(2, sheet.Cards.Count);
        Assert.IsTrue(sheet.Cards.TrueForAll(c => c.SourceRowSpan == 1));
    }

    [TestMethod]
    public void EmptyRowsInsideTheTableAreEmptyCards()
    {
        var book = new TestWorkbook();
        book.AddSheet("S").Row("A1", "A", "B").Row("A2", "a", "b").Row("A4", "c", "d");

        var sheet = ReadSingle(book);

        Assert.AreEqual(3, sheet.Cards.Count);
        AssertTexts(sheet.Cards[1], "", "");
        Assert.AreEqual(3, sheet.Cards[1].SourceFirstRow);
    }

    [TestMethod]
    public void HeaderOnlyTableHasOneEmptyCard()
    {
        var book = new TestWorkbook();
        book.AddSheet("S").Row("A1", "A", "B");

        var sheet = ReadSingle(book);

        Assert.AreEqual(1, sheet.Cards.Count);
        Assert.IsNull(sheet.Cards[0].SourceFirstRow);
        AssertTexts(sheet.Cards[0], "", "");
    }

    [TestMethod]
    public void DefaultLayoutWithoutMetadata()
    {
        var book = new TestWorkbook();
        book.AddSheet("S").Row("A1", "A");

        Assert.AreEqual(ImageLayout.Default, ReadSingle(book).ImageLayout);
    }

    [TestMethod]
    public void SkipsEmptyAndChartSheetsAndKeepsOrder()
    {
        var book = new TestWorkbook();
        book.AddSheet("First").Row("A1", "A");
        book.AddSheet("Empty").Empty("A1");
        book.AddChartSheet("Chart");
        book.AddSheet("Only metadata").Set("XFD1", "ImageHeight:100");
        book.AddSheet("Last").Row("B2", "B");

        var workbook = Read(book);

        CollectionAssert.AreEqual(ExpectedSheetNames, workbook.Sheets.Select(s => s.Name).ToArray());
        Assert.AreEqual(5u, workbook.Sheets[1].Source!.SheetId);
        Assert.AreNotEqual(workbook.Sheets[0].Source!.RelationshipId, workbook.Sheets[1].Source!.RelationshipId);
        CollectionAssert.AreEqual(ExpectedSkippedSheetNames, workbook.SkippedSheetNames);
    }

    [TestMethod]
    public void ReadsHiddenWorksheetsLikeOthers()
    {
        var book = new TestWorkbook();
        book.AddSheet("Shown").Row("A1", "A").Row("A2", "x");
        book.AddSheet("Hidden", hidden: true).Row("A1", "B").Row("A2", "y");

        var workbook = Read(book);

        Assert.AreEqual(2, workbook.Sheets.Count);
        Assert.AreEqual("Hidden", workbook.Sheets[1].Name);
        AssertTexts(workbook.Sheets[1].Cards[0], "y");
        Assert.AreEqual(0, workbook.SkippedSheetNames.Count, "hidden worksheets are shown, not skipped");
    }

    [TestMethod]
    public void WorkbookWithOnlyAChartSheetHasNoSheets()
    {
        var book = new TestWorkbook();
        book.AddChartSheet("Chart");

        var workbook = Read(book);

        Assert.AreEqual(0, workbook.Sheets.Count);
        Assert.IsNotNull(workbook.Source);
    }

    [TestMethod]
    public void KeepsThePackageBytes()
    {
        var book = new TestWorkbook();
        book.AddSheet("S").Row("A1", "A");
        var bytes = book.ToBytes();

        using var stream = new MemoryStream(bytes);
        var workbook = XlsxReader.Read(stream, Invariant);

        CollectionAssert.AreEqual(bytes, workbook.Source!.Package.ToArray());
        Assert.IsNull(workbook.SourcePath);
    }

    [TestMethod]
    public async Task ReadsFromAPathOffTheCallingThread()
    {
        var book = new TestWorkbook();
        book.AddSheet("S").Row("A1", "A").Row("A2", "x");
        var path = Path.Combine(Path.GetTempPath(), $"xlsx-reader-test-{Guid.NewGuid():N}.xlsx");
        await File.WriteAllBytesAsync(path, book.ToBytes());
        try
        {
            var workbook = await XlsxReader.ReadAsync(path, Invariant);

            Assert.AreEqual(path, workbook.SourcePath);
            Assert.AreEqual("x", workbook.Sheets[0].Cards[0].Entries[0].Text);
            // The file is not kept open: it can be deleted (or saved over) right away.
            File.Delete(path);
        }
        finally
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
    }

    [TestMethod]
    public void MissingFileIsNotFound()
    {
        var path = Path.Combine(Path.GetTempPath(), $"xlsx-reader-missing-{Guid.NewGuid():N}.xlsx");

        var error = Assert.ThrowsExactly<SpreadsheetReadException>(() => XlsxReader.Read(path, Invariant));

        Assert.AreEqual(SpreadsheetReadError.NotFound, error.Reason);
    }

    [TestMethod]
    public void GarbageIsNotAnXlsx()
    {
        Assert.AreEqual(SpreadsheetReadError.NotAnXlsx, ReadError(Garbage));
        Assert.AreEqual(SpreadsheetReadError.NotAnXlsx, ReadError(Array.Empty<byte>()));
    }

    [TestMethod]
    public void CompoundFileIsLegacyOrEncrypted() =>
        Assert.AreEqual(SpreadsheetReadError.LegacyOrEncrypted, ReadError(CompoundFileHeader));

    [TestMethod]
    public void ZipWithoutAWorkbookIsCorrupt()
    {
        using var zip = new MemoryStream();
        using (var archive = new ZipArchive(zip, ZipArchiveMode.Create, leaveOpen: true))
        {
            using var writer = new StreamWriter(archive.CreateEntry("hello.txt").Open());
            writer.Write("not a workbook");
        }

        Assert.AreEqual(SpreadsheetReadError.Corrupt, ReadError(zip.ToArray()));
    }

    [TestMethod]
    public void TruncatedPackageIsCorrupt()
    {
        var book = new TestWorkbook();
        book.AddSheet("S").Row("A1", "A");
        var bytes = book.ToBytes();

        Assert.AreEqual(SpreadsheetReadError.Corrupt, ReadError(bytes[..(bytes.Length / 2)]));
    }

    [TestMethod]
    public void CancelledReadThrows()
    {
        var book = new TestWorkbook();
        book.AddSheet("S").Row("A1", "A");
        using var stream = book.ToStream();

        Assert.ThrowsExactly<OperationCanceledException>(() => XlsxReader.Read(stream, Invariant, new CancellationToken(canceled: true)));
    }

    private static SpreadsheetReadError ReadError(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        return Assert.ThrowsExactly<SpreadsheetReadException>(() => XlsxReader.Read(stream, Invariant)).Reason;
    }
}
