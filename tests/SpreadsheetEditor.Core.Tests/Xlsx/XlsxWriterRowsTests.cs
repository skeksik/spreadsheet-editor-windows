using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using SpreadsheetEditor.Core.Model;
using SpreadsheetEditor.Core.Tests.Imaging;
using SpreadsheetEditor.Core.Xlsx;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using X = DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace SpreadsheetEditor.Core.Tests.Xlsx;

/// <summary>
/// Round trips of saves that insert, delete or move worksheet rows (docs/spec.md §5.3, §9.18, §9.20, §9.22): build a workbook, read it,
/// change the model, write, read back and compare. Every value in them is made up.
/// </summary>
[TestClass]
public sealed class XlsxWriterRowsTests
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

    /// <summary>Reads, changes the model, writes twice and checks that both saves are the same file.</summary>
    private static byte[] Change(byte[] original, Action<Workbook> change)
    {
        var workbook = Read(original);
        change(workbook);
        var first = Write(workbook);
        CollectionAssert.AreEqual(first, Write(workbook), "Two saves differ.");
        return first;
    }

    private static SpreadsheetWriteError WriteError(byte[] original, Action<Workbook> change)
    {
        var workbook = Read(original);
        change(workbook);
        return Assert.ThrowsExactly<SpreadsheetWriteException>(() => Write(workbook)).Reason;
    }

    /// <summary>Changes a built package with the Open XML SDK, for content the test builders do not make.</summary>
    private static byte[] Edit(byte[] bytes, Action<SpreadsheetDocument> edit)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        using (var document = SpreadsheetDocument.Open(stream, true))
        {
            edit(document);
        }

        return stream.ToArray();
    }

    private static X.Worksheet Sheet(SpreadsheetDocument document, string name)
    {
        var entry = document.WorkbookPart!.Workbook!.Sheets!.Elements<X.Sheet>().Single(s => s.Name == name);
        return ((WorksheetPart)document.WorkbookPart.GetPartById(entry.Id!.Value!)).Worksheet!;
    }

    private static byte[] Lilies()
    {
        var book = new TestWorkbook();
        book.AddSheet("Plants").Row("A1", "Name", "Kind").Row("A2", "Fern", "a").Row("A3", "Rose", "b").Row("A4", "Lily", "c");
        return book.ToBytes();
    }

    private static Card NewCard(params string[] texts)
    {
        var card = new Card();
        card.Entries.AddRange(texts.Select(t => new Entry { Text = t }));
        return card;
    }

    private static string[] Texts(Card card) => card.Entries.Select(e => e.Text).ToArray();

    private static void AssertSequence(IEnumerable<string> actual, params string[] expected) =>
        CollectionAssert.AreEqual(expected, actual.ToArray());

    private static string? Formula(XlsxPackage package, string sheet, string reference) => package.Cell(sheet, reference)?.CellFormula?.Text;

    /// <summary>A defined name's text; a print area is the first sheet's.</summary>
    private static string DefinedName(XlsxPackage package, string name) =>
        package.Workbook.DefinedNames!.Elements<X.DefinedName>()
            .Single(n => n.Name == name && (n.LocalSheetId is null || n.LocalSheetId.Value == 0)).Text;

    /// <summary>
    /// A table of two cards (rows 2 and 3) with content below it on row 15: a value, a total, a merge and a tall row; a sheet that refers
    /// to both; defined names; a hyperlink, a conditional format and a data validation below the table.
    /// </summary>
    private static byte[] Garden()
    {
        var book = new TestWorkbook();
        book.AddSheet("Plants").Row("A1", "Name", "Height").Row("A2", "Fern", 30).Row("A3", "Rose", 45)
            .Set("A15", "Total").Formula("B15", "SUM(B2:B3)", 75).Set("C15", "note").Merge("C15:D15").Formula("E15", "B15*2", 150);
        book.AddSheet("Report").Row("A1", "Item", "Value").Row("A2", "Total").Formula("B2", "Plants!B15", 75)
            .Row("A3", "Second").Formula("B3", "Plants!B3", 45);
        book.AddDefinedName("Total", "Plants!$B$15");
        book.AddDefinedName("_xlnm.Print_Area", "Plants!$A$1:$E$15", 0);
        return Edit(book.ToBytes(), document =>
        {
            var plants = Sheet(document, "Plants");
            plants.Descendants<X.Row>().Single(r => r.RowIndex! == 15u).Height = 30;
            plants.Descendants<X.Row>().Single(r => r.RowIndex! == 15u).CustomHeight = true;
            var conditional = new X.ConditionalFormatting(
                new X.ConditionalFormattingRule(new X.Formula("B15>100")) { Type = X.ConditionalFormatValues.Expression, Priority = 1 })
            { SequenceOfReferences = new ListValue<StringValue> { InnerText = "B15:E15" } };
            plants.InsertAfter(conditional, plants.GetFirstChild<X.MergeCells>());
            var validations = new X.DataValidations(
                new X.DataValidation(new X.Formula1("$A$2:$A$3")) { Type = X.DataValidationValues.List, SequenceOfReferences = new ListValue<StringValue> { InnerText = "F15" } })
            { Count = 1 };
            plants.InsertAfter(validations, conditional);
            plants.InsertAfter(new X.Hyperlinks(new X.Hyperlink { Reference = "A15", Location = "Plants!B15" }), validations);
        });
    }

    [TestMethod]
    public void AnAddedCardInsertsARowAndShiftsEverythingBelow()
    {
        var original = Garden();

        var bytes = Change(original, w => w.Sheets[0].Cards.Add(NewCard("Tulip", "20")));

        using var before = new XlsxPackage(original);
        using var after = new XlsxPackage(bytes);
        Assert.AreEqual("Tulip", after.Text("Plants", "A4"));
        Assert.AreEqual("20", after.Text("Plants", "B4"));
        // The content below moved down a row, with its row height, merge and formulas; the total now covers the new card.
        Assert.AreEqual("Total", after.Text("Plants", "A16"));
        Assert.AreEqual("SUM(B2:B3)", Formula(after, "Plants", "B16"));
        Assert.AreEqual("B16*2", Formula(after, "Plants", "E16"));
        Assert.IsNull(after.Cell("Plants", "A15"));
        Assert.AreEqual(30, after.Worksheet("Plants").Descendants<X.Row>().Single(r => r.RowIndex! == 16u).Height!.Value);
        Assert.AreEqual("C16:D16", after.Worksheet("Plants").Descendants<X.MergeCell>().Single().Reference!.Value);
        var conditional = after.Worksheet("Plants").Descendants<X.ConditionalFormatting>().Single();
        Assert.AreEqual("B16:E16", conditional.SequenceOfReferences!.InnerText);
        Assert.AreEqual("B16>100", conditional.Descendants<X.Formula>().Single().Text);
        Assert.AreEqual("F16", after.Worksheet("Plants").Descendants<X.DataValidation>().Single().SequenceOfReferences!.InnerText);
        var hyperlink = after.Worksheet("Plants").Descendants<X.Hyperlink>().Single();
        Assert.AreEqual(("A16", "Plants!B16"), (hyperlink.Reference!.Value, hyperlink.Location!.Value));
        // Other sheets and defined names follow; the print area keeps its columns and grows with the table.
        Assert.AreEqual("Plants!B16", Formula(after, "Report", "B2"));
        Assert.AreEqual("Plants!B3", Formula(after, "Report", "B3"));
        Assert.AreEqual("Plants!$B$16", DefinedName(after, "Total"));
        Assert.AreEqual("Plants!$A$1:$E$16", DefinedName(after, "_xlnm.Print_Area"));
        Assert.IsTrue(after.Workbook.CalculationProperties!.FullCalculationOnLoad!.Value);

        var back = Read(bytes).Sheets[0];
        Assert.AreEqual(3, back.Cards.Count);
        AssertSequence(Texts(back.Cards[2]), "Tulip", "20");
        // The other sheet is unchanged apart from its formulas.
        Assert.AreEqual(before.Text("Report", "A2"), after.Text("Report", "A2"));
    }

    [TestMethod]
    public void ARemovedCardDeletesItsRowsAndShiftsEverythingBelowUp()
    {
        var bytes = Change(Garden(), w => w.Sheets[0].Cards.RemoveAt(1));

        using var after = new XlsxPackage(bytes);
        Assert.AreEqual("Fern", after.Text("Plants", "A2"));
        Assert.AreEqual("Total", after.Text("Plants", "A14"));
        Assert.IsNull(after.Cell("Plants", "A15"));
        Assert.AreEqual("SUM(B2:B2)", Formula(after, "Plants", "B14"));
        Assert.AreEqual("C14:D14", after.Worksheet("Plants").Descendants<X.MergeCell>().Single().Reference!.Value);
        // A reference to the deleted card's cell has nothing left to point at.
        Assert.AreEqual("Plants!#REF!", Formula(after, "Report", "B3"));
        Assert.AreEqual("Plants!B14", Formula(after, "Report", "B2"));
        Assert.AreEqual("Plants!$A$1:$E$14", DefinedName(after, "_xlnm.Print_Area"));
        Assert.AreEqual(1, Read(bytes).Sheets[0].Cards.Count);
    }

    [TestMethod]
    public void ACardInsertedBetweenOthersGrowsTheRangesAroundIt()
    {
        var bytes = Change(Garden(), w => w.Sheets[0].Cards.Insert(1, NewCard("Ivy", "5")));

        using var after = new XlsxPackage(bytes);
        AssertSequence(Read(bytes).Sheets[0].Cards.Select(c => c.Entries[0].Text), "Fern", "Ivy", "Rose");
        Assert.AreEqual("SUM(B2:B4)", Formula(after, "Plants", "B16"));
        Assert.AreEqual("Plants!B4", Formula(after, "Report", "B3"));
        Assert.AreEqual("$A$2:$A$4", after.Worksheet("Plants").Descendants<X.Formula1>().Single().Text);
    }

    [TestMethod]
    public void ReorderedCardsMoveTheirWholeRowsAndReferencesStayByPosition()
    {
        var book = new TestWorkbook();
        book.AddSheet("Plants").Row("A1", "Name", "Note").Row("A2", "Fern", "shade").Row("B3", "later").Merge("A2:A3").Set("F2", "beside")
            .Row("A4", "Rose", "sun").Set("A15", "after").Formula("B15", "B4", "sun");
        var original = Edit(book.ToBytes(), document =>
            Sheet(document, "Plants").Descendants<X.Row>().Single(r => r.RowIndex! == 3u).Height = 40);

        var bytes = Change(original, w => w.Sheets[0].Cards.Reverse());

        using var after = new XlsxPackage(bytes);
        // Rose's row comes first; Fern's two rows follow with their later-row value, the cell beside them, merge and height.
        Assert.AreEqual("Rose", after.Text("Plants", "A2"));
        Assert.AreEqual("Fern", after.Text("Plants", "A3"));
        Assert.AreEqual("later", after.Text("Plants", "B4"));
        Assert.AreEqual("beside", after.Text("Plants", "F3"));
        Assert.AreEqual("A3:A4", after.Worksheet("Plants").Descendants<X.MergeCell>().Single().Reference!.Value);
        Assert.AreEqual(40, after.Worksheet("Plants").Descendants<X.Row>().Single(r => r.RowIndex! == 4u).Height!.Value);
        // Rows below keep their place; a reference to the table keeps its position (as after a sort).
        Assert.AreEqual("after", after.Text("Plants", "A15"));
        Assert.AreEqual("B4", Formula(after, "Plants", "B15"));
        var back = Read(bytes).Sheets[0];
        AssertSequence(back.Cards.Select(c => c.Entries[0].Text), "Rose", "Fern");
        AssertSequence(back.Cards.Select(c => c.SourceRowSpan.ToString(Invariant)), "1", "2");
    }

    [TestMethod]
    public void TheMetadataCellsStayInTheirRowsWhenCardsThereMove()
    {
        var saved = Change(Lilies(),
            w => w.Sheets[0].ImageLayout = new ImageLayout(120, 160, 4));

        var bytes = Change(saved, w =>
        {
            w.Sheets[0].Cards.Reverse();
            w.Sheets[0].TableName = "Garden";
        });

        using var after = new XlsxPackage(bytes);
        Assert.AreEqual("ImageHeight:160", after.Text("Plants", "XFD1"));
        Assert.AreEqual("ExcelImagesPerRow:4", after.Text("Plants", "XFD2"));
        Assert.AreEqual("ImageWidth:120", after.Text("Plants", "XFD3"));
        Assert.IsNull(after.Cell("Plants", "XFD4"));
        var back = Read(bytes).Sheets[0];
        Assert.AreEqual(new ImageLayout(120, 160, 4), back.ImageLayout);
        AssertSequence(back.Cards.Select(c => c.Entries[0].Text), "Lily", "Rose", "Fern");
    }

    [TestMethod]
    public void AMovedFormulaIsTranslatedByItsMove()
    {
        var book = new TestWorkbook();
        book.AddSheet("Sums").Row("A1", "Name", "Value", "Next").Row("A2", "One", 1).Formula("C2", "B2+1", 2).Row("A3", "Two", 2)
            .Formula("C3", "$B$2+B3", 3);

        var bytes = Change(book.ToBytes(), w => w.Sheets[0].Cards.Reverse());

        using var after = new XlsxPackage(bytes);
        // The cells move with their card; relative references move with them, absolute ones stay.
        Assert.AreEqual("$B$2+B2", Formula(after, "Sums", "C2"));
        Assert.AreEqual("B3+1", Formula(after, "Sums", "C3"));
        Assert.IsTrue(after.Workbook.CalculationProperties!.FullCalculationOnLoad!.Value);
    }

    [TestMethod]
    public void AFormulaInAMovedEntryIsTranslatedByItsColumn()
    {
        var book = new TestWorkbook();
        book.AddSheet("Sums").Row("A1", "Name", "Value", "Double").Row("A2", "One", 1).Formula("C2", "B2*2", 2);

        var bytes = Change(book.ToBytes(), w =>
        {
            var sheet = w.Sheets[0];
            (sheet.Columns[1], sheet.Columns[2]) = (sheet.Columns[2], sheet.Columns[1]);
            var card = sheet.Cards[0];
            (card.Entries[1], card.Entries[2]) = (card.Entries[2], card.Entries[1]);
        });

        using var after = new XlsxPackage(bytes);
        Assert.AreEqual("A2*2", Formula(after, "Sums", "B2"));
        Assert.AreEqual("1", after.Text("Sums", "C2"));
    }

    [TestMethod]
    public void NamingATableWithoutANameRowInsertsARowAboveTheHeader()
    {
        var original = Garden();

        var bytes = Change(original, w => w.Sheets[0].TableName = "Garden");

        using var after = new XlsxPackage(bytes);
        Assert.AreEqual("Garden", after.Text("Plants", "A1"));
        Assert.AreEqual("Name", after.Text("Plants", "A2"));
        Assert.AreEqual("Fern", after.Text("Plants", "A3"));
        Assert.AreEqual("SUM(B3:B4)", Formula(after, "Plants", "B16"));
        Assert.AreEqual("Plants!B16", Formula(after, "Report", "B2"));
        var back = Read(bytes).Sheets[0];
        Assert.AreEqual("Garden", back.TableName);
        AssertSequence(back.Cards.Select(c => c.Entries[0].Text), "Fern", "Rose");

        // Clearing the name later clears its cell; the row stays.
        var cleared = Change(bytes, w => w.Sheets[0].TableName = null);
        using var again = new XlsxPackage(cleared);
        Assert.IsNull(again.Cell("Plants", "A1")?.CellValue);
        Assert.AreEqual("Name", again.Text("Plants", "A2"));
    }

    [TestMethod]
    public void ANewCardOfSeveralRowsIsMergedOverThem()
    {
        var bytes = Change(Garden(), w =>
        {
            var card = NewCard("Ivy", "5");
            card.SourceRowSpan = 2;
            w.Sheets[0].Cards.Add(card);
        });

        using var after = new XlsxPackage(bytes);
        var merges = after.Worksheet("Plants").Descendants<X.MergeCell>().Select(m => m.Reference!.Value!).Order().ToArray();
        AssertSequence(merges, "A4:A5", "B4:B5", "C17:D17");
        var back = Read(bytes).Sheets[0];
        Assert.AreEqual(2, back.Cards[2].SourceRowSpan);
        Assert.AreEqual("Total", after.Text("Plants", "A17"));
    }

    [TestMethod]
    public void ACardWhoseSpanGrewGetsMoreRowsAndAMerge()
    {
        var book = new TestWorkbook();
        book.AddSheet("Plants").Row("A1", "Name", "Note").Row("A2", "Fern", "shade").Row("A3", "Rose").Set("A15", "after");

        var bytes = Change(book.ToBytes(), w => w.Sheets[0].Cards[0].SourceRowSpan = 2);

        using var after = new XlsxPackage(bytes);
        var merges = after.Worksheet("Plants").Descendants<X.MergeCell>().Select(m => m.Reference!.Value!).Order().ToArray();
        AssertSequence(merges, "A2:A3", "B2:B3");
        Assert.AreEqual("Rose", after.Text("Plants", "A4"));
        Assert.AreEqual("after", after.Text("Plants", "A16"));
    }

    /// <summary>A copy of a card as WorkbookEditor.DuplicateCard makes it: same values, no rows of its own, copied from the card's row.</summary>
    private static Card Duplicate(Card source)
    {
        var copy = new Card { SourceRowSpan = source.SourceRowSpan, CopiedFromRow = source.SourceFirstRow ?? source.CopiedFromRow };
        copy.Entries.AddRange(source.Entries.Select(e => new Entry { Text = e.Text, Original = e.Original, Style = e.Style }));
        return copy;
    }

    private static byte[] Doubles()
    {
        var book = new TestWorkbook();
        var bold = book.AddStyle(font: new X.Font(new X.Bold(), new X.FontSize { Val = 11 }, new X.FontName { Val = "Calibri" }));
        book.AddSheet("Sums").Row("A1", "Name", "Value", "Double", "Fixed")
            .Row("A2", "One", 2).Formula("C2", "B2*2", 4).Formula("D2", "$B$2+1", 3).Set("A3", "Two", bold).Set("B3", 3)
            .Formula("C3", "B3*2", 6).Formula("D3", "$B$2+1", 3)
            .Row("A15", "Total").Formula("B15", "SUM(C2:C3)", 10);
        return book.ToBytes();
    }

    [TestMethod]
    public void ADuplicatedCardsFormulasAreTranslatedFromTheCardItCopies()
    {
        var bytes = Change(Doubles(), w => w.Sheets[0].Cards.Insert(2, Duplicate(w.Sheets[0].Cards[1])));

        using var after = new XlsxPackage(bytes);
        Assert.AreEqual("B3*2", Formula(after, "Sums", "C3"));
        Assert.AreEqual("B4*2", Formula(after, "Sums", "C4"), "the copy below its source, as a copy and paste gives");
        Assert.AreEqual("$B$2+1", Formula(after, "Sums", "D4"), "absolute references stay");
        Assert.AreEqual("Two", after.Text("Sums", "A4"));
        Assert.AreEqual(after.Cell("Sums", "A3")!.StyleIndex?.Value, after.Cell("Sums", "A4")!.StyleIndex?.Value, "the copy takes its source's styles");
        Assert.AreEqual("SUM(C2:C3)", Formula(after, "Sums", "B16"), "a row inserted below a range does not grow it");
    }

    [TestMethod]
    public void ACopyOfACopyAboveItsSourceIsTranslatedFromTheSourcesNewRow()
    {
        var bytes = Change(Doubles(), w =>
        {
            var cards = w.Sheets[0].Cards;
            var copy = Duplicate(cards[1]);
            cards.Insert(0, Duplicate(copy));
        });

        using var after = new XlsxPackage(bytes);
        Assert.AreEqual("Two", after.Text("Sums", "A2"));
        Assert.AreEqual("B2*2", Formula(after, "Sums", "C2"));
        Assert.AreEqual("B4*2", Formula(after, "Sums", "C4"), "the source moved down a row");
    }

    [TestMethod]
    public void ADuplicateOfARemovedCardIsTranslatedFromTheRemovedRow()
    {
        var bytes = Change(Doubles(), w =>
        {
            var cards = w.Sheets[0].Cards;
            var copy = Duplicate(cards[0]);
            cards.RemoveAt(0);
            cards.Add(copy);
        });

        using var after = new XlsxPackage(bytes);
        Assert.AreEqual("One", after.Text("Sums", "A3"));
        Assert.AreEqual("B3*2", Formula(after, "Sums", "C3"), "its own row, not #REF! for the deleted row");
        Assert.AreEqual("B2*2", Formula(after, "Sums", "C2"));
        Assert.AreEqual("#REF!+1", Formula(after, "Sums", "D2"), "the kept card's reference to the deleted row");
        Assert.AreEqual("SUM(C2:C2)", Formula(after, "Sums", "B15"));
    }

    [TestMethod]
    public void AnEditedFormulaInADuplicatedCardIsWrittenAsEdited()
    {
        var bytes = Change(Doubles(), w =>
        {
            var copy = Duplicate(w.Sheets[0].Cards[0]);
            copy.Entries[2].Text = "=B3*3";
            w.Sheets[0].Cards.Add(copy);
        });

        using var after = new XlsxPackage(bytes);
        Assert.AreEqual("B3*3", Formula(after, "Sums", "C4"));
    }

    [TestMethod]
    public void AFormulaInACopiedCardWithoutASourceRowIsNotSupported()
    {
        var book = new TestWorkbook();
        book.AddSheet("Sums").Row("A1", "Name", "Double").Row("A2", "One").Formula("B2", "ROW()*2", 4);

        Assert.AreEqual(SpreadsheetWriteError.NotSupported, WriteError(book.ToBytes(), w =>
        {
            var source = w.Sheets[0].Cards[0];
            var copy = new Card();
            copy.Entries.AddRange(source.Entries.Select(e => new Entry { Text = e.Text, Original = e.Original, Style = e.Style }));
            w.Sheets[0].Cards.Add(copy);
        }));
    }

    [TestMethod]
    public void ASharedFormulaBelowTheTableMovesAndStaysShared()
    {
        var book = new TestWorkbook();
        book.AddSheet("Sums").Row("A1", "Name", "Value").Row("A2", "One", 1).Row("A3", "Two", 2)
            .Row("A15", 10).Row("A16", 20).SharedFormula("B15", "A15*2", "B15:B16", 20).SharedDependent("B16", 40);

        var bytes = Change(book.ToBytes(), w => w.Sheets[0].Cards.Add(NewCard("Three", "3")));

        using var after = new XlsxPackage(bytes);
        var master = after.Cell("Sums", "B16")!.CellFormula!;
        Assert.AreEqual((X.CellFormulaValues.Shared, "B16:B17", "A16*2"), (master.FormulaType!.Value, master.Reference!.Value, master.Text));
        Assert.AreEqual(master.SharedIndex!.Value, after.Cell("Sums", "B17")!.CellFormula!.SharedIndex!.Value);
        Assert.AreEqual("", after.Cell("Sums", "B17")!.CellFormula!.Text);
    }

    [TestMethod]
    public void ASharedFormulaSplitByTheChangeGivesEveryCellItsOwnFormula()
    {
        // The group's cells refer to a cell of the table: inserting a row moves the reference of one but not the other's position.
        var book = new TestWorkbook();
        book.AddSheet("Sums").Row("A1", "Name", "Value").Row("A2", "One", 1).Row("A3", "Two", 2)
            .SharedFormula("D2", "B2*2", "D2:D3", 2).SharedDependent("D3", 4);

        var bytes = Change(book.ToBytes(), w => w.Sheets[0].Cards.Insert(1, NewCard("Half", "1.5", "", "")));

        using var after = new XlsxPackage(bytes);
        // D2 stays; D3 moved with its card to D4. Each keeps its own formula.
        Assert.AreEqual("B2*2", Formula(after, "Sums", "D2"));
        Assert.AreEqual("B4*2", Formula(after, "Sums", "D4"));
        Assert.IsTrue(after.Worksheet("Sums").Descendants<X.CellFormula>().All(f => f.FormulaType is null && f.SharedIndex is null));
    }

    [TestMethod]
    public void ASharedFormulaInReorderedCardsIsSplitAndKeepsItsValues()
    {
        var book = new TestWorkbook();
        book.AddSheet("Sums").Row("A1", "Name", "Value", "Double").Row("A2", "One", 1).Row("A3", "Two", 2)
            .SharedFormula("C2", "B2*2", "C2:C3", 2).SharedDependent("C3", 4);

        var bytes = Change(book.ToBytes(), w => w.Sheets[0].Cards.Reverse());

        using var after = new XlsxPackage(bytes);
        // Each cell moved with its card and its formula with it, translated as a copy is.
        Assert.AreEqual("B2*2", Formula(after, "Sums", "C2"));
        Assert.AreEqual("B3*2", Formula(after, "Sums", "C3"));
        Assert.AreEqual("Two", after.Text("Sums", "A2"));
    }

    [TestMethod]
    public void AnArrayFormulaBelowTheTableMovesWhole()
    {
        var book = new TestWorkbook();
        book.AddSheet("Sums").Row("A1", "Name", "Value").Row("A2", "One", 1).Row("A3", "Two", 2)
            .ArrayFormula("B15", "B15:B16", "B2:B3*2", 2).Set("B16", 4);

        var bytes = Change(book.ToBytes(), w => w.Sheets[0].Cards.RemoveAt(0));

        using var after = new XlsxPackage(bytes);
        var formula = after.Cell("Sums", "B14")!.CellFormula!;
        Assert.AreEqual(("B14:B15", "B2:B2*2"), (formula.Reference!.Value, formula.Text));
    }

    [TestMethod]
    public void AnArrayFormulaThatTheChangeSplitsIsNotSupported()
    {
        var book = new TestWorkbook();
        book.AddSheet("Sums").Row("A1", "Name", "Value").Row("A2", "One").Row("A3", "Two").ArrayFormula("B2", "B2:B3", "ROW(A2:A3)", 2).Set("B3", 3);

        Assert.AreEqual(SpreadsheetWriteError.NotSupported, WriteError(book.ToBytes(), w => w.Sheets[0].Cards.Insert(1, NewCard("Half", ""))));
    }

    [TestMethod]
    public void PicturesMoveWithTheirCardsAndARemovedCardsPicturesGo()
    {
        var book = new TestWorkbook();
        book.AddSheet("Plants").Row("A1", "Name", "Photo").Row("A2", "Fern").Row("A3", "Rose").Row("A4", "Lily")
            .Picture("B2", PictureBytes.Png(4, 3)).Picture("B3", PictureBytes.Png(5, 3)).Picture("H20", PictureBytes.Png(6, 3));

        var bytes = Change(book.ToBytes(), w =>
        {
            var cards = w.Sheets[0].Cards;
            cards.RemoveAt(0);
            cards.Reverse();
        });

        using var after = new XlsxPackage(bytes);
        var anchors = after.Part("Plants").DrawingsPart!.WorksheetDrawing!.Elements<Xdr.TwoCellAnchor>().ToList();
        // Fern's picture went with its card; Rose's stayed with Rose (row 3, marker 2); the picture outside the table moved up a row.
        Assert.AreEqual(2, anchors.Count);
        AssertSequence(anchors.Select(a => a.FromMarker!.RowId!.Text), "2", "18");
        Assert.AreEqual(2, after.Part("Plants").DrawingsPart!.ImageParts.Count());
        var back = Read(bytes).Sheets[0];
        AssertSequence(back.Cards.Select(c => c.Entries[0].Text), "Lily", "Rose");
        Assert.AreEqual(0, back.Cards[0].Entries[1].Images.Count);
        Assert.AreEqual(5, back.Cards[1].Entries[1].Images[0].PixelWidth);
    }

    [TestMethod]
    public void ATablePartBelowTheTableMovesWithItsRows()
    {
        var original = Edit(Garden(), document =>
        {
            var entry = document.WorkbookPart!.Workbook!.Sheets!.Elements<X.Sheet>().Single(s => s.Name == "Plants");
            var part = (WorksheetPart)document.WorkbookPart.GetPartById(entry.Id!.Value!);
            var table = part.AddNewPart<TableDefinitionPart>();
            table.Table = new X.Table(
                new X.AutoFilter { Reference = "H5:I6" },
                new X.TableColumns(new X.TableColumn { Id = 1, Name = "Key" }, new X.TableColumn { Id = 2, Name = "Value" }) { Count = 2 })
            { Id = 1, Name = "Lookup", DisplayName = "Lookup", Reference = "H5:I6" };
            part.Worksheet!.AppendChild(new X.TableParts(new X.TablePart { Id = part.GetIdOfPart(table) }) { Count = 1 });
        });

        var bytes = Change(original, w => w.Sheets[0].Cards.Add(NewCard("Tulip", "20")));

        using var after = new XlsxPackage(bytes);
        var moved = after.Part("Plants").TableDefinitionParts.Single().Table!;
        Assert.AreEqual(("H6:I7", "H6:I7"), (moved.Reference!.Value, moved.AutoFilter!.Reference!.Value));
    }

    [TestMethod]
    public void ChartSeriesFollowTheRows()
    {
        var original = Edit(Garden(), document =>
        {
            var entry = document.WorkbookPart!.Workbook!.Sheets!.Elements<X.Sheet>().Single(s => s.Name == "Report");
            var part = (WorksheetPart)document.WorkbookPart.GetPartById(entry.Id!.Value!);
            var drawing = part.AddNewPart<DrawingsPart>();
            drawing.WorksheetDrawing = new Xdr.WorksheetDrawing();
            var chart = drawing.AddNewPart<ChartPart>();
            var series = new C.BarChartSeries(
                new C.Index { Val = 0 }, new C.Order { Val = 0 }, new C.Values(new C.NumberReference(new C.Formula("Plants!$B$2:$B$3"))));
            chart.ChartSpace = new C.ChartSpace(new C.Chart(new C.PlotArea(new C.BarChart(series))));
            part.Worksheet!.AppendChild(new X.Drawing { Id = part.GetIdOfPart(drawing) });
        });

        var bytes = Change(original, w => w.Sheets[0].Cards.Insert(0, NewCard("Moss", "2")));

        using var after = new XlsxPackage(bytes);
        var chart = after.Part("Report").DrawingsPart!.ChartParts.Single();
        // A row inserted at the range's first row moves the whole range, as a spreadsheet program does.
        Assert.AreEqual("Plants!$B$3:$B$4", chart.ChartSpace!.Descendants<C.Formula>().Single().Text);
    }

    [TestMethod]
    public void CommentsMoveWithTheirCells()
    {
        var original = Edit(Garden(), document =>
        {
            var entry = document.WorkbookPart!.Workbook!.Sheets!.Elements<X.Sheet>().Single(s => s.Name == "Plants");
            var part = (WorksheetPart)document.WorkbookPart.GetPartById(entry.Id!.Value!);
            var comments = part.AddNewPart<WorksheetCommentsPart>();
            comments.Comments = new X.Comments(
                new X.Authors(new X.Author("Tester")),
                new X.CommentList(
                    new X.Comment(new X.CommentText(new X.Text("first"))) { Reference = "A2", AuthorId = 0 },
                    new X.Comment(new X.CommentText(new X.Text("below"))) { Reference = "A15", AuthorId = 0 }));
            var vml = part.AddNewPart<VmlDrawingPart>();
            using (var writer = new StreamWriter(vml.GetStream(FileMode.Create)))
            {
                writer.Write("""
                    <xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
                    <v:shape id="_x0000_s1025" type="#_x0000_t202"><x:ClientData ObjectType="Note"><x:Anchor>1, 15, 0, 10, 3, 15, 4, 4</x:Anchor><x:Row>1</x:Row><x:Column>0</x:Column></x:ClientData></v:shape>
                    <v:shape id="_x0000_s1026" type="#_x0000_t202"><x:ClientData ObjectType="Note"><x:Anchor>1, 15, 13, 10, 3, 15, 17, 4</x:Anchor><x:Row>14</x:Row><x:Column>0</x:Column></x:ClientData></v:shape>
                    </xml>
                    """);
            }

            part.Worksheet!.AppendChild(new X.LegacyDrawing { Id = part.GetIdOfPart(vml) });
        });

        var bytes = Change(original, w => w.Sheets[0].Cards.RemoveAt(0));

        using var after = new XlsxPackage(bytes);
        var comment = after.Part("Plants").WorksheetCommentsPart!.Comments!.CommentList!.Elements<X.Comment>().Single();
        Assert.AreEqual(("A14", "below"), (comment.Reference!.Value, comment.InnerText));
        using var reader = new StreamReader(after.Part("Plants").VmlDrawingParts.Single().GetStream());
        var vml = reader.ReadToEnd();
        Assert.IsFalse(vml.Contains("_x0000_s1025", StringComparison.Ordinal));
        StringAssert.Contains(vml, "<x:Row>13</x:Row>");
        StringAssert.Contains(vml, "1, 15, 12, 10, 3, 15, 16, 4");
    }

    [TestMethod]
    public void APivotTableOnASheetWhoseRowsMoveIsNotSupported()
    {
        var original = Edit(Garden(), document =>
        {
            var entry = document.WorkbookPart!.Workbook!.Sheets!.Elements<X.Sheet>().Single(s => s.Name == "Plants");
            var part = (WorksheetPart)document.WorkbookPart.GetPartById(entry.Id!.Value!);
            part.AddNewPart<PivotTablePart>().PivotTableDefinition = new X.PivotTableDefinition { Name = "Pivot", CacheId = 1 };
        });

        Assert.AreEqual(SpreadsheetWriteError.NotSupported, WriteError(original, w => w.Sheets[0].Cards.Add(NewCard("Tulip", "20"))));
        // Writing values in place moves no row and is still fine.
        Change(original, w => w.Sheets[0].Cards[0].Entries[0].Text = "Moss");
    }

    [TestMethod]
    public void ADuplicatedSheetMovesRowsInItsOwnCopyOnly()
    {
        var original = Garden();

        var bytes = Change(original, w =>
        {
            var copy = new Sheet("Plants (2)") { Source = w.Sheets[0].Source! with { IsCopy = true } };
            copy.Columns.AddRange(w.Sheets[0].Columns.Select(c => new Column(c.Name)));
            foreach (var card in w.Sheets[0].Cards)
            {
                var cardCopy = new Card { SourceRowSpan = card.SourceRowSpan };
                cardCopy.Entries.AddRange(card.Entries.Select(e => new Entry { Text = e.Text, Original = e.Original, Style = e.Style }));
                copy.Cards.Add(cardCopy);
            }

            copy.Cards.Add(NewCard("Tulip", "20"));
            w.Sheets.Add(copy);
        });

        using var after = new XlsxPackage(bytes);
        Assert.AreEqual("Total", after.Text("Plants", "A15"));
        Assert.AreEqual("Total", after.Text("Plants (2)", "A16"));
        Assert.AreEqual("Plants!B15", Formula(after, "Report", "B2"));
    }
}
