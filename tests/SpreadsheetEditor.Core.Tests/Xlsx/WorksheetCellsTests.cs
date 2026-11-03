using System.Globalization;
using SpreadsheetEditor.Core.Model;
using SpreadsheetEditor.Core.Xlsx;

namespace SpreadsheetEditor.Core.Tests.Xlsx;

/// <summary>Shared, array and data-table formulas as the reader shows them (docs/spec.md §9.3; ECMA-376 Part 1 §18.3.1.40).</summary>
[TestClass]
public sealed class WorksheetCellsTests
{
    private static Sheet ReadSingle(TestWorkbook workbook)
    {
        using var stream = workbook.ToStream();
        var sheets = XlsxReader.Read(stream, CultureInfo.InvariantCulture).Sheets;
        Assert.AreEqual(1, sheets.Count);
        return sheets[0];
    }

    /// <summary>Checks the text of one column's entry on every card.</summary>
    private static void AssertColumn(Sheet sheet, int column, params string[] expected) =>
        CollectionAssert.AreEqual(expected, sheet.Cards.Select(c => c.Entries[column].Text).ToArray());

    private static void AssertTexts(IEnumerable<Entry> entries, params string[] expected) =>
        CollectionAssert.AreEqual(expected, entries.Select(e => e.Text).ToArray());

    [TestMethod]
    public void SharedFormulaDownAColumnIsShownOnEveryCell()
    {
        var book = new TestWorkbook();
        book.AddSheet("S").Row("A1", "Price", "Double")
            .Set("A2", 1).SharedMaster("B2", 0, "B2:B4", "A2*2", 2)
            .Set("A3", 2).Shared("B3", 0, 4)
            .Set("A4", 3).Shared("B4", 0, 6);

        var sheet = ReadSingle(book);

        AssertColumn(sheet, 1, "=A2*2", "=A3*2", "=A4*2");
        var dependent = sheet.Cards[2].Entries[1].Original!;
        Assert.AreEqual(
            new CellValue(CellValueKind.Number, "=A4*2", Number: 6, Formula: "A4*2", Format: NumberFormat.General), dependent);
    }

    [TestMethod]
    public void SharedFormulaAlongARowIsShownOnEveryCell()
    {
        var book = new TestWorkbook();
        book.AddSheet("S").Row("A1", "Start", "Next", "Then", "Last")
            .Set("A2", 1).SharedMaster("B2", 3, "B2:D2", "A2+1", 2).Shared("C2", 3, 3).Shared("D2", 3, 4);

        var entries = ReadSingle(book).Cards[0].Entries;

        AssertTexts(entries, "1", "=A2+1", "=B2+1", "=C2+1");
        Assert.AreEqual(4, entries[3].Original!.Number);
    }

    [TestMethod]
    public void SharedFormulaBlockKeepsAbsolutePartsAndMovesRelativeOnes()
    {
        var book = new TestWorkbook();
        book.AddSheet("S").Row("A1", "Rate", "Low", "High")
            .Set("A2", 1).SharedMaster("B2", 1, "B2:C3", "$A2*B$1+$A$1+SUM(A2:$A3)", 0).Shared("C2", 1, 0)
            .Set("A3", 2).Shared("B3", 1, 0).Shared("C3", 1, 0);

        var sheet = ReadSingle(book);

        AssertTexts(sheet.Cards[0].Entries.Skip(1), "=$A2*B$1+$A$1+SUM(A2:$A3)", "=$A2*C$1+$A$1+SUM(B2:$A3)");
        AssertTexts(sheet.Cards[1].Entries.Skip(1), "=$A3*B$1+$A$1+SUM(A3:$A4)", "=$A3*C$1+$A$1+SUM(B3:$A4)");
    }

    [TestMethod]
    public void SharedFormulaWithTextResultsKeepsTheCachedText()
    {
        var book = new TestWorkbook();
        book.AddSheet("S").Row("A1", "Name", "Loud")
            .Set("A2", "fern").SharedMaster("B2", 0, "B2:B3", "UPPER(A2)&\"!\"", "FERN!")
            .Set("A3", "rose").Shared("B3", 0, "ROSE!");

        var value = ReadSingle(book).Cards[1].Entries[1].Original!;

        Assert.AreEqual(CellValueKind.Text, value.Kind);
        Assert.AreEqual("UPPER(A3)&\"!\"", value.Formula);
        Assert.AreEqual("=UPPER(A3)&\"!\"", value.DisplayText);
    }

    [TestMethod]
    public void MasterOutsideTheTableAfterItsDependents()
    {
        // The master lies more than 10 rows below the table, so it is not part of it, and it comes after its dependents in the XML.
        var book = new TestWorkbook();
        book.AddSheet("S").Row("A1", "Count", "Total")
            .Set("A2", 1).Shared("B2", 5, 10)
            .Set("A3", 2).Shared("B3", 5, 20)
            .SharedMaster("B20", 5, "B2:B20", "A20*10", 0);

        var sheet = ReadSingle(book);

        Assert.AreEqual(3, sheet.Source!.LastRow);
        AssertColumn(sheet, 1, "=A2*10", "=A3*10");
        Assert.AreEqual(20, sheet.Cards[1].Entries[1].Original!.Number);
        var map = sheet.Source.SharedFormulas;
        Assert.AreEqual(new SharedFormulaGroup(5, 20, 2, "B2:B20", "A20*10"), map.Group(5));
        Assert.AreEqual(new SharedFormulaCell(5, true), map.Get(20, 2));
        Assert.AreEqual(new SharedFormulaCell(5, false), map.Get(2, 2));
    }

    [TestMethod]
    public void ReferenceMovedOffTheSheetBecomesRefError()
    {
        var book = new TestWorkbook();
        book.AddSheet("S").Row("A1", "A", "B").Set("A2", 1).Shared("B2", 0, 0).Set("A3", 2).SharedMaster("B3", 0, "B2:B3", "A2", 1);

        Assert.AreEqual("=A1", ReadSingle(book).Cards[0].Entries[1].Text);

        var edge = new TestWorkbook();
        edge.AddSheet("S").Row("A1", "A", "B").Shared("A2", 0, 0).SharedMaster("B2", 0, "A2:B2", "A1", 1);

        Assert.AreEqual("=#REF!", ReadSingle(edge).Cards[0].Entries[0].Text);
    }

    [TestMethod]
    public void DependentWithoutAMasterKeepsItsCachedValue()
    {
        var book = new TestWorkbook();
        book.AddSheet("S").Row("A1", "A", "B").Set("A2", 1).Shared("B2", 7, 42);

        var sheet = ReadSingle(book);

        var value = sheet.Cards[0].Entries[1].Original!;
        Assert.AreEqual(new CellValue(CellValueKind.Number, "42", Number: 42, Format: NumberFormat.General), value);
        Assert.AreEqual(new SharedFormulaCell(7, false), sheet.Source!.SharedFormulas.Get(2, 2));
        Assert.IsNull(sheet.Source.SharedFormulas.Group(7));
    }

    [TestMethod]
    public void MapListsEveryCellOfEveryGroup()
    {
        var book = new TestWorkbook();
        book.AddSheet("S").Row("A1", "A", "B", "C")
            .Set("A2", 1).SharedMaster("B2", 0, "B2:B3", "A2", 1).SharedMaster("C2", 1, "C2:C3", "B2", 1)
            .Set("A3", 2).Shared("B3", 0, 2).Shared("C3", 1, 2)
            .Formula("A4", "1+1", 2);

        var map = ReadSingle(book).Source!.SharedFormulas;

        Assert.AreEqual(4, map.CellCount);
        Assert.AreEqual(2, map.Groups.Count);
        Assert.AreEqual(new SharedFormulaCell(0, true), map.Get(2, 2));
        Assert.AreEqual(new SharedFormulaCell(1, true), map.Get(2, 3));
        Assert.AreEqual(new SharedFormulaCell(0, false), map.Get(3, 2));
        Assert.AreEqual(new SharedFormulaCell(1, false), map.Get(3, 3));
        Assert.IsNull(map.Get(4, 1));
        Assert.IsNull(map.Get(2, 1));
        Assert.AreEqual(new SharedFormulaGroup(1, 2, 3, "C2:C3", "B2"), map.Group(1));
    }

    [TestMethod]
    public void WorkbookWithoutSharedFormulasHasAnEmptyMap()
    {
        var book = new TestWorkbook();
        book.AddSheet("S").Row("A1", "A").Formula("A2", "1+1", 2);

        var map = ReadSingle(book).Source!.SharedFormulas;

        Assert.AreEqual(0, map.CellCount);
        Assert.AreEqual(0, map.Groups.Count);
    }

    [TestMethod]
    public void ArrayFormulaIsShownOnItsTopLeftCellOnly()
    {
        var book = new TestWorkbook();
        book.AddSheet("S").Row("A1", "A", "B")
            .Row("A2", 1, null).ArrayFormula("B2", "B2:B3", "A2:A3*2", 2)
            .Row("A3", 2, 4);

        var sheet = ReadSingle(book);

        AssertColumn(sheet, 1, "=A2:A3*2", "4");
        Assert.IsNull(sheet.Cards[1].Entries[1].Original!.Formula);
        Assert.AreEqual(0, sheet.Source!.SharedFormulas.CellCount);
    }

    [TestMethod]
    public void DataTableCellsShowTheirCachedValues()
    {
        var book = new TestWorkbook();
        book.AddSheet("S").Row("A1", "Input", "Result")
            .Set("A2", 1).DataTable("B2", "B2:B3", "A1", 10)
            .Row("A3", 2, 20);

        var sheet = ReadSingle(book);

        AssertColumn(sheet, 1, "10", "20");
        Assert.IsNull(sheet.Cards[0].Entries[1].Original!.Formula);
    }
}
