using SpreadsheetEditor.Core.Formulas;

namespace SpreadsheetEditor.Core.Tests.Formulas;

[TestClass]
public sealed class FormulaReferencesTests
{
    private const string Sheet = "Data";
    private const int Max = RowShift.MaxRow;

    private static string Insert(string formula, int row, int count, string? formulaSheet = Sheet) =>
        FormulaReferences.ShiftRows(formula, formulaSheet, RowShift.Insert(Sheet, row, count));

    private static string Delete(string formula, int row, int count, string? formulaSheet = Sheet) =>
        FormulaReferences.ShiftRows(formula, formulaSheet, RowShift.Delete(Sheet, row, count));

    // Insertion

    [TestMethod]
    [DataRow("=A5", "=A7")]
    [DataRow("=A3", "=A5")]
    [DataRow("=A2", "=A2")]
    [DataRow("=$A$5", "=$A$7")]
    [DataRow("=A$5", "=A$7")]
    [DataRow("=$A5", "=$A7")]
    [DataRow("=a5", "=A7")]
    [DataRow("=A1:B2", "=A1:B2")]
    [DataRow("=SUM(A1:B5)", "=SUM(A1:B7)")]
    [DataRow("=SUM(A3:B5)", "=SUM(A5:B7)")]
    [DataRow("=SUM($A$1:$A$3)", "=SUM($A$1:$A$5)")]
    [DataRow("=SUM(A5:A1)", "=SUM(A7:A1)")]
    [DataRow("=SUM(1:3)", "=SUM(1:5)")]
    [DataRow("=SUM($4:$4)", "=SUM($6:$6)")]
    [DataRow("=SUM(1:2)", "=SUM(1:2)")]
    [DataRow("=SUM(A:A)+SUM($B:$C)", "=SUM(A:A)+SUM($B:$C)")]
    [DataRow("=A10*B2+C3", "=A12*B2+C5")]
    public void InsertMovesRowsFromTheFirstInsertedRowDown(string formula, string expected) =>
        Assert.AreEqual(expected, Insert(formula, 3, 2));

    [TestMethod]
    public void InsertAtRowOneMovesEverything() => Assert.AreEqual("=A2+B3:C4", Insert("=A1+B2:C3", 1, 1));

    [TestMethod]
    public void InsertPushingACellOffTheSheetGivesRefError()
    {
        Assert.AreEqual("=#REF!", Insert($"=A{Max}", 10, 1));
        Assert.AreEqual($"=A{Max}", Insert($"=A{Max - 1}", 10, 1));
    }

    [TestMethod]
    public void InsertKeepsARangeThatEndsOnTheLastRowEndingThere()
    {
        Assert.AreEqual($"=SUM(A3:A{Max})", Insert($"=SUM(A1:A{Max})", 1, 2));
        Assert.AreEqual($"=SUM(1:{Max})", Insert($"=SUM(1:{Max})", 5, 2));
    }

    [TestMethod]
    public void InsertCutsARangePushedPastTheLastRow() =>
        Assert.AreEqual($"=SUM(A1:A{Max})", Insert($"=SUM(A1:A{Max - 1})", 5, 3));

    // Deletion

    [TestMethod]
    [DataRow("=A2", "=A2")]
    [DataRow("=A3", "=#REF!")]
    [DataRow("=A4", "=#REF!")]
    [DataRow("=A5", "=A3")]
    [DataRow("=$A$5", "=$A$3")]
    [DataRow("=A$9+$A9", "=A$7+$A7")]
    [DataRow("=SUM(A1:A10)", "=SUM(A1:A8)")]
    [DataRow("=SUM(A3:A4)", "=SUM(#REF!)")]
    [DataRow("=SUM(A2:A3)", "=SUM(A2:A2)")]
    [DataRow("=SUM(A4:A6)", "=SUM(A3:A4)")]
    [DataRow("=SUM(A6:A2)", "=SUM(A4:A2)")]
    [DataRow("=SUM(3:4)", "=SUM(#REF!)")]
    [DataRow("=SUM(1:5)", "=SUM(1:3)")]
    [DataRow("=SUM(A:A)", "=SUM(A:A)")]
    [DataRow("=A3:A3", "=#REF!")]
    public void DeleteMovesRowsBelowUpAndRemovesDeletedCells(string formula, string expected) =>
        Assert.AreEqual(expected, Delete(formula, 3, 2));

    [TestMethod]
    public void DeleteOfTheLastRows()
    {
        Assert.AreEqual("=#REF!", Delete($"=A{Max}", Max, 1));
        Assert.AreEqual("=A1", Delete("=A1", Max - 1, 2));
        Assert.AreEqual($"=A{Max - 1}", Delete($"=A{Max}", 1, 1));
    }

    [TestMethod]
    public void DeleteKeepsARangeThatEndsOnTheLastRowEndingThere() =>
        Assert.AreEqual($"=SUM(A1:A{Max})", Delete($"=SUM(A1:A{Max})", 5, 2));

    [TestMethod]
    public void DeleteKeepsTheSheetPrefixBeforeRefError() =>
        Assert.AreEqual("='Data'!#REF!+Data!#REF!", Delete("='Data'!A3+Data!A4", 3, 2, "Other"));

    // Sheets

    [TestMethod]
    public void QualifiedReferencesMoveOnlyForTheShiftedSheet()
    {
        Assert.AreEqual("=Data!A7+Other!A5+'Data'!B7+DATA!C7", Insert("=Data!A5+Other!A5+'Data'!B5+DATA!C5", 3, 2, "Other"));
    }

    [TestMethod]
    public void UnqualifiedReferencesMoveOnlyOnTheShiftedSheet()
    {
        Assert.AreEqual("=A5", Insert("=A5", 3, 2, "Other"));
        Assert.AreEqual("=A5", Insert("=A5", 3, 2, null));
        Assert.AreEqual("=A7", Insert("=A5", 3, 2, "data"));
    }

    [TestMethod]
    public void QuotedSheetNamesWithSpacesAndApostrophes()
    {
        var shift = RowShift.Insert("It's 1", 2, 1);
        Assert.AreEqual("='It''s 1'!A3+'It''s 2'!A2", FormulaReferences.ShiftRows("='It''s 1'!A2+'It''s 2'!A2", "Data", shift));
        var spaced = RowShift.Delete("Sheet 1", 1, 1);
        Assert.AreEqual("=SUM('Sheet 1'!$A$1:$A$9)", FormulaReferences.ShiftRows("=SUM('Sheet 1'!$A$2:$A$10)", null, spaced));
    }

    [TestMethod]
    public void UnicodeSheetNames() =>
        Assert.AreEqual("=Лист1!B4", FormulaReferences.ShiftRows("=Лист1!B3", null, RowShift.Insert("лист1", 1, 1)));

    [TestMethod]
    [DataRow("=Data:Other!A5")]
    [DataRow("='Data:Other'!A5")]
    [DataRow("=[1]Data!A5")]
    [DataRow("='[1]Data'!A5")]
    public void MultiSheetAndExternalReferencesAreLeftAlone(string formula) => Assert.AreEqual(formula, Insert(formula, 1, 3));

    // Text that is not a reference

    [TestMethod]
    [DataRow("=\"A5\"&\"Data!A5\"&\"'Data'!A5\"")]
    [DataRow("=\"say \"\"A5\"\"\"")]
    [DataRow("=LOG10(A)")]
    [DataRow("=MyRange5+Price")]
    [DataRow("=Table1[Price]+Table1[[#This Row],[A5]]+[@A5]")]
    [DataRow("=Table1[Col']'A5]")]
    [DataRow("=#REF!+Data!#REF!")]
    [DataRow("=1E5+2.5")]
    [DataRow("=R5C1")]
    [DataRow("=_xlfn.XLOOKUP(Key,Keys,Values)")]
    public void LeavesNonReferencesAlone(string formula) => Assert.AreEqual(formula, Insert(formula, 1, 3));

    [TestMethod]
    public void ShiftsReferencesAroundStringsAndFunctions() =>
        Assert.AreEqual("=IF(A7=\"A5\",VLOOKUP(B7,Table1[A5],2),Data!C7&\"x\")",
            Insert("=IF(A5=\"A5\",VLOOKUP(B5,Table1[A5],2),Data!C5&\"x\")", 3, 2));

    [TestMethod]
    public void ShiftsSpillAndImplicitIntersection() => Assert.AreEqual("=SUM(A7#)+@B7", Insert("=SUM(A5#)+@B5", 3, 2));

    [TestMethod]
    public void DefinedNamesAndPrintAreasAreFormulas()
    {
        var shift = RowShift.Insert("Sheet 1", 5, 3);
        Assert.AreEqual("'Sheet 1'!$A$1:$D$23", FormulaReferences.ShiftRows("'Sheet 1'!$A$1:$D$20", null, shift));
        Assert.AreEqual("'Sheet 1'!$A$1:$D$23,'Sheet 1'!$1:$2",
            FormulaReferences.ShiftRows("'Sheet 1'!$A$1:$D$20,'Sheet 1'!$1:$2", null, shift));
    }

    [TestMethod]
    public void UnchangedFormulaIsReturnedAsIs()
    {
        const string formula = "=sum( a1 , b2 )";
        Assert.AreSame(formula, Insert(formula, 10, 1));
    }

    // Round trips

    private static readonly string[] RoundTripFormulas =
    [
        "=A1", "=A2", "=A3", "=A5", "=A9", "=$A$5", "=A$5", "=$A5", "=XFD1048571", "=A1048571",
        "=SUM(A1:B10)", "=SUM(A3:A4)", "=SUM(A5:A1)", "=SUM($B$2:$C$20)", "=SUM(A1:A1048576)", "=SUM(4:1048576)",
        "=SUM(1:3)", "=SUM($3:$9)", "=SUM(A:A)+SUM($B:$XFD)",
        "=Data!A5+'Data'!B6:C9+Other!A5", "='It''s'!A5", "=Data:Other!A5+[1]Data!A5",
        "=IF(A5=\"A5\",VLOOKUP(B5,Table1[A5],2),Data!C5&\"x\")", "=SUM(Table1[[#This Row],[Price]])*[@Qty]",
        "=#REF!+MyName5+LOG10(A6)+1E5", "=SUM(A5#)+@B5", "=OFFSET(A1,1,1):B9", "=A4 B5:C6",
    ];

    [TestMethod]
    [DataRow(1, 1)]
    [DataRow(1, 5)]
    [DataRow(3, 2)]
    [DataRow(5, 1)]
    [DataRow(6, 4)]
    [DataRow(1048570, 1)]
    public void InsertThenDeleteGivesTheOriginal(int row, int count)
    {
        foreach (var formula in RoundTripFormulas)
        {
            var inserted = Insert(formula, row, count);
            Assert.AreEqual(formula, Delete(inserted, row, count), $"{formula} via {inserted}");
        }
    }

    [TestMethod]
    public void InsertThenDeleteGivesTheOriginalForEveryRowOfASmallSheet()
    {
        foreach (var formula in RoundTripFormulas)
        {
            for (var row = 1; row <= 12; row++)
            {
                for (var count = 1; count <= 3; count++)
                {
                    Assert.AreEqual(formula, Delete(Insert(formula, row, count), row, count), $"{formula} at {row}+{count}");
                }
            }
        }
    }

    // Ranges

    [TestMethod]
    [DataRow("A1:C3", "A1:C5")]
    [DataRow("A5:C6", "A7:C8")]
    [DataRow("A1:B2", "A1:B2")]
    [DataRow("$A$1:$D$20", "$A$1:$D$22")]
    [DataRow("B4", "B6")]
    [DataRow("A1:C3 E5:F9 A1", "A1:C5 E7:F11 A1")]
    public void ShiftRangeInsert(string range, string expected) =>
        Assert.AreEqual(expected, FormulaReferences.ShiftRange(range, RowShift.Insert(Sheet, 3, 2)));

    [TestMethod]
    [DataRow("A1:C3", "A1:C2")]
    [DataRow("A3:C4", null)]
    [DataRow("A5:C6", "A3:C4")]
    [DataRow("A3:C4 E1:E9 B3", "E1:E7")]
    [DataRow("A3 B4", null)]
    public void ShiftRangeDelete(string range, string? expected) =>
        Assert.AreEqual(expected, FormulaReferences.ShiftRange(range, RowShift.Delete(Sheet, 3, 2)));

    [TestMethod]
    public void ShiftRangeRoundTrip()
    {
        foreach (var range in new[] { "A1:C3", "A5:C6", "B4", "A1:C3 E5:F9", "A1:XFD1048576", "D1048573" })
        {
            var inserted = FormulaReferences.ShiftRange(range, RowShift.Insert(Sheet, 2, 3))!;
            Assert.AreEqual(range, FormulaReferences.ShiftRange(inserted, RowShift.Delete(Sheet, 2, 3)));
        }
    }

    [TestMethod]
    [DataRow("Sheet1!A1")]
    [DataRow("A1:")]
    [DataRow("SUM(A1)")]
    [DataRow("A0")]
    public void ShiftRangeRejectsTextThatIsNotARange(string range) =>
        Assert.ThrowsExactly<FormatException>(() => FormulaReferences.ShiftRange(range, RowShift.Insert(Sheet, 1, 1)));

    // Translation (shared formulas)

    [TestMethod]
    [DataRow("=A1+B2", 2, 1, "=B3+C4")]
    [DataRow("=$A$1+B2", 2, 1, "=$A$1+C4")]
    [DataRow("=A$1*$A1", 3, 2, "=C$1*$A4")]
    [DataRow("=SUM(A1:A5)", 1, 0, "=SUM(A2:A6)")]
    [DataRow("=SUM($A1:A$5)", 4, 4, "=SUM($A5:E$5)")]
    [DataRow("=SUM(1:3)+SUM(A:B)", 2, 2, "=SUM(3:5)+SUM(C:D)")]
    [DataRow("=SUM($1:3)+SUM($A:B)", 2, 2, "=SUM($1:5)+SUM($A:D)")]
    [DataRow("=Other!A1+'It''s'!B2", 1, 1, "=Other!B2+'It''s'!C3")]
    [DataRow("=\"A1\"&Table1[A1]&LOG10(A1)", 1, 1, "=\"A1\"&Table1[A1]&LOG10(B2)")]
    [DataRow("=B3", -1, -1, "=A2")]
    [DataRow("=A1", 0, 0, "=A1")]
    public void TranslateMovesRelativePartsOnly(string formula, int rows, int columns, string expected) =>
        Assert.AreEqual(expected, FormulaReferences.Translate(formula, rows, columns));

    [TestMethod]
    public void TranslateOffTheSheetGivesRefError()
    {
        Assert.AreEqual("=#REF!+$B$1", FormulaReferences.Translate("=A1+$B$1", -1, 0));
        Assert.AreEqual("=Other!#REF!", FormulaReferences.Translate("=Other!A1", 0, -1));
        Assert.AreEqual("=#REF!", FormulaReferences.Translate($"=XFD{Max}", 0, 1));
        Assert.AreEqual("=#REF!", FormulaReferences.Translate($"=A{Max}", 1, 0));
        Assert.AreEqual($"=$XFD${Max}", FormulaReferences.Translate($"=$XFD${Max}", 5, 5));
        Assert.AreEqual($"=XFD{Max}", FormulaReferences.Translate("=XFC1048575", 1, 1));
    }

    [TestMethod]
    public void TranslateBackGivesTheOriginal()
    {
        const string shared = "=SUM($A1:B$2)*C3+Data!D4-'It''s'!$E5";
        Assert.AreEqual(shared, FormulaReferences.Translate(FormulaReferences.Translate(shared, 7, 3), -7, -3));
    }

    // Arguments

    [TestMethod]
    public void RowShiftChecksItsArguments()
    {
        Assert.ThrowsExactly<ArgumentOutOfRangeException>(() => RowShift.Insert(Sheet, 0, 1));
        Assert.ThrowsExactly<ArgumentOutOfRangeException>(() => RowShift.Insert(Sheet, 1, 0));
        Assert.ThrowsExactly<ArgumentOutOfRangeException>(() => RowShift.Delete(Sheet, Max, 2));
        Assert.ThrowsExactly<ArgumentNullException>(() => RowShift.Delete(null!, 1, 1));
        Assert.ThrowsExactly<ArgumentException>(() => FormulaReferences.ShiftRows("=A1", Sheet, default));
    }
}
