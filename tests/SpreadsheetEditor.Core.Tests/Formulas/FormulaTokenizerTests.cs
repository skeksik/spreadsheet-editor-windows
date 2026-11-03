using SpreadsheetEditor.Core.Formulas;

namespace SpreadsheetEditor.Core.Tests.Formulas;

[TestClass]
public sealed class FormulaTokenizerTests
{
    private static readonly string[] EachKind = ["A1", "$B$2:C3", "1:3", "$A:$C", "Sheet2!D4", "'My sheet'!E5:F6", "a1"];
    private static readonly string[] Edges = ["XFD1048576", "A1:XFD1", "1048576:1048576", "XFD:XFD"];

    private static List<(FormulaTokenKind Kind, string Text)> Tokens(string formula) =>
        [.. FormulaTokenizer.Tokenize(formula).Select(t => (t.Kind, formula.Substring(t.Start, t.Length)))];

    private static List<string> References(string formula) =>
        [.. Tokens(formula).Where(t => t.Kind == FormulaTokenKind.Reference).Select(t => t.Text)];

    [TestMethod]
    [DataRow("=SUM(A1:B2)*'Sheet 1'!$C$3&\"x\"")]
    [DataRow("Table1[[#This Row],[Price]]+[@Qty]")]
    [DataRow("IF(ISERROR(#N/A),{1,2;3,4},1.5E-3)")]
    [DataRow("'unclosed")]
    [DataRow("\"unclosed")]
    [DataRow("[unclosed")]
    [DataRow("")]
    public void TokensCoverTheWholeText(string formula)
    {
        var tokens = FormulaTokenizer.Tokenize(formula);
        var position = 0;
        foreach (var token in tokens)
        {
            Assert.AreEqual(position, token.Start);
            Assert.IsGreaterThan(0, token.Length);
            position = token.End;
        }
        Assert.AreEqual(formula.Length, position);
    }

    [TestMethod]
    public void FindsEachKindOfReference()
    {
        CollectionAssert.AreEqual(
            EachKind,
            References("=A1+SUM($B$2:C3)+SUM(1:3)+SUM($A:$C)+Sheet2!D4+SUM('My sheet'!E5:F6)+a1"));
    }

    [TestMethod]
    [DataRow("=\"A1\"&\"B2:C3\"")]
    [DataRow("=LOG10(2)")]
    [DataRow("=MyName+_xlnm.Print_Area")]
    [DataRow("=Table1[A1]")]
    [DataRow("=R1C1")]
    [DataRow("=ABCD1")]
    [DataRow("=XFE1")]
    [DataRow("=A0")]
    [DataRow("=A1048577")]
    [DataRow("=A01")]
    [DataRow("=1E5+2.5")]
    [DataRow("=TRUE")]
    [DataRow("=Sheet1!MyName")]
    [DataRow("=#REF!+#DIV/0!")]
    public void FindsNoReference(string formula) => Assert.IsEmpty(References(formula));

    [TestMethod]
    public void ReadsFunctionsNamesStringsAndErrors()
    {
        var tokens = Tokens("=SUM(Price,\"a\"\"b\",#N/A)");
        CollectionAssert.Contains(tokens, (FormulaTokenKind.Function, "SUM"));
        CollectionAssert.Contains(tokens, (FormulaTokenKind.Name, "Price"));
        CollectionAssert.Contains(tokens, (FormulaTokenKind.String, "\"a\"\"b\""));
        CollectionAssert.Contains(tokens, (FormulaTokenKind.Error, "#N/A"));
    }

    [TestMethod]
    public void UnquotesSheetNames()
    {
        var token = FormulaTokenizer.Tokenize("='It''s here'!A1").Single(t => t.Kind == FormulaTokenKind.Reference);
        Assert.AreEqual("It's here", token.SheetName);
        Assert.AreEqual("'It''s here'!".Length, token.PrefixLength);
        Assert.IsFalse(token.IsMultiSheetOrExternal);
    }

    [TestMethod]
    [DataRow("=Sheet1:Sheet3!A1")]
    [DataRow("='Sheet 1:Sheet 3'!A1")]
    [DataRow("=[1]Sheet1!A1")]
    [DataRow("='[1]Sheet 1'!A1")]
    public void MarksMultiSheetAndExternalReferences(string formula)
    {
        var token = FormulaTokenizer.Tokenize(formula).Single(t => t.Kind == FormulaTokenKind.Reference);
        Assert.IsTrue(token.IsMultiSheetOrExternal);
    }

    [TestMethod]
    [DataRow("=Sheet1:Sheet3!A1", false)]
    [DataRow("='Sheet 1:Sheet 3'!A1", false)]
    [DataRow("=[1]Sheet1!A1", true)]
    [DataRow("=[1]Sheet1:Sheet3!A1", true)]
    [DataRow("='[1]Sheet 1'!A1", true)]
    [DataRow("=[1]'Sheet 1'!A1", true)]
    [DataRow("='C:\\Files\\[Book.xlsx]Sheet 1'!A1", true)]
    [DataRow("=Sheet1!A1", false)]
    public void MarksExternalReferences(string formula, bool external)
    {
        var token = FormulaTokenizer.Tokenize(formula).Single(t => t.Kind == FormulaTokenKind.Reference);
        Assert.AreEqual(external, token.IsExternal);
    }

    [TestMethod]
    [DataRow("=Sheet1!MyName", "Sheet1", "Sheet1!")]
    [DataRow("='My sheet'!#REF!", "My sheet", "'My sheet'!")]
    [DataRow("=Sheet1:Sheet3!Total", "Sheet1:Sheet3", "Sheet1:Sheet3!")]
    public void SheetPrefixWithoutAreaCarriesItsSheetName(string formula, string sheet, string prefix)
    {
        var token = FormulaTokenizer.Tokenize(formula).First(t => t.SheetName is not null);
        Assert.AreEqual(FormulaTokenKind.Other, token.Kind);
        Assert.AreEqual(sheet, token.SheetName);
        Assert.AreEqual(prefix.Length, token.PrefixLength);
        Assert.AreEqual(prefix, formula.Substring(token.Start, token.Length));
    }

    [TestMethod]
    public void ReadsTheEdgesOfTheSheet()
    {
        CollectionAssert.AreEqual(Edges,
            References("=XFD1048576+A1:XFD1+1048576:1048576+XFD:XFD"));
    }
}
