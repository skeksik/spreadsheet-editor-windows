using SpreadsheetEditor.Core.Xlsx;

namespace SpreadsheetEditor.Core.Tests.Xlsx;

[TestClass]
public sealed class CellReferenceTests
{
    [TestMethod]
    [DataRow("A1", 1, 1)]
    [DataRow("b12", 2, 12)]
    [DataRow("$AB$7", 28, 7)]
    [DataRow("XFD1048576", 16384, 1048576)]
    public void ParsesReferences(string text, int column, int row)
    {
        Assert.IsTrue(CellReference.TryParse(text, out var c, out var r));
        Assert.AreEqual((column, row), (c, r));
    }

    [TestMethod]
    [DataRow("")]
    [DataRow("A")]
    [DataRow("12")]
    [DataRow("A0")]
    [DataRow("XFE1")]
    [DataRow("A1048577")]
    [DataRow("A1B")]
    public void RejectsOtherText(string text) => Assert.IsFalse(CellReference.TryParse(text, out _, out _));

    [TestMethod]
    public void ParsesRanges()
    {
        Assert.IsTrue(CellReference.TryParseRange("C3:A1", out var range));
        Assert.AreEqual(new MergeRange(1, 1, 3, 3), range);
        Assert.IsTrue(CellReference.TryParseRange("B2", out range));
        Assert.AreEqual(new MergeRange(2, 2, 2, 2), range);
        Assert.IsFalse(CellReference.TryParseRange("A1:", out _));
    }

    [TestMethod]
    [DataRow(1, "A")]
    [DataRow(26, "Z")]
    [DataRow(27, "AA")]
    [DataRow(16384, "XFD")]
    public void NamesColumns(int column, string name) => Assert.AreEqual(name, CellReference.ColumnName(column));
}
