using SpreadsheetEditor.Core.Model;
using SpreadsheetEditor.Core.Values;

namespace SpreadsheetEditor.Core.Tests;

[TestClass]
public sealed class NumberFormatsTests
{
    [TestMethod]
    [DataRow(0, DateTimeFormatKind.None)]
    [DataRow(2, DateTimeFormatKind.None)]
    [DataRow(14, DateTimeFormatKind.Date)]
    [DataRow(17, DateTimeFormatKind.Date)]
    [DataRow(18, DateTimeFormatKind.Time)]
    [DataRow(21, DateTimeFormatKind.Time)]
    [DataRow(22, DateTimeFormatKind.DateAndTime)]
    [DataRow(45, DateTimeFormatKind.Time)]
    [DataRow(46, DateTimeFormatKind.Time)]
    [DataRow(47, DateTimeFormatKind.Time)]
    [DataRow(49, DateTimeFormatKind.None)]
    [DataRow(27, DateTimeFormatKind.Date)]
    [DataRow(32, DateTimeFormatKind.Time)]
    [DataRow(58, DateTimeFormatKind.Date)]
    [DataRow(77, DateTimeFormatKind.DateAndTime)]
    [DataRow(164, DateTimeFormatKind.None)]
    public void ClassifiesBuiltInIds(int id, DateTimeFormatKind kind) =>
        Assert.AreEqual(kind, NumberFormats.Classify(new NumberFormat(id)));

    [TestMethod]
    [DataRow("yyyy-mm-dd", DateTimeFormatKind.Date)]
    [DataRow("dd.mm.yyyy", DateTimeFormatKind.Date)]
    [DataRow("d-mmm", DateTimeFormatKind.Date)]
    [DataRow("mmmm", DateTimeFormatKind.Date)]
    [DataRow("[$-409]mmmm d, yyyy", DateTimeFormatKind.Date)]
    [DataRow("[$-411]ge.m.d", DateTimeFormatKind.Date)]
    [DataRow("h:mm", DateTimeFormatKind.Time)]
    [DataRow("h:mm AM/PM", DateTimeFormatKind.Time)]
    [DataRow("mm:ss", DateTimeFormatKind.Time)]
    [DataRow("[h]:mm:ss", DateTimeFormatKind.Time)]
    [DataRow("[mm]:ss", DateTimeFormatKind.Time)]
    [DataRow("ss.0", DateTimeFormatKind.Time)]
    [DataRow("m/d/yy h:mm", DateTimeFormatKind.DateAndTime)]
    [DataRow("dd.mm.yyyy hh:mm:ss", DateTimeFormatKind.DateAndTime)]
    [DataRow("yyyy-mm-dd;@", DateTimeFormatKind.Date)]
    [DataRow("General", DateTimeFormatKind.None)]
    [DataRow("0.00", DateTimeFormatKind.None)]
    [DataRow("#,##0.00", DateTimeFormatKind.None)]
    [DataRow("0.00E+00", DateTimeFormatKind.None)]
    [DataRow("[Red]0.00;[Blue]-0.00", DateTimeFormatKind.None)]
    [DataRow("#,##0 \"days\"", DateTimeFormatKind.None)]
    [DataRow("\\d0", DateTimeFormatKind.None)]
    [DataRow("[$€-407] #,##0.00", DateTimeFormatKind.None)]
    [DataRow("@", DateTimeFormatKind.None)]
    [DataRow("0.00;\"yyyy\"", DateTimeFormatKind.None)]
    public void ClassifiesFormatCodes(string code, DateTimeFormatKind kind) =>
        Assert.AreEqual(kind, NumberFormats.ClassifyCode(code));

    [TestMethod]
    public void TheCodeDecidesOverTheId()
    {
        Assert.AreEqual(DateTimeFormatKind.None, NumberFormats.Classify(new NumberFormat(14, "0.00")));
        Assert.AreEqual(DateTimeFormatKind.Date, NumberFormats.Classify(new NumberFormat(164, "yyyy")));
    }

    [TestMethod]
    public void NullIsGeneral() => Assert.IsFalse(NumberFormats.IsDateOrTime(null));
}
