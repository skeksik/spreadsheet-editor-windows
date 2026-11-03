using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Tests;

[TestClass]
public sealed class RgbColorTests
{
    [TestMethod]
    [DataRow("#1E90FF")]
    [DataRow("1e90ff")]
    [DataRow("FF1E90FF")]
    public void ParsesHexForms(string text)
    {
        Assert.IsTrue(RgbColor.TryParse(text, out var color));
        Assert.AreEqual(new RgbColor(0x1E, 0x90, 0xFF), color);
        Assert.AreEqual("#1E90FF", color.ToString());
    }

    [TestMethod]
    [DataRow(null)]
    [DataRow("")]
    [DataRow("#12345")]
    [DataRow("GGGGGG")]
    public void RejectsOtherText(string? text) => Assert.IsFalse(RgbColor.TryParse(text, out _));
}
