using SpreadsheetEditor.Core.Model;
using SpreadsheetEditor.Core.Xlsx;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetEditor.Core.Tests.Xlsx;

[TestClass]
public sealed class XlsxColorsTests
{
    private static RgbColor Hex(string hex) => RgbColor.TryParse(hex, out var color) ? color : throw new ArgumentException(hex);

    private static ResolvedColor Resolve(X.ColorType color) => XlsxColors.Default.Resolve(color);

    [TestMethod]
    [DataRow("FF112233", "#112233")]
    [DataRow("80112233", "#112233")]
    [DataRow("00112233", "#112233")]
    [DataRow("112233", "#112233")]
    [DataRow("ffaabbcc", "#AABBCC")]
    public void ExplicitColourIgnoresAlpha(string argb, string expected)
    {
        var color = XlsxColors.Argb(argb);

        Assert.AreEqual(ColorKind.Rgb, color.Kind);
        Assert.AreEqual(expected, color.Rgb.ToString());
    }

    [TestMethod]
    public void FullyTransparentIsNoColour() => Assert.AreEqual(ResolvedColor.None, XlsxColors.Argb("00000000"));

    [TestMethod]
    [DataRow("")]
    [DataRow("12345")]
    [DataRow("1234567")]
    [DataRow("GG112233")]
    [DataRow("#112233")]
    public void MalformedColourIsUnresolved(string argb) => Assert.AreEqual(ResolvedColor.Unresolved, XlsxColors.Argb(argb));

    [TestMethod]
    public void RgbAttribute() =>
        Assert.AreEqual(ResolvedColor.Of(Hex("C00000")), Resolve(new X.Color { Rgb = "FFC00000" }));

    [TestMethod]
    public void AutoComesFirst() =>
        Assert.AreEqual(ResolvedColor.Automatic, Resolve(new X.Color { Auto = true, Rgb = "FFC00000" }));

    [TestMethod]
    public void ColourWithoutAttributesIsUnresolved() => Assert.AreEqual(ResolvedColor.Unresolved, Resolve(new X.Color()));

    [TestMethod]
    [DataRow(0u, "#FFFFFF")]
    [DataRow(1u, "#000000")]
    [DataRow(2u, "#E7E6E6")]
    [DataRow(3u, "#44546A")]
    [DataRow(4u, "#5B9BD5")]
    [DataRow(9u, "#70AD47")]
    [DataRow(10u, "#0563C1")]
    [DataRow(11u, "#954F72")]
    public void ThemeColourFromTheDefaultPalette(uint index, string expected) =>
        Assert.AreEqual(expected, Resolve(new X.Color { Theme = index }).Rgb.ToString());

    [TestMethod]
    public void ThemeIndexOutOfRangeIsUnresolved() => Assert.AreEqual(ResolvedColor.Unresolved, Resolve(new X.Color { Theme = 12 }));

    [TestMethod]
    public void ThemeColourWithTint() =>
        Assert.AreEqual("#9DC3E6", Resolve(new X.Color { Theme = 4, Tint = 0.3999755851924192 }).Rgb.ToString());

    [TestMethod]
    public void FillForegroundColourResolvesToo() =>
        Assert.AreEqual("#2E75B6", XlsxColors.Default.Resolve(new X.ForegroundColor { Theme = 4, Tint = -0.249977111117893 }).Rgb.ToString());

    // Expected values computed by hand with the algorithm of docs/spec.md §4.4 step 4.
    [TestMethod]
    [DataRow("5B9BD5", 0.3999755851924192, "#9DC3E6")]
    [DataRow("5B9BD5", -0.249977111117893, "#2E75B6")]
    [DataRow("000000", 0.5, "#808080")]
    [DataRow("FFFFFF", -0.5, "#808080")]
    [DataRow("E7E6E6", -0.0999786370433668, "#D0CECE")]
    [DataRow("44546A", 0.7999816888943144, "#D6DCE5")]
    [DataRow("123456", 0.0, "#123456")]
    public void TintChangesLightness(string color, double tint, string expected) =>
        Assert.AreEqual(expected, ColorTint.Apply(Hex(color), tint).ToString());

    [TestMethod]
    [DataRow(0u, "#000000")]
    [DataRow(2u, "#FF0000")]
    [DataRow(8u, "#000000")]
    [DataRow(10u, "#FF0000")]
    [DataRow(22u, "#C0C0C0")]
    [DataRow(30u, "#0066CC")]
    [DataRow(52u, "#FF9900")]
    [DataRow(63u, "#333333")]
    public void IndexedColourFromTheStandardPalette(uint index, string expected) =>
        Assert.AreEqual(expected, Resolve(new X.Color { Indexed = index }).Rgb.ToString());

    [TestMethod]
    [DataRow(64u)]
    [DataRow(65u)]
    public void SystemColoursAreAutomatic(uint index) => Assert.AreEqual(ResolvedColor.Automatic, Resolve(new X.Color { Indexed = index }));

    [TestMethod]
    public void IndexOutsideThePaletteIsUnresolved() => Assert.AreEqual(ResolvedColor.Unresolved, Resolve(new X.Color { Indexed = 66 }));

    [TestMethod]
    public void StyleSheetPaletteReplacesTheStandardOne()
    {
        var stylesheet = new X.Stylesheet(new X.Colors(new X.IndexedColors(
            new X.RgbColor { Rgb = "00102030" },
            new X.RgbColor { Rgb = "not a colour" })));
        var palette = IndexedPalette.Read(stylesheet);

        Assert.AreEqual(ResolvedColor.Of(Hex("102030")), palette.Resolve(0));
        Assert.AreEqual(ResolvedColor.Unresolved, palette.Resolve(1));
        Assert.AreEqual(ResolvedColor.Of(Hex("FF0000")), palette.Resolve(2), "an index the custom palette does not reach keeps the standard colour");
        Assert.AreEqual(ResolvedColor.Automatic, palette.Resolve(64));
    }

    [TestMethod]
    public void TintIsNotAppliedToIndexedOrExplicitColours()
    {
        Assert.AreEqual("#FF0000", Resolve(new X.Color { Indexed = 2, Tint = 0.5 }).Rgb.ToString());
        Assert.AreEqual("#FF0000", Resolve(new X.Color { Rgb = "FFFF0000", Tint = 0.5 }).Rgb.ToString());
    }
}
