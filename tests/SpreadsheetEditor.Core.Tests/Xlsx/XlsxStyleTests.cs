using System.Globalization;
using SpreadsheetEditor.Core.Model;
using SpreadsheetEditor.Core.Xlsx;
using A = DocumentFormat.OpenXml.Drawing;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetEditor.Core.Tests.Xlsx;

/// <summary>Fills, text colours, fonts and alignment read from <c>.xlsx</c> (docs/spec.md §4.2, §4.4, §9.5, §9.6).</summary>
[TestClass]
public sealed class XlsxStyleTests
{
    private static RgbColor Hex(string hex) => RgbColor.TryParse(hex, out var color) ? color : throw new ArgumentException(hex);

    /// <summary>Reads a sheet whose header is A1:B1 and whose only card is row 2, with B2 in <paramref name="style"/>.</summary>
    private static CellStyle StyleOf(TestWorkbook book, uint style)
    {
        book.AddSheet("S").Row("A1", "Name", "Value").Set("A2", "x").Set("B2", "styled", style);
        return EntryStyle(book, 0, 1);
    }

    private static CellStyle EntryStyle(TestWorkbook book, int card, int entry)
    {
        using var stream = book.ToStream();
        var sheet = XlsxReader.Read(stream, CultureInfo.InvariantCulture).Sheets.Single();
        return sheet.Cards[card].Entries[entry].Style ?? throw new AssertFailedException("The entry has no style.");
    }

    private static X.Font FontWithColor(X.Color color) => new(color, new X.FontSize { Val = 11 }, new X.FontName { Val = "Calibri" });

    [TestMethod]
    public void SolidFillGivesItsColour()
    {
        var book = new TestWorkbook();
        var style = StyleOf(book, book.AddStyle(fill: TestWorkbook.Solid(new X.ForegroundColor { Rgb = "FFFFC000" })));

        Assert.AreEqual(Hex("FFC000"), style.Fill);
        Assert.IsFalse(style.NoFill);
    }

    [TestMethod]
    public void PatternFillOtherThanSolidGivesItsForegroundColour()
    {
        var book = new TestWorkbook();
        var fill = new X.Fill(new X.PatternFill(new X.ForegroundColor { Rgb = "FF00B050" }, new X.BackgroundColor { Rgb = "FFFFFFFF" })
        {
            PatternType = X.PatternValues.DarkGrid,
        });

        Assert.AreEqual(Hex("00B050"), StyleOf(book, book.AddStyle(fill: fill)).Fill);
    }

    [TestMethod]
    public void ThemeFillWithTint()
    {
        var book = new TestWorkbook();
        var style = StyleOf(book, book.AddStyle(fill: TestWorkbook.Solid(new X.ForegroundColor { Theme = 4, Tint = 0.3999755851924192 })));

        Assert.AreEqual(Hex("9DC3E6"), style.Fill);
    }

    [TestMethod]
    public void IndexedFill()
    {
        var book = new TestWorkbook();

        Assert.AreEqual(Hex("FFFF00"), StyleOf(book, book.AddStyle(fill: TestWorkbook.Solid(new X.ForegroundColor { Indexed = 13 }))).Fill);
    }

    [TestMethod]
    public void IndexedFillFromTheStyleSheetPalette()
    {
        var book = new TestWorkbook { IndexedColors = new X.IndexedColors(Enumerable.Range(0, 64).Select(i => new X.RgbColor { Rgb = "00ABCDEF" })) };

        Assert.AreEqual(Hex("ABCDEF"), StyleOf(book, book.AddStyle(fill: TestWorkbook.Solid(new X.ForegroundColor { Indexed = 13 }))).Fill);
    }

    [TestMethod]
    public void NoFill()
    {
        var book = new TestWorkbook();
        var style = StyleOf(book, 0);

        Assert.IsNull(style.Fill);
        Assert.IsTrue(style.NoFill);
    }

    [TestMethod]
    public void PatternNoneIsNoFillWhateverItsColour()
    {
        var book = new TestWorkbook();
        var fill = new X.Fill(new X.PatternFill(new X.ForegroundColor { Rgb = "FFFF0000" }) { PatternType = X.PatternValues.None });
        var style = StyleOf(book, book.AddStyle(fill: fill));

        Assert.IsNull(style.Fill);
        Assert.IsTrue(style.NoFill);
    }

    [TestMethod]
    public void TransparentFillIsNoFill()
    {
        var book = new TestWorkbook();
        var style = StyleOf(book, book.AddStyle(fill: TestWorkbook.Solid(new X.ForegroundColor { Rgb = "00000000" })));

        Assert.IsNull(style.Fill);
        Assert.IsTrue(style.NoFill);
    }

    [TestMethod]
    public void UnresolvableFillIsLeftAsItIs()
    {
        var fills = new[]
        {
            TestWorkbook.Solid(new X.ForegroundColor { Theme = 12 }),
            TestWorkbook.Solid(new X.ForegroundColor { Indexed = 64 }),
            TestWorkbook.Solid(new X.ForegroundColor { Auto = true }),
            new X.Fill(new X.PatternFill { PatternType = X.PatternValues.Solid }),
            new X.Fill(new X.GradientFill(new X.GradientStop(new X.Color { Rgb = "FFFF0000" }) { Position = 0 }, new X.GradientStop(new X.Color { Rgb = "FF0000FF" }) { Position = 1 })),
        };
        foreach (var fill in fills)
        {
            var book = new TestWorkbook();
            var style = StyleOf(book, book.AddStyle(fill: fill));

            Assert.IsNull(style.Fill, fill.OuterXml);
            Assert.IsFalse(style.NoFill, fill.OuterXml);
        }
    }

    [TestMethod]
    public void FontColour()
    {
        var book = new TestWorkbook();
        var style = StyleOf(book, book.AddStyle(font: FontWithColor(new X.Color { Rgb = "FFC00000" })));

        Assert.AreEqual(Hex("C00000"), style.TextColor);
        Assert.IsFalse(style.AutomaticTextColor);
    }

    [TestMethod]
    public void ThemeFontColourComesFromTheWorkbookTheme()
    {
        var book = new TestWorkbook
        {
            ThemeColors = new A.ColorScheme(
                new A.Light1Color(new A.SystemColor { Val = A.SystemColorValues.Window, LastColor = "FEFEFE" }),
                new A.Dark1Color(new A.SystemColor { Val = A.SystemColorValues.WindowText }),
                new A.Light2Color(new A.RgbColorModelHex { Val = "EEECE1" }),
                new A.Dark2Color(new A.RgbColorModelHex { Val = "1F497D" }),
                new A.Accent1Color(new A.RgbColorModelHex { Val = "4F81BD" }))
            { Name = "Made up" },
        };
        var lt1 = book.AddStyle(font: FontWithColor(new X.Color { Theme = 0 }));
        var dk1 = book.AddStyle(font: FontWithColor(new X.Color { Theme = 1 }));
        var dk2 = book.AddStyle(font: FontWithColor(new X.Color { Theme = 3 }));
        var accent1 = book.AddStyle(font: FontWithColor(new X.Color { Theme = 4, Tint = -0.5 }));
        var missing = book.AddStyle(font: FontWithColor(new X.Color { Theme = 7 }));
        book.AddSheet("S").Row("A1", "A", "B", "C", "D", "E").Set("A2", "a", lt1).Set("B2", "b", dk1).Set("C2", "c", dk2).Set("D2", "d", accent1).Set("E2", "e", missing);

        Assert.AreEqual(Hex("FEFEFE"), EntryStyle(book, 0, 0).TextColor, "a system colour gives its last colour");
        Assert.AreEqual(Hex("FFFFFF"), EntryStyle(book, 0, 1).TextColor, "a system colour without a last colour is white");
        Assert.AreEqual(Hex("1F497D"), EntryStyle(book, 0, 2).TextColor);
        Assert.AreEqual(ColorTint.Apply(Hex("4F81BD"), -0.5), EntryStyle(book, 0, 3).TextColor);
        Assert.AreEqual(Hex("FFFFFF"), EntryStyle(book, 0, 4).TextColor, "a slot the theme lacks is white");
    }

    [TestMethod]
    public void WithoutAThemeTheDefaultPaletteIsUsed()
    {
        var book = new TestWorkbook();

        Assert.AreEqual(Hex("44546A"), StyleOf(book, book.AddStyle(font: FontWithColor(new X.Color { Theme = 3 }))).TextColor);
    }

    [TestMethod]
    public void IndexedFontColour()
    {
        var book = new TestWorkbook();

        Assert.AreEqual(Hex("0066CC"), StyleOf(book, book.AddStyle(font: FontWithColor(new X.Color { Indexed = 30 }))).TextColor);
    }

    [TestMethod]
    public void FontWithoutColourIsAutomatic()
    {
        var book = new TestWorkbook();
        var style = StyleOf(book, 0);

        Assert.IsNull(style.TextColor);
        Assert.IsTrue(style.AutomaticTextColor);
    }

    [TestMethod]
    [DataRow(true, 0u)]
    [DataRow(false, 64u)]
    [DataRow(false, 65u)]
    public void AutomaticFontColour(bool auto, uint indexed)
    {
        var book = new TestWorkbook();
        var color = auto ? new X.Color { Auto = true } : new X.Color { Indexed = indexed };
        var style = StyleOf(book, book.AddStyle(font: FontWithColor(color)));

        Assert.IsNull(style.TextColor);
        Assert.IsTrue(style.AutomaticTextColor);
    }

    [TestMethod]
    public void UnresolvableFontColourIsLeftAsItIs()
    {
        foreach (var color in new[] { new X.Color { Theme = 20 }, new X.Color { Indexed = 80 }, new X.Color { Rgb = "00000000" } })
        {
            var book = new TestWorkbook();
            var style = StyleOf(book, book.AddStyle(font: FontWithColor(color)));

            Assert.IsNull(style.TextColor, color.OuterXml);
            Assert.IsFalse(style.AutomaticTextColor, color.OuterXml);
        }
    }

    [TestMethod]
    public void FontAttributes()
    {
        var book = new TestWorkbook();
        var font = new X.Font(
            new X.Bold(),
            new X.Italic(),
            new X.Strike(),
            new X.Underline { Val = X.UnderlineValues.Double },
            new X.FontSize { Val = 14.5 },
            new X.FontName { Val = "Made Up Sans" });
        var style = StyleOf(book, book.AddStyle(font: font)).Font!;

        Assert.AreEqual("Made Up Sans", style.Name);
        Assert.AreEqual(14.5, style.Size);
        Assert.IsTrue(style.Bold);
        Assert.IsTrue(style.Italic);
        Assert.IsTrue(style.Strikethrough);
        Assert.AreEqual("double", style.Underline);
    }

    [TestMethod]
    public void PlainFont()
    {
        var book = new TestWorkbook();
        var style = StyleOf(book, 0).Font!;

        Assert.AreEqual("Calibri", style.Name);
        Assert.AreEqual(11.0, style.Size);
        Assert.IsFalse(style.Bold);
        Assert.IsFalse(style.Italic);
        Assert.IsFalse(style.Strikethrough);
        Assert.IsNull(style.Underline);
    }

    [TestMethod]
    public void FlagsTurnedOffExplicitly()
    {
        var book = new TestWorkbook();
        var font = new X.Font(new X.Bold { Val = false }, new X.Italic { Val = false }, new X.Strike { Val = false }, new X.Underline { Val = X.UnderlineValues.None });
        var style = StyleOf(book, book.AddStyle(font: font)).Font!;

        Assert.IsFalse(style.Bold);
        Assert.IsFalse(style.Italic);
        Assert.IsFalse(style.Strikethrough);
        Assert.IsNull(style.Underline);
    }

    [TestMethod]
    public void BareUnderlineIsSingle()
    {
        var book = new TestWorkbook();

        Assert.AreEqual("single", StyleOf(book, book.AddStyle(font: new X.Font(new X.Underline()))).Font!.Underline);
    }

    [TestMethod]
    public void Alignment()
    {
        var book = new TestWorkbook();
        var alignment = new X.Alignment
        {
            Horizontal = X.HorizontalAlignmentValues.CenterContinuous,
            Vertical = X.VerticalAlignmentValues.Top,
            WrapText = true,
        };
        var style = StyleOf(book, book.AddStyle(alignment: alignment));

        Assert.AreEqual("centerContinuous", style.HorizontalAlignment);
        Assert.AreEqual("top", style.VerticalAlignment);
        Assert.IsTrue(style.WrapText);
    }

    [TestMethod]
    public void WrapTextTurnedOff()
    {
        var book = new TestWorkbook();
        var style = StyleOf(book, book.AddStyle(alignment: new X.Alignment { Horizontal = X.HorizontalAlignmentValues.Right, WrapText = false }));

        Assert.AreEqual("right", style.HorizontalAlignment);
        Assert.IsNull(style.VerticalAlignment);
        Assert.IsFalse(style.WrapText);
    }

    [TestMethod]
    public void NoAlignment()
    {
        var book = new TestWorkbook();
        var style = StyleOf(book, 0);

        Assert.IsNull(style.HorizontalAlignment);
        Assert.IsNull(style.VerticalAlignment);
        Assert.IsNull(style.WrapText);
    }

    [TestMethod]
    public void EmptyCellKeepsItsStyle()
    {
        var book = new TestWorkbook();
        var red = book.AddStyle(fill: TestWorkbook.Solid(new X.ForegroundColor { Rgb = "FFFF0000" }));
        book.AddSheet("S").Row("A1", "A", "B").Set("A2", "x").Empty("B2", red);

        Assert.AreEqual(Hex("FF0000"), EntryStyle(book, 0, 1).Fill);
    }

    [TestMethod]
    public void RowStyleThenColumnStyle()
    {
        var book = new TestWorkbook();
        var red = book.AddStyle(fill: TestWorkbook.Solid(new X.ForegroundColor { Rgb = "FFFF0000" }));
        var green = book.AddStyle(fill: TestWorkbook.Solid(new X.ForegroundColor { Rgb = "FF00FF00" }));
        var blue = book.AddStyle(fill: TestWorkbook.Solid(new X.ForegroundColor { Rgb = "FF0000FF" }));
        book.AddSheet("S")
            .Row("A1", "A", "B", "C")
            .Set("A2", "x").Set("A3", "y")
            .Set("C3", "own", blue)
            .RowStyle(3, red)
            .ColumnStyle(2, 3, green);

        Assert.AreEqual(Hex("00FF00"), EntryStyle(book, 0, 1).Fill, "column style for a cell without its own or a row style");
        Assert.AreEqual(Hex("FF0000"), EntryStyle(book, 1, 1).Fill, "row style before column style");
        Assert.AreEqual(Hex("0000FF"), EntryStyle(book, 1, 2).Fill, "a cell's own style first");
        Assert.IsTrue(EntryStyle(book, 0, 0).NoFill, "no style anywhere: style 0");
    }

    [TestMethod]
    public void StyleIndexOutOfRangeFallsBackToStyleZero()
    {
        var book = new TestWorkbook();
        var style = StyleOf(book, 99);

        Assert.IsTrue(style.NoFill);
        Assert.AreEqual("Calibri", style.Font!.Name);
    }

    [TestMethod]
    public void EveryReadEntryHasAStyle()
    {
        var book = new TestWorkbook();
        book.AddSheet("S").Row("A1", "A", "B", "C").Row("A2", "x", null, 3).Row("A3", null, "y", null);

        using var stream = book.ToStream();
        var sheet = XlsxReader.Read(stream, CultureInfo.InvariantCulture).Sheets.Single();

        Assert.IsTrue(sheet.Cards.SelectMany(c => c.Entries).All(e => e.Style is { NoFill: true, AutomaticTextColor: true }));
    }

    [TestMethod]
    public void StyleTableWithoutAStyleSheet()
    {
        var style = StyleTable.Read(null, null).Style(3);

        Assert.IsTrue(style.NoFill);
        Assert.IsTrue(style.AutomaticTextColor);
        Assert.IsNull(style.Font);
    }

    [TestMethod]
    public void StylesAreShared()
    {
        var table = StyleTable.Read(new X.Stylesheet(new X.Fonts(new X.Font()), new X.Fills(new X.Fill()), new X.CellFormats(new X.CellFormat())), null);

        Assert.AreSame(table.Style(0), table.Style(0));
    }
}
