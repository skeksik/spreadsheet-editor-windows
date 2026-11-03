using System.IO.Compression;
using System.Text;
using SpreadsheetEditor.Core.Imaging;
using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Tests.ExcelCheck;

/// <summary>
/// Checks of the Excel check's picture encoders that need no image decoder: CRC-32, the PNG's pixels, the JPEG's markers and EXIF, and
/// that the core reads each picture's format and size. That a decoder shows them as drawn is checked by opening the samples.
/// </summary>
[TestClass]
public sealed class CheckPicturesTests
{
    [TestMethod]
    public void Crc32MatchesTheCheckValue() =>
        // The check value of CRC-32 (the PNG and zlib polynomial) for the ASCII text "123456789".
        Assert.AreEqual(0xCBF43926u, CheckPictures.Crc32(Encoding.ASCII.GetBytes("123456789")));

    [TestMethod]
    public void APngHoldsItsPixels()
    {
        var pixels = CheckPictures.Draw(40, 90, 2, 3, 7);
        var png = CheckPictures.EncodePng(pixels);

        var picture = PictureFactory.Create(png);
        Assert.AreEqual(ImageFormat.Png, picture.Format);
        Assert.AreEqual(40, picture.PixelWidth);
        Assert.AreEqual(90, picture.PixelHeight);

        // IHDR at 8, then IDAT: its data inflates to one filter byte and 40 RGB pixels per row.
        var idatLength = (png[33] << 24) | (png[34] << 16) | (png[35] << 8) | png[36];
        Assert.AreEqual("IDAT", Encoding.ASCII.GetString(png, 37, 4));
        using var inflated = new MemoryStream();
        using (var zlib = new ZLibStream(new MemoryStream(png, 41, idatLength), CompressionMode.Decompress))
        {
            zlib.CopyTo(inflated);
        }

        var raw = inflated.ToArray();
        Assert.HasCount(90 * ((40 * 3) + 1), raw);
        for (var y = 0; y < 90; y++)
        {
            Assert.AreEqual(0, raw[y * 121], "filter type");
            for (var x = 0; x < 40; x++)
            {
                var at = (y * 121) + 1 + (x * 3);
                Assert.AreEqual(pixels[y, x], new CheckPictures.Rgb(raw[at], raw[at + 1], raw[at + 2]), $"({x}, {y})");
            }
        }
    }

    [TestMethod]
    public void TheRedBandsLieOutsideTheShownRegion()
    {
        // 40 × 90 shown at 2:3 keeps the middle 40 × 60, rows 15–74: 15 rows above and below it, the 2 nearest grey, the rest red.
        var pixels = CheckPictures.Draw(40, 90, 2, 3, 1);
        for (var y = 0; y < 90; y++)
        {
            var red = Enumerable.Range(0, 40).Count(x => pixels[y, x] == CheckPictures.Red);
            var grey = Enumerable.Range(0, 40).Count(x => pixels[y, x] == CheckPictures.Grey);
            Assert.AreEqual(y is < 13 or >= 77 ? 40 : 0, red, $"red in row {y}");
            Assert.AreEqual(y is 13 or 14 or 75 or 76 ? 40 : 0, grey, $"grey in row {y}");
        }

        // 120 × 60 shown at 4:3 keeps the middle 80 × 60, columns 20–99: columns 18–19 and 100–101 grey, the rest outside red.
        var wide = CheckPictures.Draw(120, 60, 4, 3, 4);
        for (var x = 0; x < 120; x++)
        {
            var red = Enumerable.Range(0, 60).Count(y => wide[y, x] == CheckPictures.Red);
            var grey = Enumerable.Range(0, 60).Count(y => wide[y, x] == CheckPictures.Grey);
            Assert.AreEqual(x is < 18 or >= 102 ? 60 : 0, red, $"red in column {x}");
            Assert.AreEqual(x is 18 or 19 or 100 or 101 ? 60 : 0, grey, $"grey in column {x}");
        }

        // A picture of the region's own aspect has neither.
        var fits = CheckPictures.Draw(40, 60, 2, 3, 1);
        Assert.IsFalse(fits.Cast<CheckPictures.Rgb>().Any(p => p == CheckPictures.Red || p == CheckPictures.Grey), "no band");

        Assert.AreEqual((0, 15, 40, 60), CheckPictures.Region(40, 90, 2.0 / 3));
        Assert.AreEqual((20, 0, 80, 60), CheckPictures.Region(120, 60, 4.0 / 3));
        Assert.AreEqual((0, 8, 64, 48), CheckPictures.Region(64, 64, 4.0 / 3));
    }

    [TestMethod]
    public void TheOrientationJpegCarriesExifOrientationSix()
    {
        var plain = CheckPictures.OrientationJpeg(96, 48, null);
        var turned = CheckPictures.OrientationJpeg(96, 48, 6);

        foreach (var jpeg in new[] { plain, turned })
        {
            var picture = PictureFactory.Create(jpeg);
            Assert.AreEqual(ImageFormat.Jpeg, picture.Format);
            Assert.AreEqual(96, picture.PixelWidth);
            Assert.AreEqual(48, picture.PixelHeight);
            Assert.AreEqual(0xFF, jpeg[^2]);
            Assert.AreEqual(0xD9, jpeg[^1]);
        }

        // SOI, APP0 (2 + 16 bytes), then APP1: "Exif\0\0", "MM", 42, IFD0 at 8 with one entry, tag 0x0112, SHORT, count 1, value 6.
        Assert.AreEqual(-1, IndexOf(plain, [0xFF, 0xE1]));
        var app1 = IndexOf(turned, [0xFF, 0xE1]);
        Assert.AreEqual(20, app1);
        byte[] expected = [(byte)'E', (byte)'x', (byte)'i', (byte)'f', 0, 0, (byte)'M', (byte)'M', 0, 42, 0, 0, 0, 8, 0, 1, 0x01, 0x12, 0, 3, 0, 0, 0, 1, 0, 6];
        CollectionAssert.AreEqual(expected, turned.AsSpan(app1 + 4, expected.Length).ToArray());
    }

    private static int IndexOf(byte[] data, byte[] pattern) => data.AsSpan().IndexOf(pattern);
}
