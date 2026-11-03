using SpreadsheetEditor.Core.Imaging;
using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Tests.Imaging;

[TestClass]
public sealed class PictureUnitsTests
{
    [TestMethod]
    public void TallPictureTakesOneUnit()
    {
        Assert.IsFalse(PictureUnits.IsWide(TestImages.Tall()));
        Assert.AreEqual(1, PictureUnits.Units(TestImages.Tall()));
    }

    [TestMethod]
    public void WidePictureTakesTwoUnits()
    {
        Assert.IsTrue(PictureUnits.IsWide(TestImages.Wide()));
        Assert.AreEqual(2, PictureUnits.Units(TestImages.Wide()));
    }

    [TestMethod]
    public void SquarePictureIsWide() => Assert.AreEqual(2, PictureUnits.Units(TestImages.Square()));

    [TestMethod]
    [DataRow(ImageFormat.Unknown, 40, 30)]
    [DataRow(ImageFormat.Unknown, 30, 40)]
    [DataRow(ImageFormat.Png, 0, 0)]
    [DataRow(ImageFormat.Png, 40, 0)]
    [DataRow(ImageFormat.Jpeg, 0, 40)]
    public void UndecodablePictureTakesOneUnit(ImageFormat format, int width, int height)
    {
        var image = new EntryImage([], format, width, height);
        Assert.IsFalse(PictureUnits.IsDecodable(image));
        Assert.IsFalse(PictureUnits.IsWide(image));
        Assert.AreEqual(1, PictureUnits.Units(image));
    }

    [TestMethod]
    public void PictureWithFormatAndSizeIsDecodable() => Assert.IsTrue(PictureUnits.IsDecodable(TestImages.Tall()));
}
