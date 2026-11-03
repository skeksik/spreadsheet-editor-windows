using SpreadsheetEditor.Core.Imaging;
using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Tests.Imaging;

[TestClass]
public sealed class PictureFactoryTests
{
    [TestMethod]
    public void BuildsEntryImageFromHeader()
    {
        var bytes = PictureBytes.Png(300, 400);

        var image = PictureFactory.Create(bytes);

        Assert.AreSame(bytes, image.Data);
        Assert.AreEqual(ImageFormat.Png, image.Format);
        Assert.AreEqual(300, image.PixelWidth);
        Assert.AreEqual(400, image.PixelHeight);
        Assert.IsTrue(image.IsDecodable);
        Assert.IsFalse(image.IsWide);
    }

    [TestMethod]
    public void KeepsBytesOfUndecodablePicture()
    {
        var bytes = "garbage"u8.ToArray();

        var image = PictureFactory.Create(bytes);

        Assert.AreSame(bytes, image.Data);
        CollectionAssert.AreEqual("garbage"u8.ToArray(), image.Data);
        Assert.AreEqual(ImageFormat.Unknown, image.Format);
        Assert.AreEqual(0, image.PixelWidth);
        Assert.AreEqual(0, image.PixelHeight);
        Assert.IsFalse(image.IsDecodable);
    }

    [TestMethod]
    public void RejectsNull() => Assert.ThrowsExactly<ArgumentNullException>(() => PictureFactory.Create(null!));
}
