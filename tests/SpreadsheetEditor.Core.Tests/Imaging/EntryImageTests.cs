using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Tests.Imaging;

[TestClass]
public sealed class EntryImageTests
{
    [TestMethod]
    [DataRow(200, 100, true)]
    [DataRow(100, 100, true)]
    [DataRow(100, 200, false)]
    public void WideWhenAtLeastAsWideAsTall(int width, int height, bool wide)
    {
        var image = new EntryImage([], ImageFormat.Png, width, height);

        Assert.IsTrue(image.IsDecodable);
        Assert.AreEqual(wide, image.IsWide);
    }

    [TestMethod]
    [DataRow(ImageFormat.Unknown, 0, 0)]
    [DataRow(ImageFormat.Unknown, 200, 100)]
    [DataRow(ImageFormat.Jpeg, 0, 0)]
    [DataRow(ImageFormat.Jpeg, 200, 0)]
    public void UndecodablePictureIsNotWide(ImageFormat format, int width, int height)
    {
        // Spec §5.2 step 5: a picture that cannot be decoded counts as 1 unit.
        var image = new EntryImage([], format, width, height);

        Assert.IsFalse(image.IsDecodable);
        Assert.IsFalse(image.IsWide);
    }
}
