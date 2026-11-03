using SpreadsheetEditor.Core.Imaging;

namespace SpreadsheetEditor.Core.Tests.Imaging;

[TestClass]
public sealed class CenterCropTests
{
    [TestMethod]
    [DataRow(100, 30, 2, 1, 20000)] // newW 60, crop 40, 20 px each side
    [DataRow(3, 1, 1, 1, 33333)] // 1/3 of the width, truncated
    [DataRow(7, 3, 1, 1, 28571)] // 2/7 = 0.285714..., truncated
    [DataRow(640, 480, 136, 180, 21666)] // a 4:3 photo in a tall slot (default size): 21666.67 truncated
    [DataRow(1000, 500, 272, 180, 12222)] // a 2:1 picture in a wide slot (2w x h)
    public void CropsTooWidePictureLeftAndRight(int width, int height, int targetWidth, int targetHeight, int expected) =>
        Assert.AreEqual(
            new CropRect(expected, 0, expected, 0),
            CenterCrop.Compute(width, height, targetWidth, targetHeight));

    [TestMethod]
    [DataRow(30, 100, 1, 1, 35000)] // newH 30, crop 70, 35 px each side
    [DataRow(1, 3, 1, 1, 33333)]
    [DataRow(4000, 3000, 272, 180, 5882)] // 5882.35 truncated
    [DataRow(3000, 4000, 136, 180, 367)] // 367.6 truncated
    [DataRow(5, 7, 136, 180, 2731)] // 2731.09 truncated
    public void CropsTooTallPictureTopAndBottom(int width, int height, int targetWidth, int targetHeight, int expected) =>
        Assert.AreEqual(
            new CropRect(0, expected, 0, expected),
            CenterCrop.Compute(width, height, targetWidth, targetHeight));

    [TestMethod]
    [DataRow(200, 100, 272, 136)]
    [DataRow(136, 180, 136, 180)]
    [DataRow(1360, 1800, 136, 180)]
    [DataRow(1, 1, 50, 50)]
    public void SameAspectHasNoSourceRectangle(int width, int height, int targetWidth, int targetHeight)
    {
        var crop = CenterCrop.Compute(width, height, targetWidth, targetHeight);

        Assert.AreEqual(CropRect.None, crop);
        Assert.IsTrue(crop.IsNone);
    }

    [TestMethod]
    public void ComparesAspectExactlyForLargeSizes()
    {
        // 2147483646 / 1073741823 is exactly 2: the products overflow int but not long.
        Assert.IsTrue(CenterCrop.Compute(2147483646, 1073741823, 272, 136).IsNone);
        Assert.AreEqual(new CropRect(3433, 0, 3433, 0), CenterCrop.Compute(int.MaxValue, 1000000000, 272, 136));
    }

    [TestMethod]
    [DataRow(0, 100, 136, 180)]
    [DataRow(100, 0, 136, 180)]
    [DataRow(100, 100, 0, 180)]
    [DataRow(100, 100, 136, -1)]
    public void NonPositiveSizeHasNoSourceRectangle(int width, int height, int targetWidth, int targetHeight) =>
        Assert.AreEqual(CropRect.None, CenterCrop.Compute(width, height, targetWidth, targetHeight));
}
