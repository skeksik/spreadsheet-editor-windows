using SpreadsheetEditor.Core.Model;
using SpreadsheetEditor.Core.Xlsx;

namespace SpreadsheetEditor.Core.Tests.Xlsx;

[TestClass]
public sealed class LayoutMetadataTests
{
    [TestMethod]
    public void ReadsAllThree() =>
        Assert.AreEqual(new ImageLayout(200, 150, 5), LayoutMetadata.Read("ImageHeight:150", "ExcelImagesPerRow:5", "ImageWidth:200"));

    [TestMethod]
    public void MissingCellsKeepDefaults() => Assert.AreEqual(ImageLayout.Default, LayoutMetadata.Read(null, null, null));

    [TestMethod]
    public void FirstFailureStopsReading()
    {
        // Row 1 bad: nothing is read, even though rows 2 and 3 are good.
        Assert.AreEqual(ImageLayout.Default, LayoutMetadata.Read("ImageHeight:tall", "ExcelImagesPerRow:5", "ImageWidth:200"));
        // Row 2 bad: the height is read, pictures per row and width keep their defaults.
        Assert.AreEqual(new ImageLayout(136, 150, 8), LayoutMetadata.Read("ImageHeight:150", "ImagesPerRow:5", "ImageWidth:200"));
        // The prefix must be exact, colon included.
        Assert.AreEqual(ImageLayout.Default, LayoutMetadata.Read("imageheight:150", null, null));
        Assert.AreEqual(ImageLayout.Default, LayoutMetadata.Read("ImageHeight 150", null, null));
    }

    [TestMethod]
    public void ClampsToTheSettingsRanges()
    {
        Assert.AreEqual(new ImageLayout(1000, 500, 1), LayoutMetadata.Read("ImageHeight:9999", "ExcelImagesPerRow:9", "ImageWidth:5000"));
        Assert.AreEqual(new ImageLayout(50, 50, 1), LayoutMetadata.Read("ImageHeight:-3", "ExcelImagesPerRow:0", "ImageWidth:10"));
    }

    [TestMethod]
    public void ClampsImagesPerRowToTheMaximumForTheWidth()
    {
        // floor(1762 / 300) = 5.
        Assert.AreEqual(new ImageLayout(300, 180, 5), LayoutMetadata.Read("ImageHeight:180", "ExcelImagesPerRow:20", "ImageWidth:300"));
        // With the default width 136: floor(1762 / 136) = 12.
        Assert.AreEqual(new ImageLayout(136, 180, 12), LayoutMetadata.Read("ImageHeight:180", "ExcelImagesPerRow:40", null));
    }

    [TestMethod]
    [DataRow("ImageHeight:180", true)]
    [DataRow("ExcelImagesPerRow:8", true)]
    [DataRow("ImageWidth:136", true)]
    [DataRow("ImageWidth:", false)]
    [DataRow("ImageWidth:wide", false)]
    [DataRow("Image width: 136", false)]
    [DataRow(null, false)]
    public void RecognisesMetadataStrings(string? text, bool expected) => Assert.AreEqual(expected, LayoutMetadata.IsMetadata(text));
}
