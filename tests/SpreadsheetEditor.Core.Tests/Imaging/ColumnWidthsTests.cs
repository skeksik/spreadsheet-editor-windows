using SpreadsheetEditor.Core.Imaging;

namespace SpreadsheetEditor.Core.Tests.Imaging;

[TestClass]
public sealed class ColumnWidthsTests
{
    private const double Delta = 1e-9;

    [TestMethod]
    [DataRow(0.0, 0.0)]
    [DataRow(6.0, 0.5)]
    [DataRow(12.0, 1.0)]
    [DataRow(19.0, 2.0)]
    [DataRow(68.0, 9.0)]
    [DataRow(1762.0, 251.0)]
    public void ConvertsPixelsToWidth(double px, double width) => Assert.AreEqual(width, ColumnWidths.PxToWidth(px), Delta);

    [TestMethod]
    [DataRow(0.0, 0.0)]
    [DataRow(0.5, 6.0)]
    [DataRow(1.0, 12.0)]
    [DataRow(2.0, 19.0)]
    [DataRow(8.43, 64.01)]
    [DataRow(251.0, 1762.0)]
    public void ConvertsWidthToPixels(double width, double px) => Assert.AreEqual(px, ColumnWidths.WidthToPx(width), Delta);

    [TestMethod]
    [DataRow(1.0)]
    [DataRow(6.0)]
    [DataRow(12.0)]
    [DataRow(13.0)]
    [DataRow(64.0)]
    [DataRow(136.0)]
    [DataRow(1088.0)]
    [DataRow(1762.0)]
    public void PixelsSurviveARoundTrip(double px) => Assert.AreEqual(px, ColumnWidths.WidthToPx(ColumnWidths.PxToWidth(px)), Delta);

    [TestMethod]
    public void WidensADefaultColumn()
    {
        var width = ColumnWidths.Widen(null, 136);
        Assert.IsNotNull(width);
        Assert.AreEqual(131.0 / 7, width!.Value, Delta);
    }

    [TestMethod]
    public void NeverNarrowsAColumn()
    {
        Assert.IsNull(ColumnWidths.Widen(null, 64));
        Assert.IsNull(ColumnWidths.Widen(9, 68));
        Assert.IsNull(ColumnWidths.Widen(30, 136));
    }

    [TestMethod]
    public void CapsTheTargetAt1762Pixels()
    {
        var width = ColumnWidths.Widen(10, 2720);
        Assert.IsNotNull(width);
        Assert.AreEqual(251.0, width!.Value, Delta);
        Assert.AreEqual(1762.0, ColumnWidths.WidthToPx(width!.Value), Delta);
    }

    [TestMethod]
    public void LeavesAColumnWiderThanTheCapAlone() => Assert.IsNull(ColumnWidths.Widen(300, 2720));
}
