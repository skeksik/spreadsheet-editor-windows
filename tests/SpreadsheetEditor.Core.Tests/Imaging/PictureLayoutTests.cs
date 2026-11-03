using System.Globalization;
using SpreadsheetEditor.Core.Imaging;
using SpreadsheetEditor.Core.Model;
using static SpreadsheetEditor.Core.Tests.Imaging.TestImages;

namespace SpreadsheetEditor.Core.Tests.Imaging;

[TestClass]
public sealed class PictureLayoutTests
{
    private const double Delta = 1e-9;

    /// <summary>The tallest row a worksheet allows: 409.5 points, 546 pixels.</summary>
    private const double MaxRowPoints = 409.5;

    [TestMethod]
    [DataRow(8, 136, 8)]
    [DataRow(12, 136, 12)]
    [DataRow(13, 136, 12)]
    [DataRow(8, 300, 5)]
    [DataRow(1, 1000, 1)]
    [DataRow(3, 1000, 1)]
    [DataRow(35, 50, 35)]
    [DataRow(36, 50, 35)]
    public void CapsPicturesPerRowAt1762Pixels(int perRow, int width, int capped) =>
        Assert.AreEqual(capped, PictureLayout.CapPicturesPerRow(perRow, width));

    [TestMethod]
    public void SheetLayoutUsesTheCappedPicturesPerRow()
    {
        var sheet = MakeSheet(new ImageLayout(300, 180, 8), 1, [[[]]]);
        var layout = PictureLayout.ForSheet(sheet);
        Assert.AreEqual(5, layout.PicturesPerRow);
        Assert.AreEqual(5, layout.Columns[0].Capacity);
    }

    [TestMethod]
    [DataRow(50, 10)]
    [DataRow(180, 3)]
    [DataRow(181, 2)]
    [DataRow(270, 2)]
    [DataRow(271, 1)]
    [DataRow(500, 1)]
    public void PictureRowsPerWorksheetRowFollowsTheHeight(int height, int rows) =>
        Assert.AreEqual(rows, PictureLayout.PictureRowsPerWorksheetRow(height));

    [TestMethod]
    public void EveryPictureHeightFitsTheTallestWorksheetRow()
    {
        for (var height = 50; height <= 500; height++)
        {
            var r = PictureLayout.PictureRowsPerWorksheetRow(height);
            var at = "h = " + height.ToString(CultureInfo.InvariantCulture);
            Assert.IsTrue(r >= 1, at);
            Assert.IsTrue((r * height) + 6 <= 546, at);

            // Enough tall pictures in a one-unit column to fill several worksheet rows, the last one partly.
            var sheet = MakeSheet(new ImageLayout(136, height, 1), 1, [[[.. Enumerable.Range(0, (2 * r) + 1).Select(_ => Tall())]]]);
            var layout = PictureLayout.ForSheet(sheet);
            var card = PictureLayout.ForCard(layout, sheet.Cards[0]);
            Assert.AreEqual((2 * r) + 1, card.PictureRows);
            Assert.AreEqual(3, card.WorksheetRowsNeeded);
            foreach (var points in card.RowHeightsPoints)
            {
                Assert.IsNotNull(points);
                Assert.IsTrue(points!.Value <= MaxRowPoints, at);
            }
        }
    }

    [TestMethod]
    public void TallWideTallInCapacityTwoTakesThreePictureRows()
    {
        var sheet = MakeSheet(new ImageLayout(136, 180, 2), 1, [[[Tall(), Wide(), Tall()]]]);
        var layout = PictureLayout.ForSheet(sheet);
        Assert.AreEqual(4, layout.Columns[0].MaxUnits);
        Assert.AreEqual(2, layout.Columns[0].Capacity);

        var card = PictureLayout.ForCard(layout, sheet.Cards[0]);
        Assert.IsTrue(card.HasPictures);
        Assert.AreEqual(3, card.PictureRows);
        Assert.AreEqual(1, card.WorksheetRowsNeeded);
        Assert.AreEqual(MaxRowPoints, card.RowHeightsPoints[0]!.Value, Delta);

        var placed = card.Placements[0];
        Assert.HasCount(3, placed);
        AssertPlacement(placed[0], visualRow: 0, rowOffset: 0, x: 0, y: 0, width: 136, height: 180);
        AssertPlacement(placed[1], visualRow: 1, rowOffset: 0, x: 0, y: 180, width: 272, height: 180);
        AssertPlacement(placed[2], visualRow: 2, rowOffset: 0, x: 0, y: 360, width: 136, height: 180);
    }

    [TestMethod]
    public void WidePictureInCapacityOneLeavesNoEmptyRow()
    {
        var sheet = MakeSheet(new ImageLayout(136, 180, 1), 1, [[[Wide()]], [[Wide(), Wide()]]]);
        var layout = PictureLayout.ForSheet(sheet);
        Assert.AreEqual(1, layout.Columns[0].Capacity);

        var one = PictureLayout.ForCard(layout, sheet.Cards[0]);
        Assert.AreEqual(1, one.PictureRows);
        AssertPlacement(one.Placements[0][0], visualRow: 0, rowOffset: 0, x: 0, y: 0, width: 272, height: 180);
        Assert.AreEqual(180 * 0.75, one.RowHeightsPoints[0]!.Value, Delta);

        var two = PictureLayout.ForCard(layout, sheet.Cards[1]);
        Assert.AreEqual(2, two.PictureRows);
        AssertPlacement(two.Placements[0][1], visualRow: 1, rowOffset: 0, x: 0, y: 180, width: 272, height: 180);
    }

    [TestMethod]
    public void PicturesShareARowUpToTheCapacity()
    {
        var sheet = MakeSheet(new ImageLayout(100, 180, 4), 1, [[[Tall(), Wide(), Tall(), Wide()]]]);
        var layout = PictureLayout.ForSheet(sheet);
        var placed = PictureLayout.ForCard(layout, sheet.Cards[0]).Placements[0];
        AssertPlacement(placed[0], visualRow: 0, rowOffset: 0, x: 0, y: 0, width: 100, height: 180);
        AssertPlacement(placed[1], visualRow: 0, rowOffset: 0, x: 100, y: 0, width: 200, height: 180);
        AssertPlacement(placed[2], visualRow: 0, rowOffset: 0, x: 300, y: 0, width: 100, height: 180);
        AssertPlacement(placed[3], visualRow: 1, rowOffset: 0, x: 0, y: 180, width: 200, height: 180);
    }

    [TestMethod]
    public void ColumnCapacityAndWidthTargetFollowTheLargestValue()
    {
        var sheet = MakeSheet(
            new ImageLayout(136, 180, 8),
            4,
            [[[], [Tall()], [Tall(), Tall(), Tall(), Tall(), Tall(), Tall(), Tall(), Tall(), Tall(), Tall()], [Tall()]],
            [[], [Wide(), Tall()], [], [UnknownFormat(), NoSize()]]]);
        var layout = PictureLayout.ForSheet(sheet);

        Assert.AreEqual(new ColumnPictureLayout(0, 8, null), layout.Columns[0]);
        Assert.AreEqual(new ColumnPictureLayout(3, 3, 3 * 136), layout.Columns[1]);
        Assert.AreEqual(new ColumnPictureLayout(10, 8, 8 * 136), layout.Columns[2]);
        Assert.AreEqual(new ColumnPictureLayout(2, 2, 2 * 136), layout.Columns[3]);
    }

    [TestMethod]
    public void PictureRowsBeyondOneWorksheetRowMoveTheAnchorDown()
    {
        var sheet = MakeSheet(new ImageLayout(136, 180, 1), 1, [[[Tall(), Tall(), Tall(), Tall(), Tall(), Tall(), Tall()]]]);
        var layout = PictureLayout.ForSheet(sheet);
        Assert.AreEqual(3, layout.PictureRowsPerWorksheetRow);

        var card = PictureLayout.ForCard(layout, sheet.Cards[0]);
        Assert.AreEqual(7, card.PictureRows);
        Assert.AreEqual(3, card.PictureWorksheetRows);
        Assert.AreEqual(3, card.WorksheetRowsNeeded);
        Assert.HasCount(3, card.RowHeightsPoints);
        Assert.AreEqual(MaxRowPoints, card.RowHeightsPoints[0]!.Value, Delta);
        Assert.AreEqual(MaxRowPoints, card.RowHeightsPoints[1]!.Value, Delta);
        Assert.AreEqual(180 * 0.75, card.RowHeightsPoints[2]!.Value, Delta);

        var placed = card.Placements[0];
        AssertPlacement(placed[2], visualRow: 2, rowOffset: 0, x: 0, y: 360, width: 136, height: 180);
        AssertPlacement(placed[4], visualRow: 4, rowOffset: 1, x: 0, y: 180, width: 136, height: 180);
        AssertPlacement(placed[6], visualRow: 6, rowOffset: 2, x: 0, y: 0, width: 136, height: 180);
    }

    [TestMethod]
    public void PictureRowsNeededIsTheLargestOverTheCardsColumns()
    {
        var sheet = MakeSheet(new ImageLayout(136, 300, 1), 2, [[[Tall()], [Tall(), Tall()]]]);
        var layout = PictureLayout.ForSheet(sheet);
        Assert.AreEqual(1, layout.PictureRowsPerWorksheetRow);

        var card = PictureLayout.ForCard(layout, sheet.Cards[0]);
        Assert.AreEqual(2, card.PictureRows);
        Assert.AreEqual(2, card.WorksheetRowsNeeded);
        Assert.AreEqual((300 + 6) * 0.75, card.RowHeightsPoints[0]!.Value, Delta);
        Assert.AreEqual((300 + 6) * 0.75, card.RowHeightsPoints[1]!.Value, Delta);
        Assert.AreEqual(1, card.Placements[1][1].RowOffset);
        Assert.AreEqual(0, card.Placements[1][1].OffsetYPx);
    }

    [TestMethod]
    public void RowsBeyondThePicturesKeepTheirHeight()
    {
        var sheet = MakeSheet(new ImageLayout(136, 180, 8), 1, [[[Tall()]]]);
        sheet.Cards[0].SourceRowSpan = 4;
        var layout = PictureLayout.ForSheet(sheet);

        var card = PictureLayout.ForCard(layout, sheet.Cards[0]);
        Assert.AreEqual(1, card.PictureWorksheetRows);
        Assert.AreEqual(4, card.WorksheetRowsNeeded);
        Assert.HasCount(4, card.RowHeightsPoints);
        Assert.AreEqual(180 * 0.75, card.RowHeightsPoints[0]!.Value, Delta);
        Assert.IsNull(card.RowHeightsPoints[1]);
        Assert.IsNull(card.RowHeightsPoints[2]);
        Assert.IsNull(card.RowHeightsPoints[3]);
    }

    [TestMethod]
    public void CardWithoutPicturesKeepsItsRows()
    {
        var sheet = MakeSheet(new ImageLayout(136, 180, 8), 2, [[[], []]]);
        sheet.Cards[0].SourceRowSpan = 2;
        var layout = PictureLayout.ForSheet(sheet);

        var card = PictureLayout.ForCard(layout, sheet.Cards[0]);
        Assert.IsFalse(card.HasPictures);
        Assert.AreEqual(1, card.PictureRows);
        Assert.AreEqual(1, card.PictureWorksheetRows);
        Assert.AreEqual(2, card.WorksheetRowsNeeded);
        Assert.IsNull(card.RowHeightsPoints[0]);
        Assert.IsNull(card.RowHeightsPoints[1]);
        Assert.IsEmpty(card.Placements[0]);
        Assert.IsEmpty(card.Placements[1]);
    }

    [TestMethod]
    public void CardWithFewerEntriesThanColumnsHasNoPicturesThere()
    {
        var sheet = MakeSheet(new ImageLayout(136, 180, 8), 3, [[[Tall()]]]);
        var layout = PictureLayout.ForSheet(sheet);
        var card = PictureLayout.ForCard(layout, sheet.Cards[0]);
        Assert.HasCount(3, card.Placements);
        Assert.HasCount(1, card.Placements[0]);
        Assert.IsEmpty(card.Placements[2]);
    }

    [TestMethod]
    public void UndecodablePicturesAreLaidOutAsOneUnit()
    {
        var sheet = MakeSheet(new ImageLayout(136, 180, 2), 1, [[[UnknownFormat(), NoSize(), Wide()]]]);
        var layout = PictureLayout.ForSheet(sheet);
        Assert.AreEqual(4, layout.Columns[0].MaxUnits);

        var card = PictureLayout.ForCard(layout, sheet.Cards[0]);
        var placed = card.Placements[0];
        Assert.AreEqual(2, card.PictureRows);
        AssertPlacement(placed[0], visualRow: 0, rowOffset: 0, x: 0, y: 0, width: 136, height: 180);
        AssertPlacement(placed[1], visualRow: 0, rowOffset: 0, x: 136, y: 0, width: 136, height: 180);
        AssertPlacement(placed[2], visualRow: 1, rowOffset: 0, x: 0, y: 180, width: 272, height: 180);
        Assert.IsFalse(placed[0].IsDecodable);
        Assert.IsFalse(placed[0].IsWide);
        Assert.IsFalse(placed[1].IsDecodable);
        Assert.IsTrue(placed[2].IsDecodable);
        Assert.IsTrue(placed[2].IsWide);
    }

    [TestMethod]
    public void PlacementKnowsItsPictureAndPosition()
    {
        var second = Tall();
        var sheet = MakeSheet(new ImageLayout(136, 180, 8), 2, [[[], [Tall(), second]]]);
        var layout = PictureLayout.ForSheet(sheet);
        var placed = PictureLayout.ForCard(layout, sheet.Cards[0]).Placements[1][1];
        Assert.AreSame(second, placed.Image);
        Assert.AreEqual(1, placed.ImageIndex);
        Assert.AreEqual(1, placed.ColumnIndex);
    }

    [TestMethod]
    public void OffsetsAndSizesInEmuArePixelsTimes9525()
    {
        var sheet = MakeSheet(new ImageLayout(136, 180, 1), 1, [[[Tall(), Wide()]]]);
        var layout = PictureLayout.ForSheet(sheet);
        var placed = PictureLayout.ForCard(layout, sheet.Cards[0]).Placements[0];

        Assert.AreEqual(0L, placed[0].OffsetXEmu);
        Assert.AreEqual(1_295_400L, placed[0].WidthEmu);
        Assert.AreEqual(1_714_500L, placed[0].HeightEmu);
        Assert.AreEqual(1_714_500L, placed[1].OffsetYEmu);
        Assert.AreEqual(2_590_800L, placed[1].WidthEmu);

        var sideBySide = MakeSheet(new ImageLayout(136, 180, 2), 1, [[[Tall(), Tall()]]]);
        var second = PictureLayout.ForCard(PictureLayout.ForSheet(sideBySide), sideBySide.Cards[0]).Placements[0][1];
        Assert.AreEqual(1_295_400L, second.OffsetXEmu);
    }

    private static void AssertPlacement(PicturePlacement placement, int visualRow, int rowOffset, int x, int y, int width, int height)
    {
        Assert.AreEqual(visualRow, placement.VisualRow, "visual row");
        Assert.AreEqual(rowOffset, placement.RowOffset, "row offset");
        Assert.AreEqual(x, placement.OffsetXPx, "x");
        Assert.AreEqual(y, placement.OffsetYPx, "y");
        Assert.AreEqual(width, placement.WidthPx, "width");
        Assert.AreEqual(height, placement.HeightPx, "height");
    }
}
