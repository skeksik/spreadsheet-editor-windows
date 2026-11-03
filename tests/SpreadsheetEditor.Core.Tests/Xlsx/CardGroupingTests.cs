using SpreadsheetEditor.Core.Xlsx;

namespace SpreadsheetEditor.Core.Tests.Xlsx;

[TestClass]
public sealed class CardGroupingTests
{
    private static List<(int FirstRow, int RowSpan)> Group(int firstRow, int lastRow, params MergeRange[] merges) =>
        CardGrouping.Group(firstRow, lastRow, 1, 3, merges);

    [TestMethod]
    public void OneCardPerRowWithoutMerges() =>
        CollectionAssert.AreEqual(new List<(int, int)> { (2, 1), (3, 1), (4, 1) }, Group(2, 4));

    [TestMethod]
    public void VerticalMergeMakesOneCard() =>
        CollectionAssert.AreEqual(new List<(int, int)> { (2, 3), (5, 1) }, Group(2, 5, new MergeRange(2, 1, 4, 1)));

    [TestMethod]
    public void OverlappingMergesChain() =>
        CollectionAssert.AreEqual(
            new List<(int, int)> { (2, 4) },
            Group(2, 5, new MergeRange(4, 3, 5, 3), new MergeRange(2, 1, 3, 1), new MergeRange(3, 2, 4, 2)));

    [TestMethod]
    public void MergeFromAboveTheFirstCardExtendsIt() =>
        CollectionAssert.AreEqual(new List<(int, int)> { (2, 2), (4, 1) }, Group(2, 4, new MergeRange(1, 1, 3, 1)));

    [TestMethod]
    public void HorizontalMergesDoNotGroup() =>
        CollectionAssert.AreEqual(new List<(int, int)> { (2, 1), (3, 1) }, Group(2, 3, new MergeRange(2, 1, 2, 3)));

    [TestMethod]
    public void MergesOutsideOrAcrossTheTableColumnsDoNotGroup() =>
        CollectionAssert.AreEqual(
            new List<(int, int)> { (2, 1), (3, 1) },
            Group(2, 3, new MergeRange(2, 4, 3, 4), new MergeRange(2, 3, 3, 4)));

    [TestMethod]
    public void LastCardMayReachBelowTheLastRow() =>
        CollectionAssert.AreEqual(new List<(int, int)> { (2, 1), (3, 3) }, Group(2, 3, new MergeRange(3, 1, 5, 1)));
}
