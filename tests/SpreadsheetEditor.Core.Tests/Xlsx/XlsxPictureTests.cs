using System.Globalization;
using SpreadsheetEditor.Core.Model;
using SpreadsheetEditor.Core.Tests.Imaging;
using SpreadsheetEditor.Core.Xlsx;

namespace SpreadsheetEditor.Core.Tests.Xlsx;

/// <summary>Pictures anchored to cells, read from <c>.xlsx</c> (docs/spec.md §4.1 steps 1 and 6, §4.2 step 3, §9.28).</summary>
[TestClass]
public sealed class XlsxPictureTests
{
    private static readonly byte[] Wide = PictureBytes.Png(40, 30);
    private static readonly byte[] Tall = PictureBytes.Png(30, 40);
    private static readonly byte[] Square = PictureBytes.Png(25, 25);
    private static readonly byte[] Garbage = [1, 2, 3, 4, 5, 6, 7, 8];

    private static Sheet ReadSingle(TestWorkbook book)
    {
        using var stream = book.ToStream();
        return XlsxReader.Read(stream, CultureInfo.InvariantCulture).Sheets.Single();
    }

    private static List<byte[]> Pictures(Entry entry) => [.. entry.Images.Select(i => i.Data)];

    [TestMethod]
    [DataRow(PictureAnchor.TwoCell)]
    [DataRow(PictureAnchor.OneCell)]
    public void PictureGoesToTheEntryOfItsAnchorCell(PictureAnchor anchor)
    {
        var book = new TestWorkbook();
        book.AddSheet("S").Row("A1", "Name", "Photo").Row("A2", "Fern", "caption").Row("A3", "Rose").Picture("B2", Wide, anchor);

        var sheet = ReadSingle(book);

        var image = sheet.Cards[0].Entries[1].Images.Single();
        CollectionAssert.AreEqual(Wide, image.Data, "the bytes are kept unchanged");
        Assert.AreEqual((ImageFormat.Png, 40, 30), (image.Format, image.PixelWidth, image.PixelHeight));
        Assert.AreEqual("caption", sheet.Cards[0].Entries[1].Text);
        Assert.IsTrue(sheet.Columns[1].IsImageColumn);
        Assert.IsFalse(sheet.Columns[0].IsImageColumn);
        Assert.AreEqual(0, sheet.Cards[1].Entries[1].Images.Count);
        Assert.AreEqual(0, sheet.Cards[0].Entries[0].Images.Count);
    }

    [TestMethod]
    [DataRow(PictureAnchor.Absolute)]
    [DataRow(PictureAnchor.Grouped)]
    [DataRow(PictureAnchor.Linked)]
    public void IgnoredPicturesAreNotShownAndStayInThePackage(PictureAnchor anchor)
    {
        var book = new TestWorkbook();
        book.AddSheet("S").Row("A1", "Name", "Photo").Row("A2", "Fern").Picture("B2", Wide, anchor);
        var bytes = book.ToBytes();

        using var stream = new MemoryStream(bytes);
        var workbook = XlsxReader.Read(stream, CultureInfo.InvariantCulture);
        var sheet = workbook.Sheets.Single();

        CollectionAssert.AreEqual(bytes, workbook.Source!.Package.ToArray(), "the package is kept as read");
        Assert.IsTrue(sheet.Cards.SelectMany(c => c.Entries).All(e => e.Images.Count == 0));
        Assert.IsFalse(sheet.Columns[1].IsImageColumn);
    }

    [TestMethod]
    public void UndecodablePictureIsKept()
    {
        var book = new TestWorkbook();
        book.AddSheet("S").Row("A1", "Name", "Photo").Row("A2", "Fern").Picture("B2", Garbage);

        var image = ReadSingle(book).Cards[0].Entries[1].Images.Single();

        CollectionAssert.AreEqual(Garbage, image.Data);
        Assert.IsFalse(image.IsDecodable);
    }

    [TestMethod]
    public void PicturesOfAMultiRowCardGoInRowOrderThenFileOrder()
    {
        var book = new TestWorkbook();
        book.AddSheet("S")
            .Row("A1", "Name", "Photo")
            .Row("A2", "Fern")
            .Merge("A2:A4")
            .Row("A5", "Rose")
            .Picture("B4", Square)
            .Picture("B2", Wide)
            .Picture("B3", Garbage)
            .Picture("B2", Tall)
            .Picture("B5", Square);

        var sheet = ReadSingle(book);

        Assert.AreEqual(2, sheet.Cards.Count);
        var pictures = Pictures(sheet.Cards[0].Entries[1]);
        Assert.AreEqual(4, pictures.Count);
        CollectionAssert.AreEqual(Wide, pictures[0]);
        CollectionAssert.AreEqual(Tall, pictures[1]);
        CollectionAssert.AreEqual(Garbage, pictures[2]);
        CollectionAssert.AreEqual(Square, pictures[3]);
        Assert.AreEqual(1, sheet.Cards[1].Entries[1].Images.Count);
    }

    [TestMethod]
    public void PictureAloneMakesItsCellOccupied()
    {
        var book = new TestWorkbook();
        book.AddSheet("S").Row("A1", "Name", "Photo").Row("A2", "Fern").Picture("B5", Wide);

        var sheet = ReadSingle(book);

        Assert.AreEqual(4, sheet.Cards.Count, "rows 2 to 5: the picture's row ends the table");
        Assert.AreEqual(1, sheet.Cards[3].Entries[1].Images.Count);
        Assert.AreEqual(5, sheet.Source!.LastRow);
    }

    [TestMethod]
    public void PictureColumnWidensTheColumnBlock()
    {
        var book = new TestWorkbook();
        book.AddSheet("S").Row("A1", "Name", "Size").Row("A2", "Fern", 3).Picture("D2", Wide);

        var sheet = ReadSingle(book);

        Assert.AreEqual(4, sheet.Columns.Count);
        Assert.AreEqual("", sheet.Columns[3].Name);
        Assert.IsTrue(sheet.Columns[3].IsImageColumn);
        Assert.AreEqual(1, sheet.Cards[0].Entries[3].Images.Count);
    }

    [TestMethod]
    public void PicturesOutsideTheTableAreNotShown()
    {
        var book = new TestWorkbook();
        book.AddSheet("S")
            .Row("A1", "Name", "Photo")
            .Row("A2", "Fern")
            .Picture("H2", Wide)
            .Picture("B1", Tall);

        var sheet = ReadSingle(book);

        Assert.AreEqual(2, sheet.Columns.Count, "a picture more than 3 columns away is outside the block");
        Assert.IsTrue(sheet.Cards.SelectMany(c => c.Entries).All(e => e.Images.Count == 0), "a picture in the header row belongs to no card");
        Assert.IsFalse(sheet.Columns[1].IsImageColumn);
    }

    [TestMethod]
    public void PictureInTheLastColumnIsIgnored()
    {
        var book = new TestWorkbook();
        book.AddSheet("S").Row("A1", "Name").Row("A2", "Fern").Picture("XFD2", Wide);

        var sheet = ReadSingle(book);

        Assert.AreEqual(1, sheet.Columns.Count);
    }

    [TestMethod]
    public void SheetWithOnlyAPictureHasATable()
    {
        var book = new TestWorkbook();
        book.AddSheet("S").Picture("C3", Wide);

        var sheet = ReadSingle(book);

        Assert.AreEqual(3, sheet.FirstColumn);
        Assert.AreEqual(3, sheet.HeaderRow);
        Assert.AreEqual(1, sheet.Cards.Count);
        Assert.AreEqual(0, sheet.Cards[0].Entries[0].Images.Count, "the anchor row is the header row");
    }

    [TestMethod]
    public void PicturesWithHiddenHeadersStartOnTheFirstRow()
    {
        var book = new TestWorkbook();
        book.AddSheet("S").Row("A1", 1, 2).Row("A2", 3, 4).Picture("B1", Wide);

        var sheet = ReadSingle(book);

        Assert.IsNull(sheet.HeaderRow);
        Assert.AreEqual(1, sheet.Cards[0].Entries[1].Images.Count);
    }
}
