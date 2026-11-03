using System.Globalization;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using SpreadsheetEditor.Core.Imaging;
using SpreadsheetEditor.Core.Model;
using SpreadsheetEditor.Core.Tests.Imaging;
using SpreadsheetEditor.Core.Xlsx;
using A = DocumentFormat.OpenXml.Drawing;
using X = DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace SpreadsheetEditor.Core.Tests.Xlsx;

/// <summary>
/// Round trips of saves that lay pictures out in their cells (docs/spec.md §5.2–5.4, §9.25–9.29): build a workbook, read it, change the
/// model, write twice (the same file both times), read back and compare with <see cref="PictureLayout"/> and <see cref="CenterCrop"/>.
/// Every value in them is made up.
/// </summary>
[TestClass]
public sealed class XlsxWriterPicturesTests
{
    private static readonly CultureInfo Invariant = CultureInfo.InvariantCulture;

    private static Workbook Read(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        return XlsxReader.Read(stream, Invariant);
    }

    private static byte[] Write(Workbook workbook)
    {
        using var stream = new MemoryStream();
        XlsxWriter.Write(workbook, stream, Invariant);
        return stream.ToArray();
    }

    private static byte[] Change(byte[] original, Action<Workbook> change) => Change(Read(original), change);

    /// <summary>Changes the model, writes twice and checks that both saves are the same file.</summary>
    private static byte[] Change(Workbook workbook, Action<Workbook> change)
    {
        change(workbook);
        var first = Write(workbook);
        CollectionAssert.AreEqual(first, Write(workbook), "Two saves differ.");
        return first;
    }

    /// <summary>A one-cell anchor as written: its cell (1-based), offsets and size in EMU, crop, cNvPr and picture bytes.</summary>
    private sealed record Anchor(
        int Row, int Column, long OffsetX, long OffsetY, long Width, long Height, CropRect Crop, uint Id, string Name, string? Description, byte[] Bytes);

    private static List<Anchor> OneCellAnchors(XlsxPackage package, string sheet)
    {
        var drawing = package.Part(sheet).DrawingsPart!;
        var result = new List<Anchor>();
        foreach (var anchor in drawing.WorksheetDrawing!.Elements<Xdr.OneCellAnchor>())
        {
            var from = anchor.FromMarker!;
            var picture = anchor.GetFirstChild<Xdr.Picture>()!;
            var properties = picture.NonVisualPictureProperties!.NonVisualDrawingProperties!;
            var blip = picture.BlipFill!.Blip!;
            var source = picture.BlipFill.SourceRectangle;
            var crop = source is null
                ? CropRect.None
                : new CropRect(source.Left?.Value ?? 0, source.Top?.Value ?? 0, source.Right?.Value ?? 0, source.Bottom?.Value ?? 0);
            Assert.AreEqual(A.BlipCompressionValues.Print, blip.CompressionState?.Value);
            Assert.AreEqual(A.ShapeTypeValues.Rectangle, picture.ShapeProperties!.GetFirstChild<A.PresetGeometry>()!.Preset!.Value);
            Assert.IsNotNull(picture.ShapeProperties.GetFirstChild<A.Outline>()!.GetFirstChild<A.NoFill>(), "no outline");
            result.Add(new Anchor(
                int.Parse(from.RowId!.Text, Invariant) + 1,
                int.Parse(from.ColumnId!.Text, Invariant) + 1,
                long.Parse(from.ColumnOffset!.Text, Invariant),
                long.Parse(from.RowOffset!.Text, Invariant),
                anchor.Extent!.Cx!.Value,
                anchor.Extent.Cy!.Value,
                crop,
                properties.Id!.Value,
                properties.Name!.Value!,
                properties.Description?.Value,
                Bytes((ImagePart)drawing.GetPartById(blip.Embed!.Value!))));
        }

        return result;
    }

    private static byte[] Bytes(ImagePart part)
    {
        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        using var copy = new MemoryStream();
        stream.CopyTo(copy);
        return copy.ToArray();
    }

    /// <summary>Checks the anchors of a card's value against its layout: one per picture, in order, at the card's first row.</summary>
    private static void AssertLaidOut(IReadOnlyList<Anchor> anchors, Sheet sheet, Card card, int column, int firstRow)
    {
        var placements = PictureLayout.ForCard(PictureLayout.ForSheet(sheet), card).Placements[column];
        Assert.AreEqual(placements.Count, anchors.Count);
        for (var i = 0; i < placements.Count; i++)
        {
            var p = placements[i];
            var a = anchors[i];
            Assert.AreEqual(firstRow + p.RowOffset, a.Row, $"picture {i}: row");
            Assert.AreEqual(sheet.FirstColumn + column, a.Column, $"picture {i}: column");
            Assert.AreEqual(p.OffsetXEmu, a.OffsetX, $"picture {i}: x offset");
            Assert.AreEqual(p.OffsetYEmu, a.OffsetY, $"picture {i}: y offset");
            Assert.AreEqual(p.WidthEmu, a.Width, $"picture {i}: width");
            Assert.AreEqual(p.HeightEmu, a.Height, $"picture {i}: height");
            var crop = p.IsDecodable ? CenterCrop.Compute(p.Image.PixelWidth, p.Image.PixelHeight, p.WidthPx, p.HeightPx) : CropRect.None;
            Assert.AreEqual(crop, a.Crop, $"picture {i}: crop");
            Assert.AreEqual("Image " + a.Id.ToString(Invariant), a.Name);
            Assert.AreEqual("Picture", a.Description);
            CollectionAssert.AreEqual(p.Image.Data, a.Bytes, $"picture {i}: bytes");
        }
    }

    private static double? Height(XlsxPackage package, string sheet, int row) =>
        package.Worksheet(sheet).Descendants<X.Row>().SingleOrDefault(r => r.RowIndex?.Value == (uint)row)?.Height?.Value;

    private static X.Column? ColumnElement(XlsxPackage package, string sheet, int column) =>
        package.Worksheet(sheet).GetFirstChild<X.Columns>()?.Elements<X.Column>().SingleOrDefault(c => c.Min!.Value <= column && c.Max!.Value >= column);

    private static string[] Merges(XlsxPackage package, string sheet) =>
        package.Worksheet(sheet).Descendants<X.MergeCell>().Select(m => m.Reference!.Value!).Order().ToArray();

    /// <summary>Two cards, Fern (row 2) with a picture at B2 and Rose (row 3) without, and a note below the table.</summary>
    private static byte[] Garden(PictureAnchor anchor = PictureAnchor.TwoCell)
    {
        var book = new TestWorkbook();
        book.AddSheet("Plants").Row("A1", "Name", "Photo").Row("A2", "Fern").Row("A3", "Rose").Set("A15", "after")
            .Picture("B2", PictureBytes.Png(40, 30), anchor);
        return book.ToBytes();
    }

    [TestMethod]
    public void PicturesAddedToAValueAreAnchoredAsTheLayoutPlacesThem()
    {
        var wide = PictureFactory.Create(PictureBytes.Png(400, 200));
        var tall = PictureFactory.Create(PictureBytes.Png(100, 300));
        var square = PictureFactory.Create(PictureBytes.Png(50, 50));
        var workbook = Read(Garden());
        var bytes = Change(workbook, w =>
        {
            w.Sheets[0].ImageLayout = new ImageLayout(100, 80, 4);
            w.Sheets[0].Cards[1].Entries[1].Images.AddRange([wide, tall, square]);
        });

        using var after = new XlsxPackage(bytes);
        var sheet = workbook.Sheets[0];
        var anchors = OneCellAnchors(after, "Plants");
        // The changed picture settings lay out Fern's picture again too: it comes first, in card order.
        Assert.AreEqual(4, anchors.Count);
        AssertLaidOut(anchors[..1], sheet, sheet.Cards[0], 1, 2);
        AssertLaidOut(anchors[1..], sheet, sheet.Cards[1], 1, 3);
        // Capacity 4 units of 100 px: wide (2) and tall (1) on the first picture row, the square (wide, 2) on the second, 80 px down.
        Assert.AreEqual((3, 2, 0L, 0L, 200 * 9525L, 80 * 9525L), (anchors[1].Row, anchors[1].Column, anchors[1].OffsetX, anchors[1].OffsetY, anchors[1].Width, anchors[1].Height));
        Assert.AreEqual((200 * 9525L, 0L, 100 * 9525L), (anchors[2].OffsetX, anchors[2].OffsetY, anchors[2].Width));
        Assert.AreEqual((0L, 80 * 9525L), (anchors[3].OffsetX, anchors[3].OffsetY));
        Assert.AreEqual(new CropRect(0, 10000, 0, 10000), anchors[1].Crop, "400 x 200 shown at 200 x 80: 20 px off the top and the bottom");
    }

    [TestMethod]
    public void NewPicturesGetIdsAfterTheDrawingsHighestAndReuseEqualImageParts()
    {
        var fern = PictureBytes.Png(40, 30);
        var bytes = Change(Garden(), w => w.Sheets[0].Cards[1].Entries[1].Images.Add(PictureFactory.Create([.. fern])));

        using var after = new XlsxPackage(bytes);
        var drawing = after.Part("Plants").DrawingsPart!;
        Assert.AreEqual(1, drawing.ImageParts.Count(), "the equal bytes reuse Fern's image part");
        var anchors = OneCellAnchors(after, "Plants");
        Assert.AreEqual(1, anchors.Count);
        Assert.AreEqual(2u, anchors[0].Id, "after the drawing's highest id, 1");
        Assert.AreEqual("Image 2", anchors[0].Name);
        Assert.AreEqual(1, drawing.WorksheetDrawing!.Elements<Xdr.TwoCellAnchor>().Count(), "Fern's picture keeps its anchor");
    }

    [TestMethod]
    public void AnUneditedValueKeepsItsAnchorAsRead()
    {
        var original = Garden();
        var bytes = Change(original, w => w.Sheets[0].Cards[1].Entries[1].Images.Add(PictureFactory.Create(PictureBytes.Png(60, 30))));

        using var before = new XlsxPackage(original);
        using var after = new XlsxPackage(bytes);
        var kept = before.Part("Plants").DrawingsPart!.WorksheetDrawing!.Elements<Xdr.TwoCellAnchor>().Single();
        var now = after.Part("Plants").DrawingsPart!.WorksheetDrawing!.Elements<Xdr.TwoCellAnchor>().Single();
        Assert.AreEqual(kept.OuterXml, now.OuterXml);
        Assert.IsNull(Height(after, "Plants", 2), "Fern's row keeps its height");
        Assert.IsNotNull(Height(after, "Plants", 3), "Rose's row gets the height of its pictures");
    }

    [TestMethod]
    public void ASaveWithoutPictureChangesLeavesTheDrawingAsRead()
    {
        var book = new TestWorkbook();
        book.AddSheet("Plants").Row("A1", "Name", "Photo").Row("A2", "Fern").Row("A3", "Rose").Width(2, 2, 3)
            .Picture("B2", PictureBytes.Png(40, 30)).Picture("B3", PictureBytes.Png(30, 40), PictureAnchor.OneCell)
            .Picture("D9", PictureBytes.Png(5, 5)).Picture("A1", PictureBytes.Png(6, 6), PictureAnchor.Absolute)
            .Picture("B3", PictureBytes.Png(7, 7), PictureAnchor.Grouped);
        var original = book.ToBytes();

        var bytes = Change(original, w => w.Sheets[0].Cards[0].Entries[0].Text = "Bracken");

        using var before = new XlsxPackage(original);
        using var after = new XlsxPackage(bytes);
        Assert.IsTrue(XNode.DeepEquals(XlsxPackage.Xml(before.Part("Plants").DrawingsPart!), XlsxPackage.Xml(after.Part("Plants").DrawingsPart!)));
        Assert.AreEqual(3.0, ColumnElement(after, "Plants", 2)!.Width!.Value, "a column is widened only for pictures laid out again");
        Assert.IsNull(Height(after, "Plants", 2));
    }

    [TestMethod]
    public void ACardGrowsRowsForItsPictures()
    {
        var workbook = Read(Garden());
        var bytes = Change(workbook, w =>
        {
            var sheet = w.Sheets[0];
            sheet.ImageLayout = new ImageLayout(100, 300, 1);
            sheet.Cards[1].Entries[1].Images.AddRange(
                [PictureFactory.Create(PictureBytes.Png(10, 30)), PictureFactory.Create(PictureBytes.Png(20, 30)), PictureFactory.Create(PictureBytes.Png(30, 40))]);
        });

        using var after = new XlsxPackage(bytes);
        var sheet = workbook.Sheets[0];
        var layout = PictureLayout.ForCard(PictureLayout.ForSheet(sheet), sheet.Cards[1]);
        Assert.AreEqual(3, layout.WorksheetRowsNeeded, "one picture row of 300 px per worksheet row");
        AssertLaidOut(OneCellAnchors(after, "Plants")[1..], sheet, sheet.Cards[1], 1, 3);
        AssertSequence(Merges(after, "Plants"), "A3:A5", "B3:B5");
        Assert.AreEqual("after", after.Text("Plants", "A17"), "content below moves down by the rows inserted");
        for (var row = 3; row <= 5; row++)
        {
            Assert.AreEqual(layout.RowHeightsPoints[row - 3], Height(after, "Plants", row));
            Assert.AreEqual((300 + 6) * 0.75, Height(after, "Plants", row), "a full worksheet row adds 6 px");
            Assert.IsTrue(after.Worksheet("Plants").Descendants<X.Row>().Single(r => r.RowIndex!.Value == (uint)row).CustomHeight!.Value);
        }

        var back = Read(bytes).Sheets[0];
        Assert.AreEqual(3, back.Cards[1].SourceRowSpan);
        Assert.AreEqual(3, back.Cards[1].Entries[1].Images.Count);
        Assert.AreEqual(1, back.Cards[0].Entries[1].Images.Count);
    }

    [TestMethod]
    public void PictureColumnsAreOnlyWidenedAndSplitFromTheirRange()
    {
        var book = new TestWorkbook();
        book.AddSheet("Plants").Row("A1", "Name", "Photo", "Leaf").Row("A2", "Fern").Width(1, 3, 5).Width(4, 4, 80);
        var workbook = Read(book.ToBytes());
        var bytes = Change(workbook, w =>
        {
            w.Sheets[0].ImageLayout = new ImageLayout(120, 100, 3);
            w.Sheets[0].Cards[0].Entries[1].Images.Add(PictureFactory.Create(PictureBytes.Png(50, 40)));
        });

        using var after = new XlsxPackage(bytes);
        var target = PictureLayout.ForSheet(workbook.Sheets[0]).Columns[1].WidthTargetPx!.Value;
        Assert.AreEqual(240, target, "one wide picture: 2 units of 120 px");
        var columns = after.Worksheet("Plants").GetFirstChild<X.Columns>()!.Elements<X.Column>()
            .Select(c => (c.Min!.Value, c.Max!.Value, c.Width!.Value)).ToList();
        CollectionAssert.AreEqual(
            new[] { (1u, 1u, 5.0), (2u, 2u, ColumnWidths.PxToWidth(240)), (3u, 3u, 5.0), (4u, 4u, 80.0), (16384u, 16384u, 9.140625) },
            columns);
        Assert.IsTrue(ColumnElement(after, "Plants", 2)!.CustomWidth!.Value);
        Assert.AreEqual((100 * 0.75), Height(after, "Plants", 2), "one picture row of 100 px");
    }

    [TestMethod]
    public void AColumnWiderThanItsPicturesKeepsItsWidth()
    {
        var book = new TestWorkbook();
        book.AddSheet("Plants").Row("A1", "Name", "Photo").Row("A2", "Fern").Width(2, 2, 60);
        var bytes = Change(book.ToBytes(), w => w.Sheets[0].Cards[0].Entries[1].Images.Add(PictureFactory.Create(PictureBytes.Png(50, 40))));

        using var after = new XlsxPackage(bytes);
        Assert.AreEqual(60.0, ColumnElement(after, "Plants", 2)!.Width!.Value);
    }

    [TestMethod]
    public void PicturesMoveWithTheirEntryToAnotherColumn()
    {
        var workbook = Read(Garden());
        var bytes = Change(workbook, w =>
        {
            var sheet = w.Sheets[0];
            sheet.Columns.Reverse();
            foreach (var card in sheet.Cards)
            {
                card.Entries.Reverse();
            }
        });

        using var after = new XlsxPackage(bytes);
        var drawing = after.Part("Plants").DrawingsPart!.WorksheetDrawing!;
        Assert.AreEqual(0, drawing.Elements<Xdr.TwoCellAnchor>().Count(), "the anchor as read is removed");
        var sheet = workbook.Sheets[0];
        AssertLaidOut(OneCellAnchors(after, "Plants"), sheet, sheet.Cards[0], 0, 2);
        Assert.AreEqual(1, after.Part("Plants").DrawingsPart!.ImageParts.Count());
        var back = Read(bytes).Sheets[0];
        Assert.AreEqual("Photo", back.Columns[0].Name);
        Assert.AreEqual(1, back.Cards[0].Entries[0].Images.Count);
        Assert.AreEqual("Fern", back.Cards[0].Entries[1].Text);
    }

    [TestMethod]
    public void PicturesMoveWithTheirCardAndKeepTheirAnchors()
    {
        var original = Garden();
        var bytes = Change(original, w => w.Sheets[0].Cards.Reverse());

        using var after = new XlsxPackage(bytes);
        var anchor = after.Part("Plants").DrawingsPart!.WorksheetDrawing!.Elements<Xdr.TwoCellAnchor>().Single();
        Assert.AreEqual("2", anchor.FromMarker!.RowId!.Text, "row 3, 0-based: moved with Fern");
        var back = Read(bytes).Sheets[0];
        Assert.AreEqual("Rose", back.Cards[0].Entries[0].Text);
        Assert.AreEqual(1, back.Cards[1].Entries[1].Images.Count);
    }

    [TestMethod]
    public void ADuplicatedCardsPicturesAreLaidOutInItsRowsWithTheSameImagePart()
    {
        var workbook = Read(Garden());
        var bytes = Change(workbook, w =>
        {
            var source = w.Sheets[0].Cards[0];
            var copy = new Card { SourceRowSpan = source.SourceRowSpan, CopiedFromRow = source.SourceFirstRow };
            foreach (var entry in source.Entries)
            {
                var entryCopy = new Entry { Text = entry.Text, Original = entry.Original, Style = entry.Style };
                entryCopy.Images.AddRange(entry.Images);
                copy.Entries.Add(entryCopy);
            }

            w.Sheets[0].Cards.Insert(1, copy);
        });

        using var after = new XlsxPackage(bytes);
        var sheet = workbook.Sheets[0];
        AssertLaidOut(OneCellAnchors(after, "Plants"), sheet, sheet.Cards[1], 1, 3);
        Assert.AreEqual(1, after.Part("Plants").DrawingsPart!.ImageParts.Count());
        var back = Read(bytes).Sheets[0];
        Assert.AreEqual(3, back.Cards.Count);
        Assert.AreEqual(1, back.Cards[0].Entries[1].Images.Count);
        Assert.AreEqual(1, back.Cards[1].Entries[1].Images.Count);
        Assert.AreEqual(0, back.Cards[2].Entries[1].Images.Count);
        Assert.AreEqual("after", after.Text("Plants", "A16"));
    }

    [TestMethod]
    public void RemovingAPictureRemovesItsAnchorAndUnusedImagePart()
    {
        var bytes = Change(Garden(), w => w.Sheets[0].Cards[0].Entries[1].Images.Clear());

        using var after = new XlsxPackage(bytes);
        var drawing = after.Part("Plants").DrawingsPart!;
        Assert.AreEqual(0, drawing.WorksheetDrawing!.ChildElements.Count);
        Assert.AreEqual(0, drawing.ImageParts.Count());
        Assert.IsNull(Height(after, "Plants", 2), "a card left without pictures keeps its row height");
    }

    [TestMethod]
    public void PicturesNotOwnedByTheAppStay()
    {
        var book = new TestWorkbook();
        book.AddSheet("Plants").Row("A1", "Name", "Photo").Row("A2", "Fern").Row("A3", "Rose")
            .Picture("B2", PictureBytes.Png(40, 30))
            .Picture("B2", PictureBytes.Png(6, 6), PictureAnchor.Grouped)
            .Picture("B1", PictureBytes.Png(7, 7))
            .Picture("D2", PictureBytes.Png(8, 8));
        var original = book.ToBytes();

        var bytes = Change(original, w => w.Sheets[0].Cards[0].Entries[1].Images.Clear());

        using var before = new XlsxPackage(original);
        using var after = new XlsxPackage(bytes);
        var was = before.Part("Plants").DrawingsPart!.WorksheetDrawing!.ChildElements.Select(e => e.OuterXml).ToList();
        var now = after.Part("Plants").DrawingsPart!.WorksheetDrawing!.ChildElements.Select(e => e.OuterXml).ToList();
        CollectionAssert.AreEqual(was.Skip(1).ToList(), now, "only Fern's own picture goes: the group, the header's and the one beside the table stay");
        Assert.AreEqual(3, after.Part("Plants").DrawingsPart!.ImageParts.Count());
    }

    [TestMethod]
    public void ARemovedEntrysPicturesGoWithIt()
    {
        var book = new TestWorkbook();
        book.AddSheet("Plants").Row("A1", "Name", "Photo").Row("A2", "Fern").Picture("B2", PictureBytes.Png(40, 30));

        var bytes = Change(book.ToBytes(), w =>
        {
            w.Sheets[0].Columns.RemoveAt(1);
            w.Sheets[0].Cards[0].Entries.RemoveAt(1);
        });

        using var after = new XlsxPackage(bytes);
        Assert.AreEqual(0, after.Part("Plants").DrawingsPart!.WorksheetDrawing!.ChildElements.Count);
    }

    [TestMethod]
    public void AnUndecodablePictureIsWrittenAsAOneUnitPictureWithItsBytes()
    {
        var book = new TestWorkbook();
        byte[] unknown = [1, 2, 3, 4, 5];
        book.AddSheet("Plants").Row("A1", "Name", "Photo").Row("A2", "Fern").Row("A3", "Rose").Picture("B2", unknown);
        var workbook = Read(book.ToBytes());
        Assert.IsFalse(workbook.Sheets[0].Cards[0].Entries[1].Images[0].IsDecodable);

        var bytes = Change(workbook, w => w.Sheets[0].Cards[1].Entries[1].Images.Add(w.Sheets[0].Cards[0].Entries[1].Images[0]));

        using var after = new XlsxPackage(bytes);
        var sheet = workbook.Sheets[0];
        var anchors = OneCellAnchors(after, "Plants");
        AssertLaidOut(anchors, sheet, sheet.Cards[1], 1, 3);
        Assert.AreEqual(sheet.ImageLayout.Width * 9525L, anchors[0].Width, "one unit wide");
        Assert.AreEqual(CropRect.None, anchors[0].Crop);
        CollectionAssert.AreEqual(unknown, anchors[0].Bytes);
        Assert.AreEqual(1, after.Part("Plants").DrawingsPart!.ImageParts.Count());
    }

    [TestMethod]
    public void AnUndecodablePictureNotInThePackageIsNotSupported()
    {
        var workbook = Read(Garden());
        workbook.Sheets[0].Cards[1].Entries[1].Images.Add(PictureFactory.Create([9, 9, 9]));

        var error = Assert.ThrowsExactly<SpreadsheetWriteException>(() => Write(workbook));

        Assert.AreEqual(SpreadsheetWriteError.NotSupported, error.Reason);
        StringAssert.Contains(error.Message, "its format is unknown");
    }

    [TestMethod]
    public void ASheetWithoutADrawingGetsOne()
    {
        var book = new TestWorkbook();
        book.AddSheet("Plants").Row("A1", "Name", "Photo").Row("A2", "Fern");
        var jpeg = JpegBuilder.Typical(30, 20).Build();
        var bytes = Change(book.ToBytes(), w => w.Sheets[0].Cards[0].Entries[1].Images.Add(PictureFactory.Create(jpeg)));

        using var after = new XlsxPackage(bytes);
        var part = after.Part("Plants");
        Assert.AreEqual(part.GetIdOfPart(part.DrawingsPart!), after.Worksheet("Plants").GetFirstChild<X.Drawing>()!.Id!.Value);
        var image = part.DrawingsPart!.ImageParts.Single();
        Assert.AreEqual("image/jpeg", image.ContentType);
        CollectionAssert.AreEqual(jpeg, Bytes(image));
        Assert.AreEqual(1u, OneCellAnchors(after, "Plants").Single().Id);
        Assert.AreEqual(1, Read(bytes).Sheets[0].Cards[0].Entries[1].Images.Count);
    }

    [TestMethod]
    public void ANewSheetsPicturesAreLaidOutWithRowsMergesHeightsAndWidths()
    {
        var sheet = new Sheet("Tools") { ImageLayout = new ImageLayout(100, 300, 1) };
        sheet.Columns.AddRange([new Column("Tool"), new Column("Photo") { IsImageColumn = true }]);
        var spade = new Card();
        spade.Entries.AddRange([new Entry { Text = "Spade" }, new Entry()]);
        spade.Entries[1].Images.AddRange([PictureFactory.Create(PictureBytes.Png(30, 60)), PictureFactory.Create(PictureBytes.Png(60, 30))]);
        var rake = new Card();
        rake.Entries.AddRange([new Entry { Text = "Rake" }, new Entry()]);
        var workbook = new Workbook();
        workbook.Sheets.Add(sheet);

        var bytes = Change(workbook, w => w.Sheets[0].Cards.AddRange([spade, rake]));

        using var after = new XlsxPackage(bytes);
        AssertLaidOut(OneCellAnchors(after, "Tools"), sheet, spade, 1, 2);
        AssertSequence(Merges(after, "Tools"), "A2:A3", "B2:B3");
        Assert.AreEqual("Rake", after.Text("Tools", "A4"));
        Assert.AreEqual((300 + 6) * 0.75, Height(after, "Tools", 2));
        Assert.AreEqual(ColumnWidths.PxToWidth(100), ColumnElement(after, "Tools", 2)!.Width!.Value, "capacity 1 unit of 100 px");
        var back = Read(bytes).Sheets[0];
        Assert.AreEqual(2, back.Cards.Count);
        Assert.AreEqual(2, back.Cards[0].Entries[1].Images.Count);
        Assert.AreEqual(2, back.Cards[0].SourceRowSpan);
    }

    [TestMethod]
    public void RemovingPicturesFromADuplicatedSheetLeavesTheOriginalsPictures()
    {
        var bytes = Change(Garden(), w =>
        {
            var source = w.Sheets[0];
            var copy = new Sheet("Plants (2)") { Source = source.Source! with { IsCopy = true }, FirstDataRow = source.FirstDataRow };
            copy.Columns.AddRange(source.Columns.Select(c => new Column(c.Name) { SourceIndex = c.SourceIndex }));
            foreach (var card in source.Cards)
            {
                var cardCopy = new Card { SourceFirstRow = card.SourceFirstRow, SourceRowSpan = card.SourceRowSpan };
                cardCopy.Entries.AddRange(card.Entries.Select(e => new Entry { Text = e.Text, Original = e.Original }));
                copy.Cards.Add(cardCopy);
            }

            copy.Cards[1].Entries[1].Images.Add(PictureFactory.Create(PictureBytes.Png(9, 9)));
            w.Sheets.Add(copy);
        });

        using var after = new XlsxPackage(bytes);
        Assert.AreEqual(1, after.Part("Plants").DrawingsPart!.ImageParts.Count());
        var back = Read(bytes);
        Assert.AreEqual(1, back.Sheets[0].Cards[0].Entries[1].Images.Count);
        Assert.AreEqual(40, back.Sheets[0].Cards[0].Entries[1].Images[0].PixelWidth);
        Assert.AreEqual(0, back.Sheets[1].Cards[0].Entries[1].Images.Count);
        Assert.AreEqual(9, back.Sheets[1].Cards[1].Entries[1].Images.Single().PixelWidth);
    }

    private static void AssertSequence(IEnumerable<string> actual, params string[] expected) =>
        CollectionAssert.AreEqual(expected, actual.ToArray());
}
