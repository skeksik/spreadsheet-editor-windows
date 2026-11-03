using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using SpreadsheetEditor.Core.Imaging;
using SpreadsheetEditor.Core.Model;
using A = DocumentFormat.OpenXml.Drawing;
using X = DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>
/// Removes and adds the pictures the app lays out in a worksheet's drawing (docs/spec.md §5.4 with §9.28). A picture the app owns is the
/// one the reader shows (<see cref="DrawingPictures"/>): a one- or two-cell anchor whose own child is an <c>xdr:pic</c> embedding an image
/// part. Every other object in the drawing stays as it is.
/// </summary>
internal sealed class DrawingWriter
{
    private const string DrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    private const string MainNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private const string RelationshipsNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    // The worksheet elements a <drawing> follows and precedes (ECMA-376 Part 1 §18.3.1.99, CT_Worksheet; not verified in full).
    private static readonly string[] AfterDrawing =
        ["legacyDrawing", "legacyDrawingHF", "drawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts", "extLst"];

    private readonly WorksheetPart part;
    private readonly string sheetName;
    private DrawingsPart? drawing;
    private List<(string Id, byte[] Bytes)>? images;
    private uint nextId;
    private bool removedAny;

    public DrawingWriter(WorksheetPart part, string sheetName)
    {
        this.part = part;
        this.sheetName = sheetName;
        drawing = part.DrawingsPart;
    }

    /// <summary>
    /// Removes the pictures the app owns whose top-left cell lies in rows <paramref name="firstRow"/> to <paramref name="lastRow"/> and
    /// columns <paramref name="firstColumn"/> to <paramref name="lastColumn"/> (1-based, inclusive).
    /// </summary>
    public void RemoveOwned(int firstRow, int lastRow, int firstColumn, int lastColumn)
    {
        if (drawing?.WorksheetDrawing is not { } root)
        {
            return;
        }

        foreach (var anchor in root.ChildElements.ToList())
        {
            var from = anchor switch
            {
                Xdr.TwoCellAnchor two => two.FromMarker,
                Xdr.OneCellAnchor one => one.FromMarker,
                _ => null,
            };
            if (from is not null
                && anchor.GetFirstChild<Xdr.Picture>()?.BlipFill?.Blip?.Embed?.Value is { Length: > 0 } id
                && drawing.TryGetPartById(id, out var target) && target is ImagePart
                && TryMarker(from.RowId, out var row) && row + 1 >= firstRow && row + 1 <= lastRow
                && TryMarker(from.ColumnId, out var column) && column + 1 >= firstColumn && column + 1 <= lastColumn)
            {
                anchor.Remove();
                removedAny = true;
            }
        }
    }

    /// <summary>
    /// Adds a picture as a one-cell anchor at (<paramref name="row"/>, <paramref name="column"/>), 1-based, with the placement's offsets
    /// and size, centre-cropped to its display aspect (§5.4 steps 3–4). An undecodable picture is shown at its slot size, uncropped (§9.28).
    /// </summary>
    public void Add(PicturePlacement placement, int row, int column)
    {
        var root = Root();
        var embed = ImageId(placement.Image);
        if (nextId == 0)
        {
            nextId = root.Descendants<Xdr.NonVisualDrawingProperties>().Select(p => p.Id?.Value ?? 0).DefaultIfEmpty(0u).Max() + 1;
        }

        var id = nextId++;
        var image = placement.Image;
        var crop = placement.IsDecodable
            ? CenterCrop.Compute(image.PixelWidth, image.PixelHeight, placement.WidthPx, placement.HeightPx)
            : CropRect.None;

        var blipFill = new Xdr.BlipFill(new A.Blip { Embed = embed, CompressionState = A.BlipCompressionValues.Print });
        if (!crop.IsNone)
        {
            var source = new A.SourceRectangle();
            source.Left = crop.Left != 0 ? crop.Left : null;
            source.Top = crop.Top != 0 ? crop.Top : null;
            source.Right = crop.Right != 0 ? crop.Right : null;
            source.Bottom = crop.Bottom != 0 ? crop.Bottom : null;
            blipFill.AppendChild(source);
        }

        blipFill.AppendChild(new A.Stretch(new A.FillRectangle()));

        var anchor = new Xdr.OneCellAnchor(
            new Xdr.FromMarker(
                new Xdr.ColumnId(Text(column - 1)),
                new Xdr.ColumnOffset(Text(placement.OffsetXEmu)),
                new Xdr.RowId(Text(row - 1)),
                new Xdr.RowOffset(Text(placement.OffsetYEmu))),
            new Xdr.Extent { Cx = placement.WidthEmu, Cy = placement.HeightEmu },
            new Xdr.Picture(
                new Xdr.NonVisualPictureProperties(
                    new Xdr.NonVisualDrawingProperties
                    {
                        Id = id,
                        Name = "Image " + id.ToString(CultureInfo.InvariantCulture),
                        Description = "Picture",
                    },
                    new Xdr.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true })),
                blipFill,
                new Xdr.ShapeProperties(
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle },
                    new A.Outline(new A.NoFill()))),
            new Xdr.ClientData());
        root.AppendChild(anchor);
    }

    /// <summary>Deletes the image parts no picture uses any more, once pictures were removed.</summary>
    public void Finish()
    {
        if (!removedAny || drawing?.WorksheetDrawing is not { } root)
        {
            return;
        }

        var used = root.Descendants<A.Blip>().Select(b => b.Embed?.Value).OfType<string>().ToHashSet(StringComparer.Ordinal);
        foreach (var pair in drawing.Parts.Where(p => p.OpenXmlPart is ImagePart && !used.Contains(p.RelationshipId)).ToList())
        {
            drawing.DeletePart(pair.RelationshipId);
        }
    }

    private Xdr.WorksheetDrawing Root()
    {
        if (drawing?.WorksheetDrawing is { } existing)
        {
            return existing;
        }

        if (drawing is null)
        {
            var id = RelationshipIds.Next(part);
            drawing = part.AddNewPart<DrawingsPart>(id);
            var worksheet = part.Worksheet!;
            var element = new X.Drawing { Id = id };
            var before = worksheet.ChildElements.FirstOrDefault(e => AfterDrawing.Contains(e.LocalName));
            if (before is null)
            {
                worksheet.AppendChild(element);
            }
            else
            {
                worksheet.InsertBefore(element, before);
            }
        }

        var root = new Xdr.WorksheetDrawing();
        root.AddNamespaceDeclaration("xdr", DrawingNamespace);
        root.AddNamespaceDeclaration("a", MainNamespace);
        root.AddNamespaceDeclaration("r", RelationshipsNamespace);
        drawing.WorksheetDrawing = root;
        return root;
    }

    /// <summary>The relationship id of an image part holding the picture's bytes: an equal one is reused, else one is added.</summary>
    private string ImageId(EntryImage image)
    {
        images ??= drawing?.Parts.Where(p => p.OpenXmlPart is ImagePart).Select(p => (p.RelationshipId, Bytes((ImagePart)p.OpenXmlPart))).ToList()
            ?? [];
        foreach (var (id, bytes) in images)
        {
            if (bytes.AsSpan().SequenceEqual(image.Data))
            {
                return id;
            }
        }

        var contentType = image.Format switch
        {
            ImageFormat.Png => "image/png",
            ImageFormat.Jpeg => "image/jpeg",
            ImageFormat.Gif => "image/gif",
            ImageFormat.Bmp => "image/bmp",
            _ => throw SpreadsheetWriteException.NotSupported(
                $"A picture on the sheet \"{sheetName}\" cannot be saved: its format is unknown."),
        };
        var target = drawing!;
        var newId = RelationshipIds.Next(target);
        var imagePart = target.AddNewPart<ImagePart>(contentType, newId);
        using (var stream = new MemoryStream(image.Data, writable: false))
        {
            imagePart.FeedData(stream);
        }

        images.Add((newId, image.Data));
        return newId;
    }

    private static byte[] Bytes(ImagePart image)
    {
        using var stream = image.GetStream(FileMode.Open, FileAccess.Read);
        using var copy = new MemoryStream();
        stream.CopyTo(copy);
        return copy.ToArray();
    }

    private static string Text(long value) => value.ToString(CultureInfo.InvariantCulture);

    private static bool TryMarker(OpenXmlLeafTextElement? element, out int value) =>
        int.TryParse(element?.Text, NumberStyles.None, CultureInfo.InvariantCulture, out value);
}
