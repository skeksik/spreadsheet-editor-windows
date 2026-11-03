using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using SpreadsheetEditor.Core.Imaging;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace SpreadsheetEditor.Core.Xlsx;

/// <summary>
/// Reads the pictures anchored to cells in a worksheet's drawing (docs/spec.md §4.1 steps 1 and 6): an <c>xdr:twoCellAnchor</c> or
/// <c>xdr:oneCellAnchor</c> whose own child is an <c>xdr:pic</c> with an embedded <c>a:blip</c>. Pictures with an absolute position
/// (<c>xdr:absoluteAnchor</c>), pictures inside a group (<c>xdr:grpSp</c>) and linked pictures (<c>r:link</c> without <c>r:embed</c>) are
/// not read; they stay in the package as they are. The bytes are kept unchanged; bytes that cannot be decoded still give a picture
/// (§9.28).
/// </summary>
internal static class DrawingPictures
{
    /// <summary>The worksheet's anchored pictures in file order.</summary>
    public static List<AnchoredPicture> Read(WorksheetPart part, CancellationToken cancellationToken)
    {
        var pictures = new List<AnchoredPicture>();
        var drawing = part.DrawingsPart;
        if (drawing?.WorksheetDrawing is not { } root)
        {
            return pictures;
        }
        foreach (var anchor in root.ChildElements)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var from = anchor switch
            {
                Xdr.TwoCellAnchor two => two.FromMarker,
                Xdr.OneCellAnchor one => one.FromMarker,
                _ => null,
            };
            if (from is null
                || anchor.GetFirstChild<Xdr.Picture>()?.BlipFill?.Blip?.Embed?.Value is not { Length: > 0 } id
                || !TryMarker(from.RowId, out var row)
                || !TryMarker(from.ColumnId, out var column)
                || !drawing.TryGetPartById(id, out var imagePart)
                || imagePart is not ImagePart image)
            {
                continue;
            }
            pictures.Add(new AnchoredPicture(row + 1, Math.Min(column + 1, CellReference.MaxColumn), PictureFactory.Create(Bytes(image))));
        }
        return pictures;
    }

    /// <summary>A marker's 0-based row or column.</summary>
    private static bool TryMarker(OpenXmlLeafTextElement? element, out int value) =>
        int.TryParse(element?.Text, NumberStyles.None, CultureInfo.InvariantCulture, out value);

    private static byte[] Bytes(ImagePart image)
    {
        using var stream = image.GetStream(FileMode.Open, FileAccess.Read);
        using var copy = new MemoryStream();
        stream.CopyTo(copy);
        return copy.ToArray();
    }
}
