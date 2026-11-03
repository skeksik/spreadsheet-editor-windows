using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Imaging;

/// <summary>
/// Lays out the pictures of a sheet's cards in their cells for the .xlsx writer (docs/spec.md §5.2–5.4 with §9.25 and §9.26). Pure
/// functions over the model: they change nothing and return what the writer applies.
/// </summary>
public static class PictureLayout
{
    /// <summary>The tallest a full worksheet row of pictures is, in pixels, before the 6 px a full row adds (§9.26).</summary>
    public const int MaxPictureRowsHeightPx = 540;

    /// <summary>The pixels a worksheet row holding its full number of picture rows gets on top of the pictures (§5.3 step 5).</summary>
    public const int FullRowExtraPx = 6;

    /// <summary>Points per pixel at 96 dpi, as row heights are written (§5.3 step 5).</summary>
    public const double PointsPerPixel = 0.75;

    /// <summary>The sheet-wide layout: pictures per row after the cap, picture rows per worksheet row, each column's capacity.</summary>
    public static SheetPictureLayout ForSheet(Sheet sheet)
    {
        ArgumentNullException.ThrowIfNull(sheet);
        var settings = sheet.ImageLayout;
        var width = settings.Width;
        var height = settings.Height;
        ArgumentOutOfRangeException.ThrowIfLessThan(width, 1, nameof(sheet));
        ArgumentOutOfRangeException.ThrowIfLessThan(height, 1, nameof(sheet));
        ArgumentOutOfRangeException.ThrowIfLessThan(settings.ImagesPerRow, 1, nameof(sheet));

        var perRow = CapPicturesPerRow(settings.ImagesPerRow, width);
        var columns = new ColumnPictureLayout[sheet.Columns.Count];
        for (var c = 0; c < columns.Length; c++)
        {
            var maxUnits = 0;
            foreach (var card in sheet.Cards)
            {
                if (c < card.Entries.Count)
                {
                    maxUnits = Math.Max(maxUnits, card.Entries[c].Images.Sum(image => PictureUnits.Units(image)));
                }
            }

            var capacity = maxUnits > 0 ? Math.Max(1, Math.Min(maxUnits, perRow)) : perRow;
            columns[c] = new ColumnPictureLayout(maxUnits, capacity, maxUnits > 0 ? capacity * width : null);
        }

        return new SheetPictureLayout(width, height, perRow, PictureRowsPerWorksheetRow(height), columns);
    }

    /// <summary>Where the card's pictures go, how many worksheet rows it needs and their heights.</summary>
    public static CardPictureLayout ForCard(SheetPictureLayout sheet, Card card)
    {
        ArgumentNullException.ThrowIfNull(sheet);
        ArgumentNullException.ThrowIfNull(card);

        var r = sheet.PictureRowsPerWorksheetRow;
        var placements = new IReadOnlyList<PicturePlacement>[sheet.Columns.Count];
        var pictureRows = 1;
        var hasPictures = false;
        for (var c = 0; c < placements.Length; c++)
        {
            IReadOnlyList<EntryImage> images = c < card.Entries.Count ? card.Entries[c].Images : [];
            var placed = Place(images, c, sheet.Columns[c].Capacity, sheet.Width, sheet.Height, r);
            placements[c] = placed;
            if (placed.Length > 0)
            {
                hasPictures = true;
                pictureRows = Math.Max(pictureRows, placed[^1].VisualRow + 1);
            }
        }

        var pictureWorksheetRows = (pictureRows + r - 1) / r;
        var need = Math.Max(card.SourceRowSpan, pictureWorksheetRows);
        var heights = new double?[need];
        for (var i = 0; i < need; i++)
        {
            var k = Math.Min(r, pictureRows - (r * i));
            heights[i] = hasPictures && k > 0 ? RowHeightPoints(k, r, sheet.Height) : null;
        }

        return new CardPictureLayout(hasPictures, pictureRows, pictureWorksheetRows, need, heights, placements);
    }

    /// <summary>
    /// Pictures per row after the cap: when <paramref name="perRow"/> × <paramref name="width"/> exceeds <see cref="ColumnWidths.MaxPx"/>,
    /// max(1, floor(1762 / width)) (§5.2 step 1, §2.8).
    /// </summary>
    public static int CapPicturesPerRow(int perRow, int width) =>
        (long)perRow * width > ColumnWidths.MaxPx ? Math.Max(1, ColumnWidths.MaxPx / width) : perRow;

    /// <summary>How many picture rows one worksheet row holds: max(1, floor(540 / h)) (§9.26).</summary>
    public static int PictureRowsPerWorksheetRow(int height) => Math.Max(1, MaxPictureRowsHeightPx / height);

    /// <summary>
    /// The height in points of a worksheet row holding <paramref name="pictureRows"/> picture rows of <paramref name="height"/> px, out of
    /// <paramref name="rowsPerWorksheetRow"/>: a full row adds 6 px (§5.3 step 5).
    /// </summary>
    public static double RowHeightPoints(int pictureRows, int rowsPerWorksheetRow, int height) =>
        ((pictureRows * height) + (pictureRows == rowsPerWorksheetRow ? FullRowExtraPx : 0)) * PointsPerPixel;

    // The placement loop of §5.4 step 2 with §9.25: a picture wraps to the next picture row only when the current row already holds
    // pictures, so a picture wider than the capacity still goes on an empty row. Counting rows runs this same loop.
    private static PicturePlacement[] Place(IReadOnlyList<EntryImage> images, int column, int capacity, int width, int height, int r)
    {
        var result = new PicturePlacement[images.Count];
        var visualRow = 0;
        var units = 0;
        for (var i = 0; i < images.Count; i++)
        {
            var image = images[i];
            var wide = PictureUnits.IsWide(image);
            var u = wide ? 2 : 1;
            if (units > 0 && units + u > capacity)
            {
                visualRow++;
                units = 0;
            }

            result[i] = new PicturePlacement(
                image,
                i,
                column,
                visualRow,
                visualRow / r,
                units * width,
                (visualRow % r) * height,
                wide ? 2 * width : width,
                height,
                wide,
                PictureUnits.IsDecodable(image));
            units += u;
        }

        return result;
    }
}
