using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Imaging;

/// <summary>
/// Where one picture goes: a one-cell anchor in the card's rows, its offsets inside that cell and its display size (docs/spec.md §5.4).
/// </summary>
/// <param name="Image">The picture.</param>
/// <param name="ImageIndex">Its position in the value's picture list.</param>
/// <param name="ColumnIndex">The table column (0-based); the worksheet column is the sheet's first column plus this.</param>
/// <param name="VisualRow">The picture row within the value, from 0.</param>
/// <param name="RowOffset">The anchor row, counted in worksheet rows below the card's first row.</param>
/// <param name="OffsetXPx">Horizontal offset from the anchor cell's left edge, in pixels.</param>
/// <param name="OffsetYPx">Vertical offset from the anchor cell's top edge, in pixels.</param>
/// <param name="WidthPx">Display width in pixels: 2w for a wide picture, w otherwise.</param>
/// <param name="HeightPx">Display height in pixels: h.</param>
/// <param name="IsWide">True when the picture takes 2 units.</param>
/// <param name="IsDecodable">False for a picture whose format or size is unknown; it is laid out as a 1-unit picture.</param>
public sealed record PicturePlacement(
    EntryImage Image,
    int ImageIndex,
    int ColumnIndex,
    int VisualRow,
    int RowOffset,
    int OffsetXPx,
    int OffsetYPx,
    int WidthPx,
    int HeightPx,
    bool IsWide,
    bool IsDecodable)
{
    /// <summary>English Metric Units per pixel, as the drawing part measures offsets and sizes (docs/spec.md §5.4 step 4).</summary>
    public const long EmuPerPixel = 9525;

    public long OffsetXEmu => OffsetXPx * EmuPerPixel;

    public long OffsetYEmu => OffsetYPx * EmuPerPixel;

    public long WidthEmu => WidthPx * EmuPerPixel;

    public long HeightEmu => HeightPx * EmuPerPixel;
}
