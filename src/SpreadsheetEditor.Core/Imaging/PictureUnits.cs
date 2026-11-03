using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Imaging;

/// <summary>How many horizontal layout units a picture takes in a cell (docs/spec.md §1 "Picture unit", §5.2 step 5, §9.28).</summary>
public static class PictureUnits
{
    /// <summary>True when the picture's format and pixel size are known, so it can be laid out by its shape.</summary>
    public static bool IsDecodable(EntryImage image)
    {
        ArgumentNullException.ThrowIfNull(image);
        return image.Format != ImageFormat.Unknown && image.PixelWidth > 0 && image.PixelHeight > 0;
    }

    /// <summary>True for a decodable picture at least as wide as tall. An undecodable picture is never wide.</summary>
    public static bool IsWide(EntryImage image) => IsDecodable(image) && image.PixelWidth >= image.PixelHeight;

    /// <summary>2 for a wide picture, 1 for a tall or undecodable one.</summary>
    public static int Units(EntryImage image) => IsWide(image) ? 2 : 1;
}
