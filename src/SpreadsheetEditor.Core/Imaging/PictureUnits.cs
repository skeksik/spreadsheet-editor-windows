using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Imaging;

/// <summary>How many horizontal layout units a picture takes in a cell (docs/spec.md §1 "Picture unit", §5.2 step 5, §9.28).</summary>
public static class PictureUnits
{
    /// <summary>True when the picture's format and pixel size are known, so it can be laid out by its shape (<see cref="EntryImage.IsDecodable"/>).</summary>
    public static bool IsDecodable(EntryImage image)
    {
        ArgumentNullException.ThrowIfNull(image);
        return image.IsDecodable;
    }

    /// <summary>True for a decodable picture at least as wide as tall. An undecodable picture is never wide (<see cref="EntryImage.IsWide"/>).</summary>
    public static bool IsWide(EntryImage image)
    {
        ArgumentNullException.ThrowIfNull(image);
        return image.IsWide;
    }

    /// <summary>2 for a wide picture, 1 for a tall or undecodable one.</summary>
    public static int Units(EntryImage image) => IsWide(image) ? 2 : 1;
}
