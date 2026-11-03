namespace SpreadsheetEditor.Core.Model;

public enum ImageFormat
{
    Unknown,
    Png,
    Jpeg,
    Gif,
    Bmp,
}

/// <summary>A picture in an entry: the file's bytes as read or added, unchanged, and its pixel size.</summary>
public sealed class EntryImage
{
    public EntryImage(byte[] data, ImageFormat format, int pixelWidth, int pixelHeight)
    {
        Data = data;
        Format = format;
        PixelWidth = pixelWidth;
        PixelHeight = pixelHeight;
    }

    public byte[] Data { get; }

    public ImageFormat Format { get; }

    public int PixelWidth { get; }

    public int PixelHeight { get; }

    /// <summary>Whether the header gave a format and a positive size. An undecodable picture keeps its bytes (spec §9.28).</summary>
    public bool IsDecodable => Format != ImageFormat.Unknown && PixelWidth > 0 && PixelHeight > 0;

    /// <summary>
    /// A wide picture (at least as wide as tall) takes two layout units in a cell, a tall one one (docs/spec.md). An undecodable picture
    /// is not wide: it takes 1 unit (spec §5.2 step 5).
    /// </summary>
    public bool IsWide => IsDecodable && PixelWidth >= PixelHeight;
}
