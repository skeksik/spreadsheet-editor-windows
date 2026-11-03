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

    /// <summary>A wide picture (at least as wide as tall) takes two layout units in a cell, a tall one one (docs/spec.md).</summary>
    public bool IsWide => PixelWidth >= PixelHeight;
}
