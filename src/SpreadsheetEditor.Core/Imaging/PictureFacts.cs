using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Imaging;

/// <summary>What a picture's header says: its format and pixel size. The default value is "undecodable".</summary>
public readonly record struct PictureFacts(ImageFormat Format, int PixelWidth, int PixelHeight)
{
    /// <summary>A picture whose header cannot be read: unknown format, 0 x 0. Spec §9.28 keeps its bytes and lays it out as 1 unit.</summary>
    public static PictureFacts Undecodable => default;

    public bool IsDecodable => Format != ImageFormat.Unknown && PixelWidth > 0 && PixelHeight > 0;
}
