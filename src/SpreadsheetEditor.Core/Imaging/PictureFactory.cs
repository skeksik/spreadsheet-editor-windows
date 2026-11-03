using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Imaging;

/// <summary>Builds entry pictures from file bytes (the xlsx reader, the app's "add pictures").</summary>
public static class PictureFactory
{
    /// <summary>
    /// An <see cref="EntryImage"/> holding <paramref name="data"/> unchanged, with the format and pixel size its header gives. Bytes
    /// that cannot be read give an undecodable picture (unknown format, 0 x 0), which is kept (spec §9.28).
    /// </summary>
    public static EntryImage Create(byte[] data)
    {
        ArgumentNullException.ThrowIfNull(data);
        var facts = PictureHeader.Read(data);
        return new EntryImage(data, facts.Format, facts.PixelWidth, facts.PixelHeight);
    }
}
