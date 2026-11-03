using System.Buffers.Binary;

namespace SpreadsheetEditor.Core.Tests.Imaging;

/// <summary>Builds made-up picture headers in code. Only the header fields the reader looks at carry meaning; pixel data is absent.</summary>
internal static class PictureBytes
{
    public static byte[] Png(uint width, uint height, string firstChunk = "IHDR", uint chunkLength = 13)
    {
        var bytes = new List<byte> { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };
        AddUInt32BigEndian(bytes, chunkLength);
        bytes.AddRange(firstChunk.Select(c => (byte)c));
        AddUInt32BigEndian(bytes, width);
        AddUInt32BigEndian(bytes, height);
        Append(bytes, 8, 6, 0, 0, 0); // bit depth, colour type, compression, filter, interlace
        Append(bytes, 0, 0, 0, 0); // CRC, not checked by the reader
        return [.. bytes];
    }

    /// <summary>The PNG bytes the reader needs: signature, chunk length and type, and the 13 IHDR bytes (not the CRC).</summary>
    public const int PngHeaderLength = 29;

    public static byte[] Gif(string version, ushort width, ushort height)
    {
        var bytes = new List<byte>();
        bytes.AddRange(("GIF" + version).Select(c => (byte)c));
        AddUInt16LittleEndian(bytes, width);
        AddUInt16LittleEndian(bytes, height);
        Append(bytes, 0, 0, 0); // packed fields, background colour index, pixel aspect ratio
        bytes.Add(0x3B); // trailer
        return [.. bytes];
    }

    /// <summary>Header and logical screen descriptor.</summary>
    public const int GifHeaderLength = 13;

    public static byte[] BmpCore(ushort width, ushort height)
    {
        var bytes = BmpFileHeader();
        AddUInt32LittleEndian(bytes, 12);
        AddUInt16LittleEndian(bytes, width);
        AddUInt16LittleEndian(bytes, height);
        AddUInt16LittleEndian(bytes, 1); // planes
        AddUInt16LittleEndian(bytes, 24); // bits per pixel
        return [.. bytes];
    }

    public const int BmpCoreHeaderLength = 14 + 12;

    /// <summary>A BITMAPINFOHEADER (size 40) or a later header that starts the same way (V4 108, V5 124), zero-padded to its size.</summary>
    public static byte[] BmpInfo(int width, int height, uint headerSize = 40)
    {
        var bytes = BmpFileHeader();
        AddUInt32LittleEndian(bytes, headerSize);
        AddUInt32LittleEndian(bytes, unchecked((uint)width));
        AddUInt32LittleEndian(bytes, unchecked((uint)height));
        AddUInt16LittleEndian(bytes, 1); // planes
        AddUInt16LittleEndian(bytes, 24); // bits per pixel
        while (bytes.Count < 14 + (int)headerSize)
        {
            bytes.Add(0);
        }

        return [.. bytes];
    }

    /// <summary>The BMP bytes the reader needs for an info header: the file header and the first 40 bytes of the DIB header.</summary>
    public const int BmpInfoHeaderLength = 14 + 40;

    public static void Append(List<byte> bytes, params ReadOnlySpan<byte> values) => bytes.AddRange(values);

    public static void AddUInt16BigEndian(List<byte> bytes, ushort value)
    {
        Span<byte> buffer = stackalloc byte[2];
        BinaryPrimitives.WriteUInt16BigEndian(buffer, value);
        bytes.AddRange(buffer);
    }

    private static List<byte> BmpFileHeader()
    {
        var bytes = new List<byte> { (byte)'B', (byte)'M' };
        bytes.AddRange(new byte[12]); // file size, reserved, pixel data offset: not read
        return bytes;
    }

    private static void AddUInt16LittleEndian(List<byte> bytes, ushort value)
    {
        Span<byte> buffer = stackalloc byte[2];
        BinaryPrimitives.WriteUInt16LittleEndian(buffer, value);
        bytes.AddRange(buffer);
    }

    private static void AddUInt32BigEndian(List<byte> bytes, uint value)
    {
        Span<byte> buffer = stackalloc byte[4];
        BinaryPrimitives.WriteUInt32BigEndian(buffer, value);
        bytes.AddRange(buffer);
    }

    private static void AddUInt32LittleEndian(List<byte> bytes, uint value)
    {
        Span<byte> buffer = stackalloc byte[4];
        BinaryPrimitives.WriteUInt32LittleEndian(buffer, value);
        bytes.AddRange(buffer);
    }
}
