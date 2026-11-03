using System.Buffers.Binary;
using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Imaging;

/// <summary>
/// Reads a picture's format and pixel size from its header without decoding pixels (ADR 0002, "Image facts in the core"). Unknown,
/// truncated or corrupt input gives <see cref="PictureFacts.Undecodable"/>; it never throws.
/// </summary>
public static class PictureHeader
{
    public static PictureFacts Read(ReadOnlySpan<byte> data)
    {
        var facts = ReadPng(data) ?? ReadJpeg(data) ?? ReadGif(data) ?? ReadBmp(data) ?? PictureFacts.Undecodable;
        return facts.IsDecodable ? facts : PictureFacts.Undecodable;
    }

    // PNG (W3C PNG Third Edition): the 8-byte signature, then IHDR as the first chunk: length 13, type "IHDR", width and height as
    // big-endian four-byte unsigned integers limited to 0 .. 2^31-1, zero invalid. The CRC is not checked.
    private static PictureFacts? ReadPng(ReadOnlySpan<byte> data)
    {
        ReadOnlySpan<byte> signature = [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];
        ReadOnlySpan<byte> ihdr = [(byte)'I', (byte)'H', (byte)'D', (byte)'R'];
        if (!data.StartsWith(signature))
        {
            return null;
        }

        const int HeaderEnd = 8 + 4 + 4 + 13;
        if (data.Length < HeaderEnd
            || BinaryPrimitives.ReadUInt32BigEndian(data[8..]) != 13
            || !data.Slice(12, 4).SequenceEqual(ihdr))
        {
            return PictureFacts.Undecodable;
        }

        var width = BinaryPrimitives.ReadUInt32BigEndian(data[16..]);
        var height = BinaryPrimitives.ReadUInt32BigEndian(data[20..]);
        return width > int.MaxValue || height > int.MaxValue
            ? PictureFacts.Undecodable
            : new PictureFacts(ImageFormat.Png, (int)width, (int)height);
    }

    // JPEG (ITU-T T.81, B.1.1.2, Table B.1, B.2.2): SOI, then marker segments. Any marker may be preceded by X'FF' fill bytes. SOI, EOI,
    // RSTm and TEM stand alone; every other marker starts a segment with a big-endian length that counts itself. The frame header
    // follows a SOFn marker (X'FFC0'-X'FFCF' except DHT X'FFC4', JPG X'FFC8' and DAC X'FFCC'): length, precision P, lines Y, samples
    // per line X. Y = 0 (lines defined by a later DNL marker) is undecodable here. A scan (SOS) or EOI before any SOFn is undecodable.
    private static PictureFacts? ReadJpeg(ReadOnlySpan<byte> data)
    {
        if (data.Length < 2 || data[0] != 0xFF || data[1] != 0xD8)
        {
            return null;
        }

        var pos = 2;
        while (true)
        {
            if (pos >= data.Length || data[pos] != 0xFF)
            {
                return PictureFacts.Undecodable;
            }

            while (pos < data.Length && data[pos] == 0xFF)
            {
                pos++;
            }

            if (pos >= data.Length)
            {
                return PictureFacts.Undecodable;
            }

            var marker = data[pos++];
            if (marker == 0x01 || marker is >= 0xD0 and <= 0xD7)
            {
                continue;
            }

            if (marker is 0x00 or 0xD8 or 0xD9 or 0xDA || pos + 2 > data.Length)
            {
                return PictureFacts.Undecodable;
            }

            var length = BinaryPrimitives.ReadUInt16BigEndian(data[pos..]);
            if (length < 2 || pos + length > data.Length)
            {
                return PictureFacts.Undecodable;
            }

            if (IsStartOfFrame(marker))
            {
                // Lf (2), P (1), Y (2), X (2), Nf (1).
                if (length < 8)
                {
                    return PictureFacts.Undecodable;
                }

                var height = BinaryPrimitives.ReadUInt16BigEndian(data[(pos + 3)..]);
                var width = BinaryPrimitives.ReadUInt16BigEndian(data[(pos + 5)..]);
                return new PictureFacts(ImageFormat.Jpeg, width, height);
            }

            pos += length;
        }
    }

    private static bool IsStartOfFrame(byte marker) => marker is >= 0xC0 and <= 0xCF and not (0xC4 or 0xC8 or 0xCC);

    // GIF (GIF89a specification, sections 4, 17 and 18): "GIF" and version "87a" or "89a", then the logical screen descriptor (7
    // bytes) whose first fields are the logical screen width and height, unsigned, least significant byte first.
    private static PictureFacts? ReadGif(ReadOnlySpan<byte> data)
    {
        ReadOnlySpan<byte> gif87 = [(byte)'G', (byte)'I', (byte)'F', (byte)'8', (byte)'7', (byte)'a'];
        ReadOnlySpan<byte> gif89 = [(byte)'G', (byte)'I', (byte)'F', (byte)'8', (byte)'9', (byte)'a'];
        if (!data.StartsWith(gif87) && !data.StartsWith(gif89))
        {
            return null;
        }

        const int DescriptorEnd = 6 + 7;
        return data.Length < DescriptorEnd
            ? PictureFacts.Undecodable
            : new PictureFacts(
                ImageFormat.Gif,
                BinaryPrimitives.ReadUInt16LittleEndian(data[6..]),
                BinaryPrimitives.ReadUInt16LittleEndian(data[8..]));
    }

    // BMP: "BM" and the 14-byte file header, then the DIB header, whose first field is its size (little-endian). Size 12 is
    // BITMAPCOREHEADER (OS/2 core): WORD width and height. Size 40 or more is BITMAPINFOHEADER or a later header that starts the same
    // way (V4, V5): LONG width and height; a negative height is a top-down bitmap, so the height is its absolute value. Other sizes
    // (such as the OS/2 2.x headers) are undecodable. Microsoft Learn, BITMAPCOREHEADER and BITMAPINFOHEADER (wingdi.h).
    private static PictureFacts? ReadBmp(ReadOnlySpan<byte> data)
    {
        if (data.Length < 2 || data[0] != (byte)'B' || data[1] != (byte)'M')
        {
            return null;
        }

        const int FileHeaderSize = 14;
        if (data.Length < FileHeaderSize + 4)
        {
            return PictureFacts.Undecodable;
        }

        var headerSize = BinaryPrimitives.ReadUInt32LittleEndian(data[FileHeaderSize..]);
        if (headerSize == 12)
        {
            return data.Length < FileHeaderSize + 12
                ? PictureFacts.Undecodable
                : new PictureFacts(
                    ImageFormat.Bmp,
                    BinaryPrimitives.ReadUInt16LittleEndian(data[(FileHeaderSize + 4)..]),
                    BinaryPrimitives.ReadUInt16LittleEndian(data[(FileHeaderSize + 6)..]));
        }

        if (headerSize < 40 || data.Length < FileHeaderSize + 40)
        {
            return PictureFacts.Undecodable;
        }

        var width = BinaryPrimitives.ReadInt32LittleEndian(data[(FileHeaderSize + 4)..]);
        var height = BinaryPrimitives.ReadInt32LittleEndian(data[(FileHeaderSize + 8)..]);
        return height == int.MinValue
            ? PictureFacts.Undecodable
            : new PictureFacts(ImageFormat.Bmp, width, Math.Abs(height));
    }
}
