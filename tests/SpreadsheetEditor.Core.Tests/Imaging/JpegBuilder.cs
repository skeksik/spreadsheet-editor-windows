namespace SpreadsheetEditor.Core.Tests.Imaging;

/// <summary>Builds a made-up JPEG marker stream (ITU-T T.81 Annex B) in code: SOI, then whatever the test adds.</summary>
internal sealed class JpegBuilder
{
    private readonly List<byte> _bytes = [0xFF, 0xD8];

    /// <summary>The length of the stream up to the end of the last frame header added; the reader needs exactly this much.</summary>
    public int FrameHeaderEnd { get; private set; }

    /// <summary>A marker segment: X'FF', the marker, a big-endian length that counts itself, then the payload.</summary>
    public JpegBuilder Segment(byte marker, params ReadOnlySpan<byte> payload)
    {
        _bytes.Add(0xFF);
        _bytes.Add(marker);
        PictureBytes.AddUInt16BigEndian(_bytes, (ushort)(payload.Length + 2));
        _bytes.AddRange(payload);
        return this;
    }

    /// <summary>A marker that stands alone (SOI, EOI, RSTm, TEM): no length.</summary>
    public JpegBuilder Standalone(byte marker)
    {
        _bytes.Add(0xFF);
        _bytes.Add(marker);
        return this;
    }

    /// <summary>X'FF' fill bytes, which may precede any marker.</summary>
    public JpegBuilder Fill(int count)
    {
        for (var i = 0; i < count; i++)
        {
            _bytes.Add(0xFF);
        }

        return this;
    }

    /// <summary>A JFIF APP0 segment, as most cameras and editors write first.</summary>
    public JpegBuilder App0() =>
        Segment(0xE0, (byte)'J', (byte)'F', (byte)'I', (byte)'F', 0, 1, 1, 0, 0, 1, 0, 1, 0, 0);

    /// <summary>A frame header (B.2.2): precision 8, lines, samples per line, then three components.</summary>
    public JpegBuilder Frame(byte marker, ushort width, ushort height)
    {
        Segment(
            marker,
            8,
            (byte)(height >> 8),
            (byte)height,
            (byte)(width >> 8),
            (byte)width,
            3,
            1,
            0x22,
            0,
            2,
            0x11,
            1,
            3,
            0x11,
            1);
        FrameHeaderEnd = _bytes.Count;
        return this;
    }

    /// <summary>A scan header and a few bytes of entropy-coded data.</summary>
    public JpegBuilder Scan()
    {
        Segment(0xDA, 1, 1, 0, 0, 63, 0);
        PictureBytes.Append(_bytes, 0x12, 0x34, 0xFF, 0x00, 0x56);
        return this;
    }

    public byte[] Build() => [.. _bytes];

    /// <summary>A typical baseline file: SOI, APP0, DQT, SOF0, DHT, SOS, data, EOI.</summary>
    public static JpegBuilder Typical(ushort width, ushort height, byte frameMarker = 0xC0) =>
        new JpegBuilder()
            .App0()
            .Segment(0xDB, 0, 1, 2, 3, 4, 5, 6, 7)
            .Frame(frameMarker, width, height)
            .Segment(0xC4, 0, 0, 1, 5, 1, 1, 1)
            .Scan()
            .Standalone(0xD9);
}
