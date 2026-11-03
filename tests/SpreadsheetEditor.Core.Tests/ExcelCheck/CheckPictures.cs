using System.IO.Compression;

namespace SpreadsheetEditor.Core.Tests.ExcelCheck;

/// <summary>
/// Small made-up pictures for the Excel check, drawn in code and encoded as real, decodable PNG and baseline JPEG files, so that Excel
/// shows them. A sample picture has a white region of the aspect it is shown at, holding a black ring that touches the region's longer
/// sides and a number; everything outside the region is red. In the saved file the region alone must show: a red edge means a wrong
/// crop, an oval ring means a stretched picture.
/// </summary>
internal static class CheckPictures
{
    public static readonly Rgb White = new(255, 255, 255);
    public static readonly Rgb Black = new(0, 0, 0);
    public static readonly Rgb Red = new(220, 0, 0);
    public static readonly Rgb Blue = new(0, 70, 200);

    /// <summary>
    /// A PNG of <paramref name="width"/> × <paramref name="height"/> pixels whose centre region has the aspect
    /// <paramref name="targetWidth"/> : <paramref name="targetHeight"/> (the box the writer shows it in, docs/spec.md §5.4 step 3).
    /// </summary>
    public static byte[] Png(int width, int height, int targetWidth, int targetHeight, int number) =>
        EncodePng(Draw(width, height, targetWidth, targetHeight, number));

    /// <summary>The pixels of a sample picture; see the class summary.</summary>
    public static Rgb[,] Draw(int width, int height, int targetWidth, int targetHeight, int number)
    {
        var pixels = new Rgb[height, width];
        var (left, top, regionWidth, regionHeight) = Region(width, height, (double)targetWidth / targetHeight);
        var radius = Math.Min(regionWidth, regionHeight) / 2.0;
        var thickness = Math.Max(2.0, radius / 8);
        var centreX = left + (regionWidth / 2.0);
        var centreY = top + (regionHeight / 2.0);
        for (var y = 0; y < height; y++)
        {
            for (var x = 0; x < width; x++)
            {
                var inside = x >= left && x < left + regionWidth && y >= top && y < top + regionHeight;
                var distance = Math.Sqrt(Math.Pow(x + 0.5 - centreX, 2) + Math.Pow(y + 0.5 - centreY, 2));
                pixels[y, x] = !inside ? Red : distance <= radius && distance >= radius - thickness ? Black : White;
            }
        }

        var scale = Math.Max(1, (int)(radius / 6));
        DrawNumber(pixels, number, (int)centreX, (int)centreY, scale);
        return pixels;
    }

    /// <summary>
    /// The centre region of the aspect <paramref name="target"/> (width / height), as the writer's centre crop keeps it (docs/spec.md
    /// §5.4 step 3); whole pixels, rounded.
    /// </summary>
    public static (int Left, int Top, int Width, int Height) Region(int width, int height, double target)
    {
        var aspect = (double)width / height;
        if (aspect > target)
        {
            var kept = (int)Math.Round(height * target);
            return ((width - kept) / 2, 0, kept, height);
        }

        if (aspect < target)
        {
            var kept = (int)Math.Round(width / target);
            return (0, (height - kept) / 2, width, kept);
        }

        return (0, 0, width, height);
    }

    /// <summary>
    /// A baseline JPEG of <paramref name="width"/> × <paramref name="height"/> pixels (multiples of 8): the top half red, the bottom half
    /// blue, and a black square at the top-left corner. With <paramref name="orientation"/> it carries an EXIF APP1 segment whose
    /// Orientation tag (0x0112) has that value; 6 means "the stored top row is the right side of the picture as shown".
    /// </summary>
    public static byte[] OrientationJpeg(int width, int height, ushort? orientation)
    {
        var pixels = new Rgb[height, width];
        var square = Math.Min(width, height) / 2 / 8 * 8;
        for (var y = 0; y < height; y++)
        {
            for (var x = 0; x < width; x++)
            {
                pixels[y, x] = x < square && y < square ? Black : y < height / 2 ? Red : Blue;
            }
        }

        return EncodeJpeg(pixels, orientation);
    }

    public readonly record struct Rgb(byte R, byte G, byte B);

    // A 3 × 5 bitmap font for the digits, one string of 15 cells per digit, row by row.
    private static readonly string[] Digits =
    [
        "111101101101111", "010110010010111", "111001111100111", "111001111001111", "101101111001001",
        "111100111001111", "111100111101111", "111001010010010", "111101111101111", "111101111001111",
    ];

    private static void DrawNumber(Rgb[,] pixels, int number, int centreX, int centreY, int scale)
    {
        var text = number.ToString(System.Globalization.CultureInfo.InvariantCulture);
        var totalWidth = ((text.Length * 4) - 1) * scale;
        var x0 = centreX - (totalWidth / 2);
        var y0 = centreY - (5 * scale / 2);
        for (var i = 0; i < text.Length; i++)
        {
            var glyph = Digits[text[i] - '0'];
            for (var cell = 0; cell < 15; cell++)
            {
                if (glyph[cell] != '1')
                {
                    continue;
                }

                for (var dy = 0; dy < scale; dy++)
                {
                    for (var dx = 0; dx < scale; dx++)
                    {
                        var x = x0 + (i * 4 * scale) + (cell % 3 * scale) + dx;
                        var y = y0 + (cell / 3 * scale) + dy;
                        if (y >= 0 && y < pixels.GetLength(0) && x >= 0 && x < pixels.GetLength(1))
                        {
                            pixels[y, x] = Black;
                        }
                    }
                }
            }
        }
    }

    /// <summary>
    /// A PNG (ISO/IEC 15948, W3C PNG second edition): IHDR with 8-bit RGB, one IDAT of zlib-compressed scanlines each with filter type 0,
    /// and IEND; every chunk with its CRC-32.
    /// </summary>
    public static byte[] EncodePng(Rgb[,] pixels)
    {
        int height = pixels.GetLength(0), width = pixels.GetLength(1);
        var raw = new byte[height * ((width * 3) + 1)];
        var i = 0;
        for (var y = 0; y < height; y++)
        {
            raw[i++] = 0;
            for (var x = 0; x < width; x++)
            {
                raw[i++] = pixels[y, x].R;
                raw[i++] = pixels[y, x].G;
                raw[i++] = pixels[y, x].B;
            }
        }

        using var compressed = new MemoryStream();
        using (var zlib = new ZLibStream(compressed, CompressionLevel.SmallestSize, leaveOpen: true))
        {
            zlib.Write(raw);
        }

        var output = new List<byte> { 0x89, (byte)'P', (byte)'N', (byte)'G', 0x0D, 0x0A, 0x1A, 0x0A };
        var header = new List<byte>();
        AddUInt32(header, (uint)width);
        AddUInt32(header, (uint)height);
        header.AddRange([8, 2, 0, 0, 0]);
        Chunk(output, "IHDR", [.. header]);
        Chunk(output, "IDAT", compressed.ToArray());
        Chunk(output, "IEND", []);
        return [.. output];
    }

    private static void Chunk(List<byte> output, string type, byte[] data)
    {
        AddUInt32(output, (uint)data.Length);
        var typed = new byte[4 + data.Length];
        for (var k = 0; k < 4; k++)
        {
            typed[k] = (byte)type[k];
        }

        data.CopyTo(typed, 4);
        output.AddRange(typed);
        AddUInt32(output, Crc32(typed));
    }

    // CRC-32 as the PNG specification's annex D gives it: polynomial 0xEDB88320 (reflected), initial and final XOR 0xFFFFFFFF.
    internal static uint Crc32(byte[] data)
    {
        var crc = 0xFFFFFFFFu;
        foreach (var b in data)
        {
            crc ^= b;
            for (var k = 0; k < 8; k++)
            {
                crc = (crc & 1) != 0 ? 0xEDB88320u ^ (crc >> 1) : crc >> 1;
            }
        }

        return crc ^ 0xFFFFFFFFu;
    }

    private static void AddUInt32(List<byte> output, uint value)
    {
        output.Add((byte)(value >> 24));
        output.Add((byte)(value >> 16));
        output.Add((byte)(value >> 8));
        output.Add((byte)value);
    }

    /// <summary>
    /// A baseline sequential JPEG (ITU-T T.81): three components (JFIF YCbCr, no subsampling), every 8 × 8 block a flat colour, so each
    /// block is its DC coefficient and an end-of-block. One quantisation table of 8s; the DC table is the standard luminance DC table
    /// (T.81 Table K.3); the AC table holds the one symbol used, end-of-block (0x00), as a 1-bit code. Width and height are multiples of 8.
    /// </summary>
    public static byte[] EncodeJpeg(Rgb[,] pixels, ushort? orientation)
    {
        int height = pixels.GetLength(0), width = pixels.GetLength(1);
        if (width % 8 != 0 || height % 8 != 0)
        {
            throw new ArgumentException("The size must be a multiple of 8.", nameof(pixels));
        }

        var output = new List<byte> { 0xFF, 0xD8 };
        Segment(output, 0xE0, [(byte)'J', (byte)'F', (byte)'I', (byte)'F', 0, 1, 1, 0, 0, 1, 0, 1, 0, 0]);
        if (orientation is { } value)
        {
            // APP1 "Exif\0\0", then a big-endian TIFF header and IFD0 with one entry: Orientation, SHORT, count 1 (Exif 2.3, 4.6.4).
            Segment(output, 0xE1,
            [
                (byte)'E', (byte)'x', (byte)'i', (byte)'f', 0, 0,
                (byte)'M', (byte)'M', 0, 42, 0, 0, 0, 8,
                0, 1,
                0x01, 0x12, 0, 3, 0, 0, 0, 1, (byte)(value >> 8), (byte)value, 0, 0,
                0, 0, 0, 0,
            ]);
        }

        // DQT: table 0, 8-bit precision, 64 values of 8.
        Segment(output, 0xDB, [0, .. Enumerable.Repeat((byte)8, 64)]);

        // SOF0: precision 8, height, width, three components with sampling 1x1, all using quantisation table 0.
        Segment(output, 0xC0,
        [
            8, (byte)(height >> 8), (byte)height, (byte)(width >> 8), (byte)width, 3,
            1, 0x11, 0, 2, 0x11, 0, 3, 0x11, 0,
        ]);

        // DHT: DC table 0 (Table K.3), then AC table 0 with the one code "0" for end-of-block.
        Segment(output, 0xC4, [0x00, .. DcCounts, .. Enumerable.Range(0, 12).Select(v => (byte)v)]);
        Segment(output, 0xC4, [0x10, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0x00]);

        // SOS: three components, each with DC and AC table 0; spectral selection 0..63, no approximation.
        Segment(output, 0xDA, [3, 1, 0x00, 2, 0x00, 3, 0x00, 0, 63, 0]);

        var (dcCodes, dcLengths) = CanonicalCodes(DcCounts);
        var writer = new BitWriter(output);
        var predictors = new int[3];
        for (var by = 0; by < height / 8; by++)
        {
            for (var bx = 0; bx < width / 8; bx++)
            {
                var (y, cb, cr) = BlockMean(pixels, bx * 8, by * 8);
                double[] samples = [y, cb, cr];
                for (var c = 0; c < 3; c++)
                {
                    // The DC coefficient of a flat block is 8 × (sample − 128) (T.81 A.3.3); quantised by 8 it is sample − 128.
                    var dc = (int)Math.Round(samples[c] - 128);
                    var diff = dc - predictors[c];
                    predictors[c] = dc;
                    var magnitude = Math.Abs(diff);
                    var size = 0;
                    while (magnitude >> size != 0)
                    {
                        size++;
                    }

                    writer.Write(dcCodes[size], dcLengths[size]);
                    if (size > 0)
                    {
                        // Negative differences are written as diff − 1 in size bits (T.81 F.1.2.1.1).
                        writer.Write(diff >= 0 ? diff : diff + (1 << size) - 1, size);
                    }

                    writer.Write(0, 1);
                }
            }
        }

        writer.Flush();
        output.AddRange([0xFF, 0xD9]);
        return [.. output];
    }

    // Table K.3: the number of codes of each length 1..16 for the luminance DC differences, whose symbols are the sizes 0..11.
    private static readonly byte[] DcCounts = [0, 1, 5, 1, 1, 1, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0];

    private static (int[] Codes, int[] Lengths) CanonicalCodes(byte[] counts)
    {
        var codes = new List<int>();
        var lengths = new List<int>();
        var code = 0;
        for (var length = 1; length <= 16; length++)
        {
            for (var k = 0; k < counts[length - 1]; k++)
            {
                codes.Add(code++);
                lengths.Add(length);
            }

            code <<= 1;
        }

        return ([.. codes], [.. lengths]);
    }

    // JFIF's RGB to YCbCr conversion (JFIF 1.02, "Conversion to and from RGB"), averaged over one 8 × 8 block.
    private static (double Y, double Cb, double Cr) BlockMean(Rgb[,] pixels, int x0, int y0)
    {
        double r = 0, g = 0, b = 0;
        for (var y = y0; y < y0 + 8; y++)
        {
            for (var x = x0; x < x0 + 8; x++)
            {
                r += pixels[y, x].R;
                g += pixels[y, x].G;
                b += pixels[y, x].B;
            }
        }

        r /= 64;
        g /= 64;
        b /= 64;
        return (
            Math.Clamp((0.299 * r) + (0.587 * g) + (0.114 * b), 0, 255),
            Math.Clamp((-0.1687 * r) - (0.3313 * g) + (0.5 * b) + 128, 0, 255),
            Math.Clamp((0.5 * r) - (0.4187 * g) - (0.0813 * b) + 128, 0, 255));
    }

    private static void Segment(List<byte> output, byte marker, byte[] payload)
    {
        output.Add(0xFF);
        output.Add(marker);
        output.Add((byte)((payload.Length + 2) >> 8));
        output.Add((byte)(payload.Length + 2));
        output.AddRange(payload);
    }

    /// <summary>Writes entropy-coded bits most significant first, stuffing a zero byte after every X'FF' (T.81 F.1.2.3).</summary>
    private sealed class BitWriter(List<byte> output)
    {
        private int buffer;
        private int count;

        public void Write(int bits, int length)
        {
            for (var k = length - 1; k >= 0; k--)
            {
                buffer = (buffer << 1) | ((bits >> k) & 1);
                count++;
                if (count == 8)
                {
                    Emit();
                }
            }
        }

        // The last byte is padded with 1-bits (T.81 F.1.2.3).
        public void Flush()
        {
            while (count != 0)
            {
                Write(1, 1);
            }
        }

        private void Emit()
        {
            output.Add((byte)buffer);
            if (buffer == 0xFF)
            {
                output.Add(0);
            }

            buffer = 0;
            count = 0;
        }
    }
}
