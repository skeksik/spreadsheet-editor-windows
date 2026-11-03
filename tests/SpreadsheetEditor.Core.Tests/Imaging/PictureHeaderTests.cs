using System.Buffers.Binary;
using SpreadsheetEditor.Core.Imaging;
using SpreadsheetEditor.Core.Model;

namespace SpreadsheetEditor.Core.Tests.Imaging;

[TestClass]
public sealed class PictureHeaderTests
{
    // PNG

    [TestMethod]
    public void ReadsPngSize() =>
        Assert.AreEqual(new PictureFacts(ImageFormat.Png, 640, 480), PictureHeader.Read(PictureBytes.Png(640, 480)));

    [TestMethod]
    public void ReadsLargestPngSize() =>
        Assert.AreEqual(
            new PictureFacts(ImageFormat.Png, int.MaxValue, 1),
            PictureHeader.Read(PictureBytes.Png(int.MaxValue, 1)));

    [TestMethod]
    [DataRow(0u, 10u)]
    [DataRow(10u, 0u)]
    [DataRow(0x8000_0000u, 10u)]
    [DataRow(10u, 0xFFFF_FFFFu)]
    public void RejectsPngSizeOutOfRange(uint width, uint height) =>
        Assert.AreEqual(PictureFacts.Undecodable, PictureHeader.Read(PictureBytes.Png(width, height)));

    [TestMethod]
    public void RejectsPngWhoseFirstChunkIsNotIhdr() =>
        Assert.AreEqual(PictureFacts.Undecodable, PictureHeader.Read(PictureBytes.Png(640, 480, firstChunk: "IDAT")));

    [TestMethod]
    public void RejectsPngWithWrongIhdrLength() =>
        Assert.AreEqual(PictureFacts.Undecodable, PictureHeader.Read(PictureBytes.Png(640, 480, chunkLength: 12)));

    // JPEG

    [TestMethod]
    public void ReadsTypicalJpegSize() =>
        Assert.AreEqual(new PictureFacts(ImageFormat.Jpeg, 1920, 1080), PictureHeader.Read(JpegBuilder.Typical(1920, 1080).Build()));

    [TestMethod]
    [DataRow(0xC0)]
    [DataRow(0xC1)]
    [DataRow(0xC2)]
    [DataRow(0xC3)]
    [DataRow(0xC5)]
    [DataRow(0xC6)]
    [DataRow(0xC7)]
    [DataRow(0xC9)]
    [DataRow(0xCA)]
    [DataRow(0xCB)]
    [DataRow(0xCD)]
    [DataRow(0xCE)]
    [DataRow(0xCF)]
    public void ReadsEveryStartOfFrameMarker(int marker) =>
        Assert.AreEqual(
            new PictureFacts(ImageFormat.Jpeg, 300, 200),
            PictureHeader.Read(JpegBuilder.Typical(300, 200, (byte)marker).Build()));

    [TestMethod]
    [DataRow(0xC4)] // DHT
    [DataRow(0xC8)] // JPG
    [DataRow(0xCC)] // DAC
    public void DoesNotTakeDhtJpgOrDacForAFrameHeader(int marker)
    {
        // A segment with this marker whose payload looks like a frame header of 999 x 999, then the real frame.
        var bytes = new JpegBuilder()
            .Segment((byte)marker, 8, 0x03, 0xE7, 0x03, 0xE7, 1, 1, 0x11, 0)
            .Frame(0xC0, 40, 30)
            .Standalone(0xD9)
            .Build();

        Assert.AreEqual(new PictureFacts(ImageFormat.Jpeg, 40, 30), PictureHeader.Read(bytes));
    }

    [TestMethod]
    [DataRow(0xC4)]
    [DataRow(0xC8)]
    [DataRow(0xCC)]
    public void RejectsJpegWhoseOnlyFrameLikeSegmentIsDhtJpgOrDac(int marker)
    {
        var bytes = new JpegBuilder().Segment((byte)marker, 8, 0, 10, 0, 10, 1, 1, 0x11, 0).Standalone(0xD9).Build();
        Assert.AreEqual(PictureFacts.Undecodable, PictureHeader.Read(bytes));
    }

    [TestMethod]
    public void SkipsFillBytesAndStandaloneMarkers()
    {
        var bytes = new JpegBuilder()
            .Fill(3)
            .App0()
            .Standalone(0x01) // TEM
            .Standalone(0xD3) // RST3
            .Segment(0xFE, (byte)'h', (byte)'i') // COM
            .Fill(5)
            .Frame(0xC2, 800, 600)
            .Build();

        Assert.AreEqual(new PictureFacts(ImageFormat.Jpeg, 800, 600), PictureHeader.Read(bytes));
    }

    [TestMethod]
    public void RejectsJpegWithScanBeforeFrame() =>
        Assert.AreEqual(
            PictureFacts.Undecodable,
            PictureHeader.Read(new JpegBuilder().App0().Scan().Frame(0xC0, 10, 10).Build()));

    [TestMethod]
    public void RejectsJpegEndingBeforeFrame() =>
        Assert.AreEqual(
            PictureFacts.Undecodable,
            PictureHeader.Read(new JpegBuilder().App0().Standalone(0xD9).Frame(0xC0, 10, 10).Build()));

    [TestMethod]
    public void RejectsJpegWithLinesDefinedByDnl() =>
        Assert.AreEqual(PictureFacts.Undecodable, PictureHeader.Read(JpegBuilder.Typical(100, 0).Build()));

    [TestMethod]
    public void RejectsJpegWithZeroWidth() =>
        Assert.AreEqual(PictureFacts.Undecodable, PictureHeader.Read(JpegBuilder.Typical(0, 100).Build()));

    [TestMethod]
    public void RejectsJpegSegmentLengthBelowTwo()
    {
        byte[] bytes = [0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x01, 0xFF, 0xC0, 0x00, 0x0B, 8, 0, 10, 0, 10, 1, 1, 0x11, 0];
        Assert.AreEqual(PictureFacts.Undecodable, PictureHeader.Read(bytes));
    }

    [TestMethod]
    public void RejectsShortFrameHeader()
    {
        // SOF0 with length 7: one byte short of Lf, P, Y, X and Nf.
        var bytes = new JpegBuilder().Segment(0xC0, 8, 0, 10, 0, 10).Standalone(0xD9).Build();
        Assert.AreEqual(PictureFacts.Undecodable, PictureHeader.Read(bytes));
    }

    [TestMethod]
    public void RejectsJpegWithDataWhereAMarkerShouldBe()
    {
        byte[] bytes = [0xFF, 0xD8, 0x12, 0xFF, 0xC0, 0x00, 0x0B, 8, 0, 10, 0, 10, 1, 1, 0x11, 0];
        Assert.AreEqual(PictureFacts.Undecodable, PictureHeader.Read(bytes));
    }

    [TestMethod]
    public void RejectsJpegOfFillBytesOnly() =>
        Assert.AreEqual(PictureFacts.Undecodable, PictureHeader.Read(new JpegBuilder().Fill(50).Build()));

    // GIF

    [TestMethod]
    [DataRow("87a")]
    [DataRow("89a")]
    public void ReadsGifLogicalScreenSize(string version) =>
        Assert.AreEqual(new PictureFacts(ImageFormat.Gif, 513, 65535), PictureHeader.Read(PictureBytes.Gif(version, 513, 65535)));

    [TestMethod]
    public void RejectsUnknownGifVersion() =>
        Assert.AreEqual(PictureFacts.Undecodable, PictureHeader.Read(PictureBytes.Gif("88a", 10, 10)));

    [TestMethod]
    [DataRow(0, 10)]
    [DataRow(10, 0)]
    public void RejectsGifWithZeroSize(int width, int height) =>
        Assert.AreEqual(PictureFacts.Undecodable, PictureHeader.Read(PictureBytes.Gif("89a", (ushort)width, (ushort)height)));

    // BMP

    [TestMethod]
    public void ReadsBmpCoreHeader() =>
        Assert.AreEqual(new PictureFacts(ImageFormat.Bmp, 65535, 3), PictureHeader.Read(PictureBytes.BmpCore(65535, 3)));

    [TestMethod]
    [DataRow(40u)]
    [DataRow(108u)]
    [DataRow(124u)]
    public void ReadsBmpInfoHeaderBottomUp(uint headerSize) =>
        Assert.AreEqual(new PictureFacts(ImageFormat.Bmp, 320, 240), PictureHeader.Read(PictureBytes.BmpInfo(320, 240, headerSize)));

    [TestMethod]
    public void ReadsBmpTopDownHeightAsPositive() =>
        Assert.AreEqual(new PictureFacts(ImageFormat.Bmp, 320, 240), PictureHeader.Read(PictureBytes.BmpInfo(320, -240)));

    [TestMethod]
    [DataRow(-320, 240)]
    [DataRow(0, 240)]
    [DataRow(320, 0)]
    [DataRow(320, int.MinValue)]
    public void RejectsBmpInfoSizeOutOfRange(int width, int height) =>
        Assert.AreEqual(PictureFacts.Undecodable, PictureHeader.Read(PictureBytes.BmpInfo(width, height)));

    [TestMethod]
    [DataRow(0, 1)]
    [DataRow(1, 0)]
    public void RejectsBmpCoreWithZeroSize(int width, int height) =>
        Assert.AreEqual(PictureFacts.Undecodable, PictureHeader.Read(PictureBytes.BmpCore((ushort)width, (ushort)height)));

    [TestMethod]
    [DataRow(0u)]
    [DataRow(16u)] // OS/2 2.x short header: no source for its layout was checked, so it is not read
    [DataRow(39u)]
    public void RejectsUnknownBmpHeaderSize(uint headerSize)
    {
        var bytes = PictureBytes.BmpInfo(320, 240);
        BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(14), headerSize);

        Assert.AreEqual(PictureFacts.Undecodable, PictureHeader.Read(bytes));
    }

    // Truncated and garbage input

    public static IEnumerable<object[]> ValidPictures =>
    [
        ["PNG", PictureBytes.Png(640, 480), PictureBytes.PngHeaderLength],
        ["JPEG", JpegBuilder.Typical(1920, 1080).Build(), JpegBuilder.Typical(1920, 1080).FrameHeaderEnd],
        ["GIF", PictureBytes.Gif("89a", 100, 50), PictureBytes.GifHeaderLength],
        ["BMP core", PictureBytes.BmpCore(100, 50), PictureBytes.BmpCoreHeaderLength],
        ["BMP info", PictureBytes.BmpInfo(100, -50, 124), PictureBytes.BmpInfoHeaderLength],
    ];

    [TestMethod]
    [DynamicData(nameof(ValidPictures))]
    public void EveryTruncatedHeaderIsUndecodable(string name, byte[] bytes, int headerLength)
    {
        var full = PictureHeader.Read(bytes);
        Assert.IsTrue(full.IsDecodable, name);

        for (var length = 0; length < bytes.Length; length++)
        {
            var expected = length < headerLength ? PictureFacts.Undecodable : full;
            Assert.AreEqual(expected, PictureHeader.Read(bytes.AsSpan(0, length)), $"{name}, first {length} bytes");
        }
    }

    [TestMethod]
    public void EmptyInputIsUndecodable() => Assert.AreEqual(PictureFacts.Undecodable, PictureHeader.Read([]));

    [TestMethod]
    public void TextIsUndecodable() =>
        Assert.AreEqual(PictureFacts.Undecodable, PictureHeader.Read("not a picture at all"u8));

    [TestMethod]
    public void GarbageNeverThrowsAndIsUndecodableOrPositive()
    {
        // Deterministic pseudo-random bytes, alone and after each format's signature, so the parsers walk into garbage.
        byte[][] prefixes =
        [
            [],
            [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A],
            [0xFF, 0xD8],
            [(byte)'G', (byte)'I', (byte)'F', (byte)'8', (byte)'9', (byte)'a'],
            [(byte)'B', (byte)'M'],
        ];
        var state = 0x2545F491u;
        foreach (var prefix in prefixes)
        {
            for (var length = 0; length < 300; length++)
            {
                var bytes = new byte[prefix.Length + length];
                prefix.CopyTo(bytes, 0);
                for (var i = prefix.Length; i < bytes.Length; i++)
                {
                    state ^= state << 13;
                    state ^= state >> 17;
                    state ^= state << 5;
                    bytes[i] = (byte)state;
                }

                var facts = PictureHeader.Read(bytes);
                Assert.IsTrue(
                    facts == PictureFacts.Undecodable
                        || (facts.Format != ImageFormat.Unknown && facts.PixelWidth > 0 && facts.PixelHeight > 0),
                    $"prefix {prefix.Length} bytes, {length} random bytes: {facts}");
            }
        }
    }
}
