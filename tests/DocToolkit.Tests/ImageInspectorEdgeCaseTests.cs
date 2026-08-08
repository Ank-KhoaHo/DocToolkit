using Xunit;

namespace DocToolkit.Tests;

/// <summary>
/// The malformed-input paths of <see cref="ImageInspector"/>.
///
/// Written to kill surviving mutants: the file scored <b>57.1%</b> when mutation testing was added
/// — 54 survivors and 18 mutants no test reached at all, the worst in the repository. That matters
/// more here than the number suggests, because this is the code that decides what a byte array
/// <i>is</i> from its magic bytes, and getting it wrong renders as a blank frame in Word with no
/// exception anywhere. <c>ImageInspectorTests</c> covers the happy paths; this covers the ways a
/// file can be broken.
/// </summary>
public class ImageInspectorEdgeCaseTests
{
    private const string PngHeader = "\x89PNG\r\n\x1a\n";

    /// <summary>A PNG whose IHDR declares <paramref name="width"/> x <paramref name="height"/>.</summary>
    private static byte[] PngWithSize(uint width, uint height)
    {
        var bytes = new byte[24];
        for (var i = 0; i < 8; i++) bytes[i] = (byte)PngHeader[i];
        System.Buffers.Binary.BinaryPrimitives.WriteUInt32BigEndian(bytes.AsSpan(16, 4), width);
        System.Buffers.Binary.BinaryPrimitives.WriteUInt32BigEndian(bytes.AsSpan(20, 4), height);
        return bytes;
    }

    /// <summary>SOI, then whatever segments are given.</summary>
    private static byte[] Jpeg(params byte[][] segments)
    {
        var all = new List<byte> { 0xFF, 0xD8 };
        foreach (var segment in segments) all.AddRange(segment);
        return all.ToArray();
    }

    /// <summary>A Start-Of-Frame carrying a size. <paramref name="marker"/> selects which SOF.</summary>
    private static byte[] Sof(ushort width, ushort height, byte marker = 0xC0) => new byte[]
    {
        0xFF, marker, 0x00, 0x11, 0x08,
        (byte)(height >> 8), (byte)(height & 0xFF),
        (byte)(width >> 8), (byte)(width & 0xFF),
    };

    /// <summary>A length-carrying segment of <paramref name="payload"/> zero bytes.</summary>
    private static byte[] Segment(byte marker, int payload)
    {
        var length = payload + 2;
        var bytes = new byte[2 + length];
        bytes[0] = 0xFF;
        bytes[1] = marker;
        bytes[2] = (byte)(length >> 8);
        bytes[3] = (byte)(length & 0xFF);
        return bytes;
    }

    // =====================================================================================
    // PNG
    // =====================================================================================

    // A zero dimension is schema-valid as bytes and meaningless as an image. Left unchecked it
    // reaches OOXML as a zero extent, which renders as nothing at all.
    [Theory]
    [InlineData(0u, 10u)]
    [InlineData(10u, 0u)]
    [InlineData(0u, 0u)]
    public void Png_RejectsAZeroDimension(uint width, uint height)
    {
        var ex = Assert.Throws<DocumentConversionException>(
            () => ImageInspector.Inspect(PngWithSize(width, height)));

        Assert.Contains("nonsensical size", ex.Message, StringComparison.Ordinal);
    }

    // Read as a signed int, a dimension with the top bit set comes back negative rather than huge.
    [Fact]
    public void Png_RejectsADimensionThatOverflowsIntoNegative()
    {
        var ex = Assert.Throws<DocumentConversionException>(
            () => ImageInspector.Inspect(PngWithSize(0x80000000, 10)));

        Assert.Contains("nonsensical size", ex.Message, StringComparison.Ordinal);
    }

    // 23 bytes: enough to be recognised as PNG, one short of a complete IHDR. The boundary matters -
    // an off-by-one here reads past the array.
    [Fact]
    public void Png_RejectsAnIhdrOneByteShort()
    {
        var truncated = PngWithSize(10, 10)[..23];

        var ex = Assert.Throws<DocumentConversionException>(() => ImageInspector.Inspect(truncated));

        Assert.Contains("truncated", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Png_AcceptsExactlyTwentyFourBytes()
    {
        var info = ImageInspector.Inspect(PngWithSize(7, 9));

        Assert.Equal(7, info.WidthPx);
        Assert.Equal(9, info.HeightPx);
    }

    // =====================================================================================
    // JPEG segment walking
    // =====================================================================================

    [Fact]
    public void Jpeg_RejectsAStreamWithNoStartOfFrame()
    {
        // A comment segment and nothing else: well-formed, and carries no size.
        var ex = Assert.Throws<DocumentConversionException>(
            () => ImageInspector.Inspect(Jpeg(Segment(0xFE, 4))));

        Assert.Contains("no Start-Of-Frame", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Jpeg_RejectsATruncatedFrameHeader()
    {
        // SOF marker present, then the stream ends before the 9 bytes the size needs.
        var truncated = Jpeg(new byte[] { 0xFF, 0xC0, 0x00, 0x11, 0x08, 0x00 });

        var ex = Assert.Throws<DocumentConversionException>(() => ImageInspector.Inspect(truncated));

        Assert.Contains("truncated inside its frame header", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Jpeg_RejectsATruncatedSegmentHeader()
    {
        // A non-standalone marker whose two length bytes never arrive.
        var truncated = Jpeg(new byte[] { 0xFF, 0xE0, 0x00 });

        var ex = Assert.Throws<DocumentConversionException>(() => ImageInspector.Inspect(truncated));

        Assert.Contains("truncated inside a segment header", ex.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(0, 10)]
    [InlineData(10, 0)]
    public void Jpeg_RejectsAZeroDimension(ushort width, ushort height)
    {
        var ex = Assert.Throws<DocumentConversionException>(
            () => ImageInspector.Inspect(Jpeg(Sof(width, height))));

        Assert.Contains("nonsensical size", ex.Message, StringComparison.Ordinal);
    }

    // C4, C8 and CC sit inside the C0..CF range but are NOT frame headers - Huffman table, reserved
    // and arithmetic conditioning. Reading a size from one yields whatever those bytes happen to be,
    // silently. The three are asserted individually because a mutation could drop any one of them.
    [Theory]
    [InlineData((byte)0xC4)]
    [InlineData((byte)0xC8)]
    [InlineData((byte)0xCC)]
    public void Jpeg_DoesNotReadASizeFromANonFrameMarkerInTheSofRange(byte marker)
    {
        // The impostor declares 999x999 and is skipped as a length-carrying segment; the real SOF
        // that follows declares 4x5. Reading the impostor would return 999x999.
        var impostor = new byte[] { 0xFF, marker, 0x00, 0x11, 0x08, 0x03, 0xE7, 0x03, 0xE7, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };

        var info = ImageInspector.Inspect(Jpeg(impostor, Sof(4, 5)));

        Assert.Equal(4, info.WidthPx);
        Assert.Equal(5, info.HeightPx);
    }

    // Standalone markers carry no length word. Treating one as a segment reads its next two bytes
    // as a length and skips an arbitrary distance - usually past the real SOF.
    [Theory]
    [InlineData((byte)0xD0)]
    [InlineData((byte)0xD7)]
    [InlineData((byte)0xD8)]
    public void Jpeg_WalksPastStandaloneMarkersToReachTheFrame(byte marker)
    {
        var info = ImageInspector.Inspect(Jpeg(new byte[] { 0xFF, marker }, Sof(11, 13)));

        Assert.Equal(11, info.WidthPx);
        Assert.Equal(13, info.HeightPx);
    }

    // Fill bytes and entropy-coded data are not markers; the walker advances one byte at a time
    // through them rather than mistaking them for segments.
    [Fact]
    public void Jpeg_SkipsNonMarkerBytesBeforeTheFrame()
    {
        // The filler follows a real segment rather than leading: Inspect only recognises JPEG
        // at all when the bytes begin FF D8 FF, so a stream whose third byte is not a marker
        // never reaches the walker. Found by this test failing, which is the point of it.
        var info = ImageInspector.Inspect(
            Jpeg(Segment(0xE0, 2), new byte[] { 0x00, 0x12, 0x34 }, Sof(6, 8)));

        Assert.Equal(6, info.WidthPx);
        Assert.Equal(8, info.HeightPx);
    }

    // A length-carrying segment must be skipped by its declared length, landing exactly on the next
    // marker. An arithmetic slip here lands mid-segment and reads noise as a marker.
    [Fact]
    public void Jpeg_SkipsALengthCarryingSegmentByItsDeclaredLength()
    {
        var info = ImageInspector.Inspect(Jpeg(Segment(0xE0, 16), Sof(3, 21)));

        Assert.Equal(3, info.WidthPx);
        Assert.Equal(21, info.HeightPx);
    }

    // =====================================================================================
    // Format naming - the error message is the only clue a caller gets.
    // =====================================================================================

    [Theory]
    [InlineData(new byte[] { 0x42, 0x4D, 1, 2, 3, 4, 5, 6, 7, 8 }, "BMP")]
    [InlineData(new byte[] { 0x52, 0x49, 0x46, 0x46, 1, 2, 3, 4, 5, 6 }, "RIFF")]
    [InlineData(new byte[] { 0x3C, 0x73, 0x76, 0x67, 1, 2, 3, 4, 5, 6 }, "SVG")]
    [InlineData(new byte[] { 0x3C, 0x3F, 0x78, 0x6D, 0x6C, 2, 3, 4, 5, 6 }, "SVG")]
    [InlineData(new byte[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 }, "unrecognised")]
    public void Inspect_NamesWhatTheBytesLookLike(byte[] image, string expected)
    {
        var ex = Assert.Throws<DocumentConversionException>(() => ImageInspector.Inspect(image));

        Assert.Contains(expected, ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    // Under 8 bytes there is not enough to identify anything, and saying so beats "unrecognised".
    [Fact]
    public void Inspect_SaysWhenThereAreTooFewBytesToIdentify()
    {
        var ex = Assert.Throws<DocumentConversionException>(
            () => ImageInspector.Inspect(new byte[] { 1, 2, 3 }));

        Assert.Contains("too short to identify", ex.Message, StringComparison.Ordinal);
        Assert.Contains("3 bytes", ex.Message, StringComparison.Ordinal);
    }

    // The prefix check must not read past a buffer shorter than the prefix it is comparing.
    [Fact]
    public void Inspect_HandlesABufferShorterThanTheLongestMagicPrefix()
    {
        var ex = Assert.Throws<DocumentConversionException>(
            () => ImageInspector.Inspect(new byte[] { 0x3C }));

        Assert.Contains("1 bytes", ex.Message, StringComparison.Ordinal);
    }
}
