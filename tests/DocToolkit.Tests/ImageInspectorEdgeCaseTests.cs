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

    // =====================================================================================
    // Exact boundaries.
    //
    // Every test below exists to kill a specific surviving mutant: a comparison whose `<=` can be
    // swapped for `<`, or `>` for `>=`, without any existing test noticing. A boundary is only
    // tested by something sitting exactly ON it, and these are the values that do.
    // =====================================================================================

    [Fact]
    public void Inspect_RejectsNull()
    {
        Assert.Throws<ArgumentNullException>(() => ImageInspector.Inspect(null!));
    }

    // The JPEG magic test needs three bytes. At exactly three it must still fire - `>= 3` mutated
    // to `> 3` makes the shortest possible JPEG header unrecognisable.
    [Fact]
    public void Inspect_RecognisesJpegFromExactlyThreeBytes()
    {
        var ex = Assert.Throws<DocumentConversionException>(
            () => ImageInspector.Inspect(new byte[] { 0xFF, 0xD8, 0xFF }));

        // Detected as JPEG, then rejected for having no frame - NOT reported as an unknown format,
        // which is what a narrowed length check would produce.
        Assert.Contains("JPEG", ex.Message, StringComparison.Ordinal);
    }

    // Zero is the boundary `is <= 0` guards. `< 0` lets it through, and a zero-size image reaches
    // OOXML as a zero extent that renders as nothing.
    [Theory]
    [InlineData(0.0, null)]
    [InlineData(null, 0.0)]
    public void Resolve_RejectsExactlyZero(double? width, double? height)
    {
        var info = ImageInspector.Inspect(ImageFixtures.Png(10, 10));

        Assert.Throws<ArgumentOutOfRangeException>(() => ImageInspector.Resolve(info, width, height));
    }

    // Exactly the ceiling must be ACCEPTED - `> MaxCoordinateEmu` mutated to `>=` rejects the
    // largest legal size, which OOXML can represent perfectly well.
    [Fact]
    public void Resolve_AcceptsASizeLandingExactlyOnTheCeiling()
    {
        const long MaxCoordinateEmu = int.MaxValue;
        const long EmuPerPoint = 12700;

        var info = ImageInspector.Inspect(ImageFixtures.Png(10, 10));
        double points = (double)MaxCoordinateEmu / EmuPerPoint;

        var (widthEmu, heightEmu) = ImageInspector.Resolve(info, points, points);

        Assert.True(widthEmu <= MaxCoordinateEmu && widthEmu > 0, $"width {widthEmu} left the legal range");
        Assert.True(heightEmu <= MaxCoordinateEmu && heightEmu > 0, $"height {heightEmu} left the legal range");
    }

    // 0xCF is the last marker in the SOF range. `<= 0xCF` mutated to `< 0xCF` stops SOF15 being
    // read as a frame, and the size is then silently taken from wherever the walk lands next.
    [Fact]
    public void Jpeg_ReadsASizeFromMarkerCf_TheLastInTheSofRange()
    {
        var info = ImageInspector.Inspect(Jpeg(Sof(17, 19, marker: 0xCF)));

        Assert.Equal(17, info.WidthPx);
        Assert.Equal(19, info.HeightPx);
    }

    // Exactly 8 bytes is the boundary between "too short to identify" and "unrecognised". `< 8`
    // mutated to `<= 8` changes which message an 8-byte buffer gets.
    [Fact]
    public void Inspect_AtExactlyEightBytesSaysUnrecognisedRatherThanTooShort()
    {
        var ex = Assert.Throws<DocumentConversionException>(
            () => ImageInspector.Inspect(new byte[] { 1, 2, 3, 4, 5, 6, 7, 8 }));

        Assert.Contains("unrecognised", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("too short", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    // A buffer exactly as long as the prefix it is compared against. `>= prefix.Length` mutated to
    // `>` makes a two-byte "BM" fail to match the two-byte BMP prefix.
    [Fact]
    public void Inspect_MatchesAMagicPrefixWhenTheBufferIsExactlyThatLong()
    {
        var ex = Assert.Throws<DocumentConversionException>(
            () => ImageInspector.Inspect(new byte[] { 0x42, 0x4D }));   // "BM", and nothing else

        Assert.Contains("BMP", ex.Message, StringComparison.Ordinal);
    }


    // =============================================================================================
    // A positive size that resolves to zero EMU.
    // =============================================================================================
    //
    // Resolve checks the RESOLVED extents, not the arguments, and the check is `is <= 0 or > Max`.
    // Mutation turned `<= 0` into `< 0` in three places and nothing failed: every existing test
    // that reaches the guard does so with a value comfortably below zero or far above the ceiling,
    // and none sits on zero itself.
    //
    // Landing exactly on zero is not hypothetical. The conversion truncates - 0.00001 pt is
    // 0.127 EMU, which is 0 - so a caller passing a legitimately positive but tiny size produces a
    // zero-extent image. Without the guard that is written into the document as a valid but
    // invisible picture, and nothing downstream validates these numbers.

    [Fact]
    public void AWidthThatTruncatesToZeroEmuIsRejectedAndBlamedOnTheWidth()
    {
        var info = new ImageInfo(ImageFormat.Png, 100, 50);

        var ex = Assert.Throws<ArgumentOutOfRangeException>(
            () => ImageInspector.Resolve(info, widthPoints: 0.00001, heightPoints: 10));

        // Not merely "it threw": with `<= 0` weakened to `< 0` on the width test alone, the height
        // is in range and Resolve returns a zero-width image instead of throwing. And the guard
        // that decides WHICH argument to name reads the same comparison, so a weakened one blames
        // heightPoints - a value the caller can see is fine.
        Assert.Equal("widthPoints", ex.ParamName);
    }

    [Fact]
    public void AHeightThatTruncatesToZeroEmuIsRejectedAndBlamedOnTheHeight()
    {
        var info = new ImageInfo(ImageFormat.Png, 100, 50);

        var ex = Assert.Throws<ArgumentOutOfRangeException>(
            () => ImageInspector.Resolve(info, widthPoints: 10, heightPoints: 0.00001));

        Assert.Equal("heightPoints", ex.ParamName);
    }


    // =============================================================================================
    // A dimension of exactly zero, rejected as an ARGUMENT rather than as an overflow.
    // =============================================================================================
    //
    // Resolve guards its arguments up front and the resolved extents afterwards, and both raise
    // ArgumentOutOfRangeException naming the same parameter. So the type and the parameter name
    // cannot tell the two apart - only the message can, which is why these assert on it.
    //
    // That distinction is worth holding. The argument guard says "Width must be positive", which
    // names what the caller did; the overflow guard talks about what OOXML can represent, which
    // for an input of 0 is a confusing thing to be told. Mutation deleted the first guard outright
    // and nothing failed, because the second one caught the fallout and threw something close
    // enough to pass every existing assertion.

    [Theory]
    [InlineData(0.0, null, "widthPoints", "Width")]
    [InlineData(null, 0.0, "heightPoints", "Height")]
    public void ADimensionOfExactlyZeroIsRejectedAsAnArgument(
        double? widthPoints, double? heightPoints, string expectedParam, string expectedSubject)
    {
        var info = new ImageInfo(ImageFormat.Png, 100, 50);

        var ex = Assert.Throws<ArgumentOutOfRangeException>(
            () => ImageInspector.Resolve(info, widthPoints, heightPoints));

        Assert.Equal(expectedParam, ex.ParamName);
        Assert.Contains($"{expectedSubject} must be positive", ex.Message, StringComparison.Ordinal);
    }


    // =============================================================================================
    // The JPEG recognition prefix, and the segment-header boundary.
    // =============================================================================================
    //
    // Three mutants lived here through five passes. All three are boundaries, and a boundary is
    // only tested by an input sitting exactly ON it - every existing JPEG case is comfortably
    // longer than the prefix and ends comfortably inside a segment.

    /// <summary>
    /// Two bytes that look like the start of a JPEG. The length test guards the two index reads
    /// that follow it, and it is an AND for that reason: loosened to OR, <c>image[2]</c> is read
    /// off the end of a two-byte array and the caller gets an IndexOutOfRangeException from inside
    /// a library that promises one exception type.
    /// </summary>
    [Fact]
    public void TwoBytesThatBeginLikeAJpegAreRejected_NotIndexedPastTheEnd()
    {
        var ex = Assert.Throws<DocumentConversionException>(() => ImageInspector.Inspect(Jpeg()));

        Assert.Contains("too short to identify", ex.Message, StringComparison.Ordinal);
    }

    /// <summary>
    /// Exactly three bytes - the shortest input the prefix test accepts. With <c>&gt;= 3</c>
    /// weakened to <c>&gt; 3</c> this is not recognised as a JPEG at all, and the caller is told
    /// the format is unsupported rather than that the file is truncated. Both are
    /// DocumentConversionException, so only the message separates them, and the message is the
    /// entire diagnostic value here: "unsupported format" sends someone converting their file,
    /// "no Start-Of-Frame" sends them looking at a truncated upload.
    /// </summary>
    [Fact]
    public void AJpegIsRecognisedFromExactlyThreeBytes()
    {
        var ex = Assert.Throws<DocumentConversionException>(
            () => ImageInspector.Inspect(Jpeg(new byte[] { 0xFF })));

        Assert.Contains("Start-Of-Frame", ex.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("Unsupported image format", ex.Message, StringComparison.Ordinal);
    }

    /// <summary>
    /// A segment header ending exactly at the last byte: <c>i + 4 == image.Length</c>. The guard is
    /// <c>&gt;</c>, so the two length bytes ARE there and are read; mutated to <c>&gt;=</c> it
    /// bails one byte early and reports a truncation that has not happened.
    /// </summary>
    [Fact]
    public void ASegmentHeaderEndingOnTheLastByteIsRead_NotReportedAsTruncated()
    {
        // FF D8 | FF E0 00 02 - an APP0 segment whose declared length (2) is just the length field
        // itself, so the walk steps past it and runs out of file looking for a frame header.
        var jpeg = Jpeg(new byte[] { 0xFF, 0xE0, 0x00, 0x02 });

        var ex = Assert.Throws<DocumentConversionException>(() => ImageInspector.Inspect(jpeg));

        Assert.Contains("Start-Of-Frame", ex.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("truncated inside a segment header", ex.Message, StringComparison.Ordinal);
    }


    /// <summary>
    /// A width resolving to <b>exactly</b> MaxCoordinateEmu, with a height that overflows.
    ///
    /// The blame logic asks whether the WIDTH is out of range to decide which argument to name, and
    /// it uses <c>&gt; MaxCoordinateEmu</c>. Sitting exactly on the ceiling, the width is legal, so
    /// the height is named. Mutated to <c>&gt;=</c> the width is judged bad and the caller is handed
    /// back the one number that was fine - which is the specific confusion that comment in
    /// ImageInspector exists to prevent.
    ///
    /// int.MaxValue / 12700.0 multiplies back to exactly int.MaxValue in IEEE 754 double, so this
    /// really does land on the boundary rather than near it. Checked, because most tidy-looking
    /// decimals nearby do not: 169093.2 resolves seven EMU short.
    /// </summary>
    [Fact]
    public void AWidthExactlyOnTheCeilingIsLegal_SoTheOverflowingHeightIsBlamed()
    {
        const double onTheCeiling = int.MaxValue / 12700.0;

        var ex = Assert.Throws<ArgumentOutOfRangeException>(
            () => ImageInspector.Resolve(
                new ImageInfo(ImageFormat.Png, 100, 50),
                widthPoints: onTheCeiling,
                heightPoints: onTheCeiling * 2));

        Assert.Equal("heightPoints", ex.ParamName);
    }
}
