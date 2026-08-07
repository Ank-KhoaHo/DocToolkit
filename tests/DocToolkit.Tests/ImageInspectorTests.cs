namespace DocToolkit.Tests;

/// <summary>
/// The header parsers and the EMU arithmetic, tested without a document in sight.
///
/// Both are the kind of code that fails quietly: a byte-order slip gives plausible dimensions, and
/// a factor-of-ten in the EMU conversion produces a perfectly valid document containing a wrongly
/// sized image. Nothing downstream — not the schema validator, not any text assertion — would catch
/// either.
/// </summary>
public class ImageInspectorTests
{
    [Fact]
    public void Inspect_ReadsPngDimensionsFromTheIhdrChunk()
    {
        var info = ImageInspector.Inspect(ImageFixtures.Png(width: 2, height: 3));

        Assert.Equal(ImageFormat.Png, info.Format);
        Assert.Equal(2, info.WidthPx);
        Assert.Equal(3, info.HeightPx);
    }

    [Fact]
    public void Inspect_ReadsPngDimensionsNeedingMoreThanOneByte()
    {
        // 300 x 260 exercises the big-endian assembly and distinguishes width from height;
        // a byte-order or transposition bug survives 2 x 3 unnoticed.
        var info = ImageInspector.Inspect(ImageFixtures.Png(width: 300, height: 260));

        Assert.Equal(300, info.WidthPx);
        Assert.Equal(260, info.HeightPx);
    }

    [Fact]
    public void Inspect_ReadsJpegDimensionsByWalkingToTheStartOfFrame()
    {
        var info = ImageInspector.Inspect(ImageFixtures.Jpeg());

        Assert.Equal(ImageFormat.Jpeg, info.Format);
        Assert.Equal(256, info.WidthPx);
        Assert.Equal(144, info.HeightPx);
    }

    [Fact]
    public void Inspect_RejectsAnUnsupportedFormatByName()
    {
        var ex = Assert.Throws<DocumentConversionException>(
            () => ImageInspector.Inspect(ImageFixtures.Gif()));

        Assert.Contains("GIF", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Inspect_RejectsBytesThatAreNoImageAtAll()
    {
        var ex = Assert.Throws<DocumentConversionException>(
            () => ImageInspector.Inspect(new byte[] { 1, 2, 3, 4, 5, 6, 7, 8 }));

        Assert.Contains("PNG", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("JPEG", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Inspect_RejectsATruncatedPng()
    {
        var truncated = ImageFixtures.Png()[..12];

        Assert.Throws<DocumentConversionException>(() => ImageInspector.Inspect(truncated));
    }

    [Theory]
    // 96 DPI: one pixel is 9,525 EMU. 2 x 3 px -> 19,050 x 28,575.
    [InlineData(2, 3, null, null, 19050L, 28575L)]
    // widthPoints only: 1 pt = 12,700 EMU; height scales to keep 2:3.
    [InlineData(2, 3, 10.0, null, 127000L, 190500L)]
    // heightPoints only: width scales to keep 2:3.
    [InlineData(2, 3, null, 30.0, 254000L, 381000L)]
    // both: exactly what was asked for, aspect ratio deliberately ignored.
    [InlineData(2, 3, 10.0, 10.0, 127000L, 127000L)]
    public void Resolve_ConvertsToEmusAndPreservesAspectWhenOnlyOneSideIsGiven(
        int widthPx, int heightPx, double? widthPoints, double? heightPoints,
        long expectedWidthEmu, long expectedHeightEmu)
    {
        var info = new ImageInfo(ImageFormat.Png, widthPx, heightPx);

        var (width, height) = ImageInspector.Resolve(info, widthPoints, heightPoints);

        Assert.Equal(expectedWidthEmu, width);
        Assert.Equal(expectedHeightEmu, height);
    }

    [Theory]
    [InlineData(0.0, null)]
    [InlineData(-5.0, null)]
    [InlineData(null, 0.0)]
    [InlineData(null, -5.0)]
    public void Resolve_RejectsANonPositiveSize(double? widthPoints, double? heightPoints)
    {
        var info = new ImageInfo(ImageFormat.Png, 2, 3);

        Assert.Throws<ArgumentOutOfRangeException>(
            () => ImageInspector.Resolve(info, widthPoints, heightPoints));
    }

    /// <summary>
    /// The upper bound, which the lower bound above had no counterpart for. An extent is capped at
    /// 2,147,483,647 EMU — see <c>MaxCoordinateEmu</c> for why that is the SDK's limit rather than
    /// the spec's, which is four orders of magnitude higher.
    ///
    /// Only the LAST row tests the new bound. The <c>1e300</c> rows are caught by the pre-existing
    /// <c>&lt;= 0</c> check, because <c>(long)(1e300 * 12700)</c> saturates to
    /// <see cref="long.MinValue"/> — they pin that saturation, which is real, but deleting the upper
    /// bound entirely leaves all three of them green. Recorded so the theory is not read as four
    /// independent guards on one property.
    /// </summary>
    [Theory]
    [InlineData(1e300, null)]
    [InlineData(null, 1e300)]
    [InlineData(1e300, 1e300)]
    [InlineData(200_000.0, null)]
    public void Resolve_RejectsASizeBeyondWhatOoxmlCanRepresent(double? widthPoints, double? heightPoints)
    {
        var info = new ImageInfo(ImageFormat.Png, 2, 3);

        Assert.Throws<ArgumentOutOfRangeException>(
            () => ImageInspector.Resolve(info, widthPoints, heightPoints));
    }

    /// <summary>
    /// When both sides are given they are independent, so the exception must name the one actually
    /// out of range. Blaming <c>widthPoints</c> for an oversized <c>heightPoints</c> reports a value
    /// the caller can see is fine. The all-<c>1e300</c> row above cannot catch this, because there
    /// both sides overflow and either name would look right.
    /// </summary>
    [Theory]
    [InlineData(10.0, 1e300, "heightPoints")]
    [InlineData(1e300, 10.0, "widthPoints")]
    public void Resolve_NamesTheSizeThatIsActuallyOutOfRange(
        double widthPoints, double heightPoints, string expected)
    {
        var info = new ImageInfo(ImageFormat.Png, 2, 3);

        var ex = Assert.Throws<ArgumentOutOfRangeException>(
            () => ImageInspector.Resolve(info, widthPoints, heightPoints));

        Assert.Equal(expected, ex.ParamName);
    }

    /// <summary>
    /// The boundary is inclusive: the largest extent the validator accepts is still resolved.
    /// Without this, tightening the comparison to <c>&gt;=</c> would reject a legal value and every
    /// rejection test above would still pass.
    /// </summary>
    [Fact]
    public void Resolve_AcceptsTheLargestRepresentableSize()
    {
        var info = new ImageInfo(ImageFormat.Png, 2, 2);

        var (width, height) = ImageInspector.Resolve(info, 169_093.0, 169_093.0);

        Assert.Equal(2_147_481_100L, width);
        Assert.Equal(2_147_481_100L, height);
        Assert.True(width <= int.MaxValue, "the accepted maximum must satisfy the validator");
    }

    /// <summary>
    /// Multiply-before-divide in the single-dimension paths. Written the other way — scaling by a
    /// precomputed ratio — the result truncates one EMU low, and NOTHING else in this suite catches
    /// it: the 2x3 fixture used everywhere else yields an identical value under both orderings,
    /// because the rounding error falls below the double spacing at that magnitude. This fixture is
    /// chosen to discriminate, which matters now that Resolve has been restructured once.
    /// </summary>
    [Theory]
    [InlineData(13, 3, 72.0, null, 914400L, 211015L)]
    [InlineData(3, 13, null, 72.0, 211015L, 914400L)]
    public void Resolve_MultipliesBeforeDividing(
        int widthPx, int heightPx, double? widthPoints, double? heightPoints,
        long expectedWidthEmu, long expectedHeightEmu)
    {
        var info = new ImageInfo(ImageFormat.Png, widthPx, heightPx);

        var (width, height) = ImageInspector.Resolve(info, widthPoints, heightPoints);

        Assert.Equal(expectedWidthEmu, width);
        Assert.Equal(expectedHeightEmu, height);
    }

    /// <summary>
    /// The intrinsic path is bounded too, and not by caller input: dimensions come from the file's
    /// own header, so a crafted PNG claiming billions of pixels reaches the limit with no size
    /// argument passed at all. <c>WidthPx</c> is an <see cref="int"/>, so a header can claim up to
    /// <c>int.MaxValue * 9525 = 20,454,781,737,675</c> EMU — far above the 2,147,483,647 an extent
    /// may hold, so this must throw.
    ///
    /// No practical image is affected: the bound admits any image up to 225,457 px per side, which
    /// is a 50-gigapixel picture. Only a header lying about its dimensions gets here.
    ///
    /// This assertion has now been wrong in BOTH directions — first asserting a throw the code could
    /// not produce, then accepting a size the validator rejects. Both times the number, not the
    /// reasoning, was the error.
    /// </summary>
    [Fact]
    public void Resolve_RejectsAnIntrinsicSizeTooLargeForAnExtent()
    {
        var info = new ImageInfo(ImageFormat.Png, int.MaxValue, int.MaxValue);

        // Now REJECTED, and correctly so: 20,454,781,737,675 EMU is far above the 2,147,483,647 the
        // validator accepts. The original version of this test asserted the opposite bound and
        // passed only because the guard was four orders of magnitude too loose.
        var ex = Assert.Throws<ArgumentOutOfRangeException>(
            () => ImageInspector.Resolve(info, null, null));

        Assert.Equal("info", ex.ParamName);
    }

    [Fact]
    public void ContentType_FollowsTheDetectedFormat()
    {
        Assert.Equal("image/png", new ImageInfo(ImageFormat.Png, 1, 1).ContentType);
        Assert.Equal("image/jpeg", new ImageInfo(ImageFormat.Jpeg, 1, 1).ContentType);
    }
}
