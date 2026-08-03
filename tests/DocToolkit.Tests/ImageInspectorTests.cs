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

    [Fact]
    public void ContentType_FollowsTheDetectedFormat()
    {
        Assert.Equal("image/png", new ImageInfo(ImageFormat.Png, 1, 1).ContentType);
        Assert.Equal("image/jpeg", new ImageInfo(ImageFormat.Jpeg, 1, 1).ContentType);
    }
}
