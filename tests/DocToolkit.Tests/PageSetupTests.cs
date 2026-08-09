using Xunit;

namespace DocToolkit.Tests;

/// <summary>
/// <see cref="PageSetup"/> on its own — geometry and validation, with no document built.
///
/// The XML and PDF ends of the chain are <c>PageSetupOutputTests</c>. Keeping them apart is what
/// lets these run without OpenXml and lets those assert on bytes rather than on properties.
/// </summary>
public class PageSetupTests
{
    // ISO A4 is 210 x 297 mm = 595.2756 x 841.8898 pt. The one-decimal values are what land on
    // Word's own 11906 x 16838 twentieths; whole points would not.
    [Fact]
    public void A4_IsPortraitIsoA4WithOneInchMargins()
    {
        var page = PageSetup.A4;

        Assert.Equal(595.3, page.WidthPoints, 3);
        Assert.Equal(841.9, page.HeightPoints, 3);
        Assert.Equal(72, page.TopPoints, 3);
        Assert.Equal(72, page.RightPoints, 3);
        Assert.Equal(72, page.BottomPoints, 3);
        Assert.Equal(72, page.LeftPoints, 3);
    }

    [Fact]
    public void Letter_IsPortraitUsLetterWithOneInchMargins()
    {
        var page = PageSetup.Letter;

        Assert.Equal(612, page.WidthPoints, 3);
        Assert.Equal(792, page.HeightPoints, 3);
        Assert.Equal(72, page.TopPoints, 3);
    }

    [Fact]
    public void Custom_KeepsTheGivenSizeAndDefaultsMarginsToOneInch()
    {
        var page = PageSetup.Custom(300, 400);

        Assert.Equal(300, page.WidthPoints, 3);
        Assert.Equal(400, page.HeightPoints, 3);
        Assert.Equal(72, page.LeftPoints, 3);
    }

    [Fact]
    public void Landscape_SwapsWidthAndHeightAndLeavesMarginsAlone()
    {
        var page = PageSetup.A4.WithMargins(10, 20, 30, 40).Landscape();

        Assert.Equal(841.9, page.WidthPoints, 3);
        Assert.Equal(595.3, page.HeightPoints, 3);
        Assert.Equal(10, page.TopPoints, 3);
        Assert.Equal(20, page.RightPoints, 3);
        Assert.Equal(30, page.BottomPoints, 3);
        Assert.Equal(40, page.LeftPoints, 3);
    }

    // Landscape() swaps rather than "makes the long side horizontal", so calling it twice returns
    // to portrait. This pins the semantics: an implementation that normalised instead would be
    // idempotent, and this test is the only thing that tells the two apart.
    [Fact]
    public void Landscape_TwiceReturnsToTheOriginalOrientation()
    {
        var page = PageSetup.A4.Landscape().Landscape();

        Assert.Equal(595.3, page.WidthPoints, 3);
        Assert.Equal(841.9, page.HeightPoints, 3);
    }

    [Fact]
    public void Landscape_DoesNotMutateTheOriginal()
    {
        var portrait = PageSetup.A4;

        _ = portrait.Landscape();

        Assert.Equal(595.3, portrait.WidthPoints, 3);
    }

    [Fact]
    public void WithMargins_UniformSetsAllFour()
    {
        var page = PageSetup.A4.WithMargins(36);

        Assert.Equal(36, page.TopPoints, 3);
        Assert.Equal(36, page.RightPoints, 3);
        Assert.Equal(36, page.BottomPoints, 3);
        Assert.Equal(36, page.LeftPoints, 3);
    }

    // The four-argument order is top, right, bottom, left - clockwise from the top, as CSS does it.
    // Getting it wrong produces a document that opens fine and is laid out wrongly, so the order is
    // asserted with four distinct values rather than left to the doc comment.
    [Fact]
    public void WithMargins_FourArgumentsAreTopRightBottomLeftClockwise()
    {
        var page = PageSetup.A4.WithMargins(1, 2, 3, 4);

        Assert.Equal(1, page.TopPoints, 3);
        Assert.Equal(2, page.RightPoints, 3);
        Assert.Equal(3, page.BottomPoints, 3);
        Assert.Equal(4, page.LeftPoints, 3);
    }

    [Fact]
    public void WithMargins_KeepsTheSize()
    {
        var page = PageSetup.Custom(300, 400).WithMargins(10);

        Assert.Equal(300, page.WidthPoints, 3);
        Assert.Equal(400, page.HeightPoints, 3);
    }

    [Fact]
    public void WithMargins_DoesNotMutateTheOriginal()
    {
        var original = PageSetup.A4;

        _ = original.WithMargins(10);

        Assert.Equal(72, original.TopPoints, 3);
    }

    [Theory]
    [InlineData(0, 400)]
    [InlineData(-1, 400)]
    [InlineData(300, 0)]
    [InlineData(300, -1)]
    [InlineData(double.NaN, 400)]
    [InlineData(300, double.NaN)]
    public void Custom_RejectsNonPositiveOrNaNDimensions(double width, double height)
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => PageSetup.Custom(width, height));
    }

    // Width x 20 must fit the integer the OOXML attributes hold. Without this the value silently
    // wraps and the document is schema-invalid rather than rejected - the same failure
    // ImageInspector needed an overflow guard for.
    [Theory]
    [InlineData(double.MaxValue, 400)]
    [InlineData(300, double.MaxValue)]
    [InlineData(double.PositiveInfinity, 400)]
    public void Custom_RejectsDimensionsThatOverflowTwentiethsOfAPoint(double width, double height)
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => PageSetup.Custom(width, height));
    }

    [Fact]
    public void Custom_NamesTheOffendingParameter()
    {
        var ex = Assert.Throws<ArgumentOutOfRangeException>(() => PageSetup.Custom(0, 400));

        Assert.Equal("widthPoints", ex.ParamName);
    }

    [Theory]
    [InlineData(-1, 0, 0, 0, "topPoints")]
    [InlineData(0, -1, 0, 0, "rightPoints")]
    [InlineData(0, 0, -1, 0, "bottomPoints")]
    [InlineData(0, 0, 0, -1, "leftPoints")]
    [InlineData(double.NaN, 0, 0, 0, "topPoints")]
    public void WithMargins_RejectsNegativeOrNaNMarginsAndNamesTheParameter(
        double top, double right, double bottom, double left, string expectedParam)
    {
        var ex = Assert.Throws<ArgumentOutOfRangeException>(
            () => PageSetup.A4.WithMargins(top, right, bottom, left));

        Assert.Equal(expectedParam, ex.ParamName);
    }

    [Fact]
    public void WithMargins_UniformRejectsANegativeValue()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => PageSetup.A4.WithMargins(-1));
    }

    // A margin pair wider than the page produces a document that opens and renders blank. Nothing
    // downstream errors, so this is the only place it can be caught.
    [Fact]
    public void WithMargins_RejectsHorizontalMarginsThatLeaveNoContentArea()
    {
        var page = PageSetup.Custom(100, 400);

        var ex = Assert.Throws<ArgumentException>(() => page.WithMargins(0, 60, 0, 40));

        Assert.Contains("left", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("right", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void WithMargins_RejectsVerticalMarginsThatLeaveNoContentArea()
    {
        var page = PageSetup.Custom(400, 100);

        var ex = Assert.Throws<ArgumentException>(() => page.WithMargins(60, 0, 40, 0));

        Assert.Contains("top", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("bottom", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    // Landscape() swaps the dimensions under margins that were valid in portrait, so it has to
    // re-check rather than assume. A tall page with big horizontal margins becomes a short one.
    [Fact]
    public void Landscape_RejectsASwapThatLeavesNoContentArea()
    {
        var page = PageSetup.Custom(400, 100).WithMargins(0, 150, 0, 150);

        Assert.Throws<ArgumentException>(() => page.Landscape());
    }

    [Fact]
    public void ToString_NamesTheSizeAndMargins()
    {
        Assert.Equal("595.3 x 841.9 pt, margins 72/72/72/72 pt", PageSetup.A4.ToString());
    }

    // The exact ceiling. `value > MaxPoints` mutated to `>=` survives unless something sits ON the
    // boundary - and the boundary is the whole point of the guard, since one twentieth of a point
    // further overflows the integer the OOXML attribute holds.
    [Fact]
    public void Custom_AcceptsExactlyTheLargestRepresentableSize()
    {
        const double Max = int.MaxValue / 20d;

        var page = PageSetup.Custom(Max, Max);

        Assert.Equal(Max, page.WidthPoints, 3);
    }

    [Fact]
    public void WithMargins_AcceptsExactlyTheLargestRepresentableMargin()
    {
        const double Max = int.MaxValue / 20d;

        // A margin at the ceiling needs a page larger still, or the content-area check fires first
        // and this would pass for the wrong reason.
        var page = PageSetup.Custom(Max, Max).WithMargins(0, Max / 4, 0, Max / 4);

        Assert.Equal(Max / 4, page.RightPoints, 3);
    }
}
