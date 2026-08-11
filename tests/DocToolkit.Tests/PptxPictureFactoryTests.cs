using Xunit;

namespace DocToolkit.Tests;

/// <summary>
/// The fit arithmetic on its own. Tested here rather than through a .pptx package because a wrong
/// number is the likely bug, and a package round trip would hide which number was wrong.
/// </summary>
public class PptxPictureFactoryTests
{
    [Fact]
    public void AnImageMatchingTheBoxRatioFillsItExactly()
    {
        var (x, y, cx, cy) = PptxPictureFactory.Fit(
            boxX: 100, boxY: 200, boxCx: 400, boxCy: 300, imageCx: 800, imageCy: 600);

        Assert.Equal(400, cx);
        Assert.Equal(300, cy);
        Assert.Equal(100, x);
        Assert.Equal(200, y);
    }

    [Fact]
    public void AWideImageIsLetterboxedTopAndBottom()
    {
        // 16:9 into a 4:3 box. Width binds, so height is short and the slack splits evenly.
        var (x, y, cx, cy) = PptxPictureFactory.Fit(
            boxX: 0, boxY: 0, boxCx: 400, boxCy: 300, imageCx: 1600, imageCy: 900);

        Assert.Equal(400, cx);
        Assert.Equal(225, cy);
        Assert.Equal(0, x);
        Assert.Equal(37, y);            // (300 - 225) / 2, integer division
    }

    [Fact]
    public void ATallImageIsPillarboxedLeftAndRight()
    {
        var (x, y, cx, cy) = PptxPictureFactory.Fit(
            boxX: 0, boxY: 0, boxCx: 400, boxCy: 300, imageCx: 600, imageCy: 1200);

        Assert.Equal(150, cx);
        Assert.Equal(300, cy);
        Assert.Equal(125, x);           // (400 - 150) / 2
        Assert.Equal(0, y);
    }

    [Fact]
    public void ASmallImageIsEnlargedToFit()
    {
        // Pins the both-directions decision from the spec. A "never upscale" rule would leave this
        // at 40x30 and the box mostly empty, which is the behaviour deliberately NOT chosen.
        var (_, _, cx, cy) = PptxPictureFactory.Fit(
            boxX: 0, boxY: 0, boxCx: 400, boxCy: 300, imageCx: 40, imageCy: 30);

        Assert.Equal(400, cx);
        Assert.Equal(300, cy);
    }

    [Fact]
    public void TheOffsetIsRelativeToTheBoxNotTheSlide()
    {
        // The box is at 1000,2000. A result of 1000,2000 proves the box origin was carried through;
        // 0,0 would mean the image landed in the slide corner, which is the failure that looks like
        // a rendering bug rather than a maths bug.
        var (x, y, _, _) = PptxPictureFactory.Fit(
            boxX: 1000, boxY: 2000, boxCx: 400, boxCy: 300, imageCx: 800, imageCy: 600);

        Assert.Equal(1000, x);
        Assert.Equal(2000, y);
    }

    [Theory]
    [InlineData(0, 300, 800, 600)]      // box has no width
    [InlineData(400, 0, 800, 600)]      // box has no height
    [InlineData(400, 300, 0, 600)]      // image has no width
    [InlineData(400, 300, 800, 0)]      // image has no height
    [InlineData(-400, 300, 800, 600)]   // negative box
    public void ADimensionThatIsNotPositiveIsRejected(long boxCx, long boxCy, long imageCx, long imageCy)
    {
        Assert.ThrowsAny<ArgumentException>(
            () => PptxPictureFactory.Fit(0, 0, boxCx, boxCy, imageCx, imageCy));
    }
}
