namespace DocToolkit.Tests;

public class PptxSlideTests
{
    [Fact]
    public void Titled_KeepsTheTitleAndBulletsInOrder()
    {
        var slide = PptxSlide.Titled("Q3 Results", "Revenue up 12%", "Costs flat");

        Assert.Equal("Q3 Results", slide.Title);
        Assert.Equal(new[] { "Revenue up 12%", "Costs flat" }, slide.Bullets);
    }

    [Fact]
    public void Titled_AcceptsATitleWithNoBullets()
    {
        var slide = PptxSlide.Titled("Section Break");

        Assert.Equal("Section Break", slide.Title);
        Assert.Empty(slide.Bullets);
    }

    /// <summary>
    /// Eager validation, matching DocxBlock's factories: the throw lands on the line that built the
    /// bad slide rather than later inside Create, where a caller assembling many slides would have
    /// no idea which one was wrong.
    /// </summary>
    [Fact]
    public void Titled_RejectsANullTitle()
    {
        var ex = Assert.Throws<ArgumentNullException>(() => PptxSlide.Titled(null!));
        Assert.Equal("title", ex.ParamName);
    }

    [Fact]
    public void Titled_RejectsANullBulletsArray()
    {
        var ex = Assert.Throws<ArgumentNullException>(() => PptxSlide.Titled("T", null!));
        Assert.Equal("bullets", ex.ParamName);
    }

    /// <summary>
    /// The message names the 1-based index, matching DocxBlock.Table's "Row 1 was null." A caller
    /// with twenty bullets needs to know which one.
    /// </summary>
    [Fact]
    public void Titled_RejectsANullBulletAndNamesIt()
    {
        var ex = Assert.Throws<ArgumentException>(
            () => PptxSlide.Titled("T", "first", null!, "third"));

        Assert.Equal("bullets", ex.ParamName);
        Assert.StartsWith("Bullet 2 was null.", ex.Message);
    }

    /// <summary>
    /// Materialised eagerly, so mutating the caller's array afterwards cannot change the slide.
    /// </summary>
    [Fact]
    public void Titled_MaterialisesBulletsImmediately()
    {
        var bullets = new[] { "original" };
        var slide = PptxSlide.Titled("T", bullets);

        bullets[0] = "mutated";

        Assert.Equal(new[] { "original" }, slide.Bullets);
    }
}
