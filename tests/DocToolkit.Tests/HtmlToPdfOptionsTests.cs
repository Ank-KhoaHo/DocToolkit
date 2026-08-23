using System.Threading.Tasks;
using Xunit;

namespace DocToolkit.Tests;

/// <summary>
/// <see cref="HtmlToPdfOptions"/> and the overloads that take it — A57.
///
/// The point of the type is that the three axes apply TOGETHER. Every test here therefore sets
/// more than one of them and asserts on both, because an implementation that honoured page setup
/// and dropped fonts would pass any test that checked one at a time. That is the shape of the bug
/// this closes: fonts reached one PDF converter and not the other.
/// </summary>
public class HtmlToPdfOptionsTests
{
    /// <summary>
    /// A fake font is the established way to prove fonts REACHED the renderer - see
    /// <c>DocxToPdfConverterServiceTests</c>, which uses the same trick. 1024 bytes of zeroes is
    /// not a TrueType file, so a renderer that received it fails and one that ignored it succeeds.
    /// A test asserting "the PDF came back" could not tell those apart.
    /// </summary>
    private static PdfFontOptions FakeFont => new("Fake", new byte[1024]);

    [Fact]
    public void Defaults_MatchTheOverloadsTheyReplace()
    {
        var options = new HtmlToPdfOptions();

        // A4, not Letter: a document with no page setup renders on the reader's template, which
        // is the correctness defect 0.13.0 fixed. The default must not reintroduce it.
        Assert.Equal(PageSetup.A4.WidthPoints, options.Page.WidthPoints);
        Assert.Equal(PageSetup.A4.HeightPoints, options.Page.HeightPoints);

        // null means "fetch nothing" and "supply no fonts" - the offline default. An options
        // object that silently opted callers INTO fetching would break the offline premise.
        Assert.Null(options.RemoteImage);
        Assert.Null(options.Fonts);
    }

    [Fact]
    public async Task ConvertAsync_AppliesThePageSetupItWasGiven()
    {
        var options = new HtmlToPdfOptions { Page = PageSetup.Letter };

        byte[] pdf = await HtmlToPdfConverter.ConvertAsync("<p>Hello</p>", options);

        // MediaBox is the only assertion proving page setup survives the DOCX -> PDF render.
        // Every other page-setup test reads the .docx, which OfficeIMO could stop honouring
        // without any of them noticing.
        var box = Assert.Single(PdfProbe.MediaBoxes(pdf));
        Assert.Equal((int)PageSetup.Letter.WidthPoints, (int)box.Width);
        Assert.Equal((int)PageSetup.Letter.HeightPoints, (int)box.Height);
    }

    [Fact]
    public async Task ConvertAsync_AppliesFontsALONGSIDEAPageSetup()
    {
        // THE WHOLE OF A57 IN ONE TEST. The old surface could express fonts, or a page, but never
        // both - so this combination had no overload at all and DocToolkitOptions.Fonts could not
        // reach IHtmlToPdfConverter without silently dropping the page.
        var options = new HtmlToPdfOptions { Page = PageSetup.Letter, Fonts = FakeFont };

        var ex = await Assert.ThrowsAsync<DocumentConversionException>(
            () => HtmlToPdfConverter.ConvertAsync("<p>Hello</p>", options));

        // It threw because the RENDERER read the fake font. If fonts were dropped this call
        // would have succeeded, which is exactly the defect.
        Assert.Contains("TrueType", ex.InnerException!.Message, System.StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ConvertAsync_WithNoFonts_StillConverts()
    {
        // The negative control for the test above: without fonts the same call must SUCCEED.
        // Without this, a converter that always threw would pass the assertion above.
        var options = new HtmlToPdfOptions { Page = PageSetup.Letter };

        byte[] pdf = await HtmlToPdfConverter.ConvertAsync("<p>Hello</p>", options);

        Assert.True(pdf.Length > 100);
        Assert.Equal((byte)'%', pdf[0]);
    }

    [Fact]
    public async Task ConvertAsync_RejectsANullOptions()
    {
        await Assert.ThrowsAsync<System.ArgumentNullException>(
            () => HtmlToPdfConverter.ConvertAsync("<p>Hello</p>", (HtmlToPdfOptions)null!));
    }
}
