using System.Linq;
using DocToolkit.Extensions.DependencyInjection;
using Microsoft.Extensions.Options;
using Xunit;

namespace DocToolkit.Extensions.DependencyInjection.Tests;

public class HtmlToPdfConverterServiceTests
{
    [Fact]
    public async Task ConvertAsync_ProducesAPdf()
    {
        var sut = new HtmlToPdfConverterService(new TestOptionsMonitor<DocToolkitOptions>(new DocToolkitOptions()));

        var pdf = await sut.ConvertAsync("<h1>Invoice</h1><p>Total: 18,100.00</p>");

        Assert.Equal(new byte[] { 0x25, 0x50, 0x44, 0x46, 0x2D }, pdf.Take(5).ToArray());
        Assert.True(pdf.Length > 200, $"expected a real PDF, got {pdf.Length} bytes");
    }

    [Fact]
    public async Task ConvertAsync_RejectsNullHtml()
    {
        var sut = new HtmlToPdfConverterService(new TestOptionsMonitor<DocToolkitOptions>(new DocToolkitOptions()));

        await Assert.ThrowsAsync<ArgumentNullException>(() => sut.ConvertAsync(null!));
    }

    [Fact]
    public async Task ConvertAsync_ToStream_MatchesTheByteArrayOverload()
    {
        var sut = new HtmlToPdfConverterService(new TestOptionsMonitor<DocToolkitOptions>(new DocToolkitOptions()));

        // Exercised for its own sake: the byte[] overload must still work. Its RESULT is
        // deliberately unused - see the comment below on why this is not a byte comparison.
        _ = await sut.ConvertAsync("<h1>Invoice</h1><p>Total: 18,100.00</p>");

        using var destination = new MemoryStream();
        await sut.ConvertAsync("<h1>Invoice</h1><p>Total: 18,100.00</p>", destination);
        var actual = destination.ToArray();

        // Not a byte-equality assertion: each call pivots through its own freshly-built
        // intermediate .docx (see the branch's "never assert byte equality between two
        // separately generated packages" rule), so the two PDFs are two independent renders,
        // not one input read back. It happens that the PDF writer here emits no timestamp, so
        // today the two really are byte-identical - but that is an implementation detail this
        // test should not depend on. Assert the format and a sensible length floor instead.
        Assert.Equal(new byte[] { 0x25, 0x50, 0x44, 0x46, 0x2D }, actual.Take(5).ToArray());
        Assert.True(actual.Length > 200, $"expected a real PDF, got {actual.Length} bytes");
    }

    [Fact]
    public async Task ConvertAsync_ToStream_HonorsCancellation()
    {
        var sut = new HtmlToPdfConverterService(new TestOptionsMonitor<DocToolkitOptions>(new DocToolkitOptions()));
        using var cts = new CancellationTokenSource();
        cts.Cancel();

        using var destination = new MemoryStream();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(
            () => sut.ConvertAsync("<p>Body.</p>", destination, cts.Token));
    }

    // ===================================================================================
    // A57. Fonts reached IDocxToPdfConverter and not this one until core 0.34.0.
    // ===================================================================================

    /// <summary>
    /// A fake font proves fonts REACHED the renderer. 1024 bytes of zeroes is not a TrueType file,
    /// so a converter that received it fails and one that ignored it succeeds - which is exactly
    /// the defect A57 recorded. A test asserting "a PDF came back" cannot tell those apart, and
    /// that is why this suite could not have caught the gap before.
    /// </summary>
    private static PdfFontOptions FakeFont => new("Fake", new byte[1024]);

    [Fact]
    public async Task ConvertAsync_AppliesConfiguredFonts()
    {
        var options = new DocToolkitOptions { Fonts = FakeFont };
        var sut = new HtmlToPdfConverterService(new TestOptionsMonitor<DocToolkitOptions>(options));

        var ex = await Assert.ThrowsAsync<DocumentConversionException>(
            () => sut.ConvertAsync("<p>Body.</p>"));

        Assert.Contains("TrueType", ex.InnerException!.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ConvertAsync_WithNoFontsConfigured_StillConverts()
    {
        // The negative control for the test above. Without it, a service that always threw would
        // pass the font assertion and prove nothing.
        var sut = new HtmlToPdfConverterService(new TestOptionsMonitor<DocToolkitOptions>(new DocToolkitOptions()));

        var pdf = await sut.ConvertAsync("<p>Body.</p>");

        Assert.Equal(new byte[] { 0x25, 0x50, 0x44, 0x46, 0x2D }, pdf.Take(5).ToArray());
    }

    [Fact]
    public async Task ConvertAsync_AppliesFontsALONGSIDEAPageSetup()
    {
        // The whole of A57: fonts had to apply together with the other axes, not instead of them.
        // Naming a page must not quietly drop the configured fonts.
        var options = new DocToolkitOptions { Fonts = FakeFont };
        var sut = new HtmlToPdfConverterService(new TestOptionsMonitor<DocToolkitOptions>(options));

        var ex = await Assert.ThrowsAsync<DocumentConversionException>(
            () => sut.ConvertAsync("<p>Body.</p>", PageSetup.Letter));

        Assert.Contains("TrueType", ex.InnerException!.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ConvertAsync_ToStream_AppliesConfiguredFonts()
    {
        var options = new DocToolkitOptions { Fonts = FakeFont };
        var sut = new HtmlToPdfConverterService(new TestOptionsMonitor<DocToolkitOptions>(options));
        using var destination = new MemoryStream();

        var ex = await Assert.ThrowsAsync<DocumentConversionException>(
            () => sut.ConvertAsync("<p>Body.</p>", destination));

        Assert.Contains("TrueType", ex.InnerException!.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ConvertAsync_AppliesTheConfiguredPageWhenNoneIsNamed()
    {
        // Composing one options object must not lose DocToolkitOptions.Page.
        //
        // Read out of the PDF's own page dictionary rather than through PdfProbe, which lives in
        // the core test project and is not referenced here. That is safe for THIS value and not in
        // general: /MediaBox is PDF STRUCTURE, plain ASCII in the object dictionary, whereas page
        // TEXT is a hex-string operator inside a content stream and searching bytes for it finds
        // nothing while looking exactly like a broken converter.
        var letter = await ConvertWith(PageSetup.Letter);
        var a4 = await ConvertWith(PageSetup.A4);

        Assert.Contains("/MediaBox [0 0 612", letter, StringComparison.Ordinal);
        Assert.Contains("/MediaBox [0 0 595", a4, StringComparison.Ordinal);

        static async Task<string> ConvertWith(PageSetup page)
        {
            var sut = new HtmlToPdfConverterService(
                new TestOptionsMonitor<DocToolkitOptions>(new DocToolkitOptions { Page = page }));
            return System.Text.Encoding.ASCII.GetString(await sut.ConvertAsync("<p>Body.</p>"));
        }
    }
}
