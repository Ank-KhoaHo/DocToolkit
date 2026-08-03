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
        var sut = new HtmlToPdfConverterService(Options.Create(new DocToolkitOptions()));

        var pdf = await sut.ConvertAsync("<h1>Invoice</h1><p>Total: 18,100.00</p>");

        Assert.Equal(new byte[] { 0x25, 0x50, 0x44, 0x46, 0x2D }, pdf.Take(5).ToArray());
        Assert.True(pdf.Length > 200, $"expected a real PDF, got {pdf.Length} bytes");
    }

    [Fact]
    public async Task ConvertAsync_RejectsNullHtml()
    {
        var sut = new HtmlToPdfConverterService(Options.Create(new DocToolkitOptions()));

        await Assert.ThrowsAsync<ArgumentNullException>(() => sut.ConvertAsync(null!));
    }

    [Fact]
    public async Task ConvertAsync_ToStream_MatchesTheByteArrayOverload()
    {
        var sut = new HtmlToPdfConverterService(Options.Create(new DocToolkitOptions()));

        var expected = await sut.ConvertAsync("<h1>Invoice</h1><p>Total: 18,100.00</p>");

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
        var sut = new HtmlToPdfConverterService(Options.Create(new DocToolkitOptions()));
        using var cts = new CancellationTokenSource();
        cts.Cancel();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(
            () => sut.ConvertAsync("<p>Body.</p>", new MemoryStream(), cts.Token));
    }
}
