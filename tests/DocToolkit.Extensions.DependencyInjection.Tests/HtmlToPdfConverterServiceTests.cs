using System.IO;
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

        Assert.Equal(expected, destination.ToArray());
    }
}
