using System.IO;
using System.Linq;
using DocToolkit;
using DocToolkit.Extensions.DependencyInjection;
using Xunit;

namespace DocToolkit.Extensions.DependencyInjection.Tests;

public class DocxToPdfConverterServiceTests
{
    [Fact]
    public async Task Convert_ProducesAPdf()
    {
        var sut = new DocxToPdfConverterService();
        var docx = await HtmlToDocxConverter.ConvertAsync("<h1>Invoice</h1><p>Total: 18,100.00</p>");

        var pdf = sut.Convert(docx);

        Assert.Equal(new byte[] { 0x25, 0x50, 0x44, 0x46, 0x2D }, pdf.Take(5).ToArray()); // "%PDF-"
        Assert.True(pdf.Length > 200, $"expected a real PDF, got {pdf.Length} bytes");
    }

    [Fact]
    public void Convert_RejectsEmptyInput()
    {
        var sut = new DocxToPdfConverterService();

        Assert.Throws<ArgumentException>(() => sut.Convert(Array.Empty<byte>()));
    }

    [Fact]
    public async Task ConvertAsync_Stream_MatchesTheByteArrayOverload()
    {
        var sut = new DocxToPdfConverterService();
        var docx = await HtmlToDocxConverter.ConvertAsync("<h1>Invoice</h1><p>Total: 18,100.00</p>");

        var expected = sut.Convert(docx);

        using var destination = new MemoryStream();
        await sut.ConvertAsync(new MemoryStream(docx), destination);

        Assert.Equal(expected, destination.ToArray());
    }

    [Fact]
    public async Task ConvertAsync_HonorsCancellation()
    {
        var sut = new DocxToPdfConverterService();
        using var cts = new CancellationTokenSource();
        cts.Cancel();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(
            () => sut.ConvertAsync(new MemoryStream(), new MemoryStream(), cts.Token));
    }
}
