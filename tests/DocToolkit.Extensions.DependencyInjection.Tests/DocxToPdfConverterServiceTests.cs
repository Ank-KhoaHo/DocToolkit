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
        var sut = new DocxToPdfConverterService(new TestOptionsMonitor<DocToolkitOptions>(new DocToolkitOptions()));
        var docx = await HtmlToDocxConverter.ConvertAsync("<h1>Invoice</h1><p>Total: 18,100.00</p>");

        var pdf = sut.Convert(docx);

        Assert.Equal(new byte[] { 0x25, 0x50, 0x44, 0x46, 0x2D }, pdf.Take(5).ToArray()); // "%PDF-"
        Assert.True(pdf.Length > 200, $"expected a real PDF, got {pdf.Length} bytes");
    }

    [Fact]
    public void Convert_RejectsEmptyInput()
    {
        var sut = new DocxToPdfConverterService(new TestOptionsMonitor<DocToolkitOptions>(new DocToolkitOptions()));

        Assert.Throws<ArgumentException>(() => sut.Convert(Array.Empty<byte>()));
    }

    [Fact]
    public async Task ConvertAsync_Stream_MatchesTheByteArrayOverload()
    {
        var sut = new DocxToPdfConverterService(new TestOptionsMonitor<DocToolkitOptions>(new DocToolkitOptions()));
        var docx = await HtmlToDocxConverter.ConvertAsync("<h1>Invoice</h1><p>Total: 18,100.00</p>");

        var expected = sut.Convert(docx);

        using var source = new MemoryStream(docx);
        using var destination = new MemoryStream();
        await sut.ConvertAsync(source, destination);

        Assert.Equal(expected, destination.ToArray());
    }

    [Fact]
    public async Task ConvertAsync_HonorsCancellation()
    {
        var sut = new DocxToPdfConverterService(new TestOptionsMonitor<DocToolkitOptions>(new DocToolkitOptions()));
        using var cts = new CancellationTokenSource();
        cts.Cancel();

        using var source = new MemoryStream();
        using var destination = new MemoryStream();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(
            () => sut.ConvertAsync(source, destination, cts.Token));
    }

    /// <summary>
    /// The configured fonts reach the converter.
    /// </summary>
    /// <remarks>
    /// <b>Proved by making the font INVALID, which is the only host-independent way to see it.</b>
    /// A successful render with a real font would need a font file, and this project ships none - the
    /// licence and size reasons are on <see cref="PdfFontOptions"/>. The renderer validates font data
    /// eagerly, so its refusal is proof the bytes travelled from the options object through the
    /// service to the converter; nothing else in this package produces that message.
    ///
    /// Same shape as the core test for this, and the same reason.
    /// </remarks>
    [Fact]
    public void Convert_PassesTheConfiguredFontsThrough()
    {
        var options = new DocToolkitOptions { Fonts = new PdfFontOptions("Fake", new byte[1024]) };
        var sut = new DocxToPdfConverterService(new TestOptionsMonitor<DocToolkitOptions>(options));
        var docx = DocxEditor.Create([DocxBlock.Paragraph("Plain Latin text.")]);

        var ex = Assert.Throws<DocumentConversionException>(() => sut.Convert(docx));

        Assert.Contains("TrueType", ex.InnerException!.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Convert_WithNoFontsConfigured_BehavesAsBefore()
    {
        // The default must be indistinguishable from the old signature: nobody who does not
        // configure fonts should notice this option exists.
        var sut = new DocxToPdfConverterService(new TestOptionsMonitor<DocToolkitOptions>(new DocToolkitOptions()));

        Assert.NotEmpty(sut.Convert(DocxEditor.Create([DocxBlock.Paragraph("Plain text.")])));
    }

    [Fact]
    public void Convert_ReadsTheOptionsOnEveryCall()
    {
        // These services are singletons, so capturing at construction would make a configuration
        // change need a restart - and the moment somebody most wants to change configuration is when
        // something is going wrong. Same reasoning as HtmlToDocxConverterService.
        var options = new DocToolkitOptions();
        var monitor = new TestOptionsMonitor<DocToolkitOptions>(options);
        var sut = new DocxToPdfConverterService(monitor);
        var docx = DocxEditor.Create([DocxBlock.Paragraph("Plain text.")]);

        Assert.NotEmpty(sut.Convert(docx));

        options.Fonts = new PdfFontOptions("Fake", new byte[1024]);
        Assert.Throws<DocumentConversionException>(() => sut.Convert(docx));
    }
}
