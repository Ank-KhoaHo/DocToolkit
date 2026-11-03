using DocToolkit;
using DocToolkit.Extensions.DependencyInjection;
using Xunit;

namespace DocToolkit.Extensions.DependencyInjection.Tests;

/// <summary>
/// <see cref="XlsxToPdfConverterService"/> carries <see cref="DocToolkitOptions.Fonts"/> to the core converter.
/// </summary>
/// <remarks>
/// <b>Proved by making the font INVALID</b>, as in <see cref="DocxToPdfConverterServiceTests"/> and
/// for the same reason: the renderer's "TrueType" refusal is the one host-independent sign that the
/// bytes travelled from the options object to the renderer. Until A125 this service called the core
/// overload that takes no fonts, so a configured font was silently dropped here.
/// </remarks>
public class XlsxToPdfConverterServiceTests
{
    private static byte[] Sample() => WorkbookEditor.Create("Sheet", new[] { new object?[] { "Plain Latin text." } });

    private static XlsxToPdfConverterService With(DocToolkitOptions options) =>
        new(new TestOptionsMonitor<DocToolkitOptions>(options));

    private static DocToolkitOptions InvalidFont() =>
        new() { Fonts = new PdfFontOptions("Fake", new byte[1024]) };

    [Fact]
    public void Convert_PassesTheConfiguredFontsThrough()
    {
        var ex = Assert.Throws<DocumentConversionException>(() => With(InvalidFont()).Convert(Sample()));

        Assert.Contains("TrueType", ex.InnerException!.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ConvertAsync_PassesTheConfiguredFontsThrough()
    {
        using var source = new MemoryStream(Sample());
        using var destination = new MemoryStream();

        var ex = await Assert.ThrowsAsync<DocumentConversionException>(
            () => With(InvalidFont()).ConvertAsync(source, destination));

        Assert.Contains("TrueType", ex.InnerException!.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Convert_WithNoFontsConfigured_IsTheCoreDefaultRender()
    {
        Assert.Equal(DocToolkit.XlsxToPdfConverter.Convert(Sample()), With(new DocToolkitOptions()).Convert(Sample()));
    }

    [Fact]
    public void Convert_ReadsTheOptionsOnEveryCall()
    {
        // A singleton: capturing the options at construction would make a font change need a
        // restart. Same reasoning as DocxToPdfConverterService.
        var options = new DocToolkitOptions();
        var sut = With(options);

        Assert.NotEmpty(sut.Convert(Sample()));

        options.Fonts = new PdfFontOptions("Fake", new byte[1024]);
        Assert.Throws<DocumentConversionException>(() => sut.Convert(Sample()));
    }
}
