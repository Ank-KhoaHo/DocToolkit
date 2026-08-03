using System.IO;
using System.Linq;
using DocToolkit;
using DocToolkit.Extensions.DependencyInjection;
using Microsoft.Extensions.Options;
using Xunit;

namespace DocToolkit.Extensions.DependencyInjection.Tests;

public class HtmlToDocxConverterServiceTests
{
    [Fact]
    public async Task ConvertAsync_ProducesADocxContainingTheGivenContent()
    {
        var sut = new HtmlToDocxConverterService(Options.Create(new DocToolkitOptions()));

        var docx = await sut.ConvertAsync("<h1>Title</h1><p>Body copy.</p>");

        Assert.Equal(new byte[] { 0x50, 0x4B, 0x03, 0x04 }, docx.Take(4).ToArray());
        Assert.Contains("Body copy.", DocxEditor.ExtractText(docx));
    }

    [Fact]
    public async Task ConvertAsync_RejectsNullHtml()
    {
        var sut = new HtmlToDocxConverterService(Options.Create(new DocToolkitOptions()));

        await Assert.ThrowsAsync<ArgumentNullException>(() => sut.ConvertAsync(null!));
    }

    [Fact]
    public async Task ConvertAsync_ToStream_MatchesTheByteArrayOverload()
    {
        var sut = new HtmlToDocxConverterService(Options.Create(new DocToolkitOptions()));

        var expected = await sut.ConvertAsync("<h1>Title</h1><p>Body copy.</p>");

        using var destination = new MemoryStream();
        await sut.ConvertAsync("<h1>Title</h1><p>Body copy.</p>", destination);
        var actual = destination.ToArray();

        // Parity is asserted on readable content rather than on bytes: building a .docx stamps
        // the package with fresh metadata, so two conversions of identical markup never produce
        // identical bytes - not even two calls to the same static method.
        Assert.Equal(new byte[] { 0x50, 0x4B, 0x03, 0x04 }, actual.Take(4).ToArray());
        Assert.Equal(DocxEditor.ExtractText(expected), DocxEditor.ExtractText(actual));
        Assert.Contains("Body copy.", DocxEditor.ExtractText(actual));
    }
}
