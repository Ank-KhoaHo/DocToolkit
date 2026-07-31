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
}
