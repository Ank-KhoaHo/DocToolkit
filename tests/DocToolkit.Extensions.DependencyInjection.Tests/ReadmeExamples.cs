using DocToolkit;
using DocToolkit.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Options;
using Xunit;

namespace DocToolkit.Extensions.DependencyInjection.Tests;

#region readme-di-consume
public class InvoiceService
{
    private readonly IHtmlToDocxConverter _toDocx;
    private readonly IHtmlToPdfConverter _toPdf;

    public InvoiceService(IHtmlToDocxConverter toDocx, IHtmlToPdfConverter toPdf)
    {
        _toDocx = toDocx;
        _toPdf = toPdf;
    }

    public Task<byte[]> RenderDocxAsync(string html) => _toDocx.ConvertAsync(html);
    public Task<byte[]> RenderAsync(string html) => _toPdf.ConvertAsync(html);
}
#endregion

/// <summary>
/// The extensions README's code blocks, as tests - the DI-flavoured counterpart to
/// tests/DocToolkit.Tests/ReadmeExamples.cs. They live HERE rather than there because this
/// project references Ank.DocToolkit as a published PackageReference, never a ProjectReference -
/// so a snippet that compiles here is proven against what an external consumer's restore
/// actually gets, not against whatever is currently on main. See scripts/gen-readme-snippets.py.
///
/// Setup ABOVE the region, assertions BELOW it - the reader sees only the capability.
/// </summary>
public class ReadmeExamples
{
    [Fact]
    public void RegistrationExample()
    {
        var services = new ServiceCollection();

        #region readme-di-registration
        services.AddDocToolkit();

        // Or opt in to remote image download for HTML->DOCX/PDF. This still succeeds in an
        // air-gapped environment - an unreachable host leaves that image out rather than failing
        // the conversion.
        services.AddDocToolkit(o => o.AllowRemoteImageDownload = true);
        #endregion

        using var provider = services.BuildServiceProvider();

        Assert.NotNull(provider.GetService<IHtmlToDocxConverter>());
        Assert.True(provider.GetRequiredService<IOptions<DocToolkitOptions>>().Value.AllowRemoteImageDownload);
    }

    [Fact]
    public void BoundingOptionsExample()
    {
        var services = new ServiceCollection();

        #region readme-di-options
        services.AddDocToolkit(o =>
        {
            o.AllowRemoteImageDownload = true;
            o.RemoteImage.Timeout = TimeSpan.FromSeconds(3);
            o.RemoteImage.AllowedHosts.Add("cdn.example.com");   // empty means "any public host"
        });
        #endregion

        using var provider = services.BuildServiceProvider();
        var options = provider.GetRequiredService<IOptions<DocToolkitOptions>>().Value;

        Assert.True(options.AllowRemoteImageDownload);
        Assert.Equal(TimeSpan.FromSeconds(3), options.RemoteImage.Timeout);
        Assert.Contains("cdn.example.com", options.RemoteImage.AllowedHosts);
    }

    [Fact]
    public async Task ConsumeExample()
    {
        using var provider = new ServiceCollection().AddDocToolkit().BuildServiceProvider();
        var invoiceService = new InvoiceService(
            provider.GetRequiredService<IHtmlToDocxConverter>(),
            provider.GetRequiredService<IHtmlToPdfConverter>());

        byte[] docx = await invoiceService.RenderDocxAsync("<h1>Invoice</h1>");
        byte[] pdf = await invoiceService.RenderAsync("<h1>Invoice</h1>");

        Assert.Contains("Invoice", DocxEditor.ExtractText(docx), StringComparison.Ordinal);
        Assert.Equal(1, PdfEditor.PageCount(pdf));
    }

    [Fact]
    public async Task DefaultPageSetupExample()
    {
        var services = new ServiceCollection();

        #region readme-di-page-setup
        services.AddDocToolkit(o => o.Page = PageSetup.Letter);
        #endregion

        using var provider = services.BuildServiceProvider();
        byte[] docx = await provider.GetRequiredService<IHtmlToDocxConverter>().ConvertAsync("<p>Hi</p>");

        Assert.Equal(LetterWidth, PageWidthOf(docx));
    }

    private const string LetterWidth = "12240";

    /// <summary>
    /// The page width in twentieths of a point, read directly out of the package's document.xml -
    /// deliberately not through DocToolkit's own reader, so this stays independent of the library
    /// under test. Matched on the attribute NAME only; the digits start one character later, past
    /// the opening quote.
    /// </summary>
    private static string PageWidthOf(byte[] docx)
    {
        using var archive = new System.IO.Compression.ZipArchive(new MemoryStream(docx));
        using var reader = new StreamReader(archive.GetEntry("word/document.xml")!.Open());
        var xml = reader.ReadToEnd();

        const string marker = "w:w=";
        var at = xml.IndexOf(marker, StringComparison.Ordinal);
        if (at < 0) return "(no w:w in document.xml)";

        var digits = xml.AsSpan(at + marker.Length + 1);
        var length = 0;
        while (length < digits.Length && char.IsAsciiDigit(digits[length])) length++;

        return digits[..length].ToString();
    }
}
