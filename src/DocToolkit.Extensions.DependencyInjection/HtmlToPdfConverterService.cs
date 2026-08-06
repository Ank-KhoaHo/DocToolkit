using Microsoft.Extensions.Options;

namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>
/// Default <see cref="IHtmlToPdfConverter"/>, delegating to
/// <see cref="DocToolkit.HtmlToPdfConverter"/> with <see cref="DocToolkitOptions.AllowRemoteImageDownload"/>
/// and <see cref="DocToolkitOptions.RemoteImage"/>.
/// </summary>
internal sealed class HtmlToPdfConverterService : IHtmlToPdfConverter
{
    private readonly DocToolkitOptions _options;

    public HtmlToPdfConverterService(IOptions<DocToolkitOptions> options) => _options = options.Value;

    // Same overload selection as HtmlToDocxConverterService, and for the same reason - see the
    // comment there. The fetch itself happens in the HTML stage either way; this converter only
    // composes that with the DOCX-to-PDF render.
    public Task<byte[]> ConvertAsync(string html, CancellationToken ct = default)
        => _options.AllowRemoteImageDownload
            ? DocToolkit.HtmlToPdfConverter.ConvertAsync(html, _options.RemoteImage, ct)
            : DocToolkit.HtmlToPdfConverter.ConvertAsync(html, false, ct);

    public Task ConvertAsync(string html, Stream destination, CancellationToken ct = default)
        => _options.AllowRemoteImageDownload
            ? DocToolkit.HtmlToPdfConverter.ConvertAsync(html, _options.RemoteImage, destination, ct)
            : DocToolkit.HtmlToPdfConverter.ConvertAsync(html, false, destination, ct);
}
