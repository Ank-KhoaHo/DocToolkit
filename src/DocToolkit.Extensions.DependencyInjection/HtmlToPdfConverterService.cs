using System.IO;
using Microsoft.Extensions.Options;

namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>
/// Default <see cref="IHtmlToPdfConverter"/>, delegating to
/// <see cref="DocToolkit.HtmlToPdfConverter"/> with <see cref="DocToolkitOptions.AllowRemoteImageDownload"/>.
/// </summary>
internal sealed class HtmlToPdfConverterService : IHtmlToPdfConverter
{
    private readonly DocToolkitOptions _options;

    public HtmlToPdfConverterService(IOptions<DocToolkitOptions> options) => _options = options.Value;

    public Task<byte[]> ConvertAsync(string html, CancellationToken ct = default)
        => DocToolkit.HtmlToPdfConverter.ConvertAsync(html, _options.AllowRemoteImageDownload, ct);

    public Task ConvertAsync(string html, Stream destination, CancellationToken ct = default)
        => DocToolkit.HtmlToPdfConverter.ConvertAsync(html, _options.AllowRemoteImageDownload, destination, ct);
}
