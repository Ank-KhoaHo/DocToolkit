using Microsoft.Extensions.Options;

namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>
/// Default <see cref="IHtmlToDocxConverter"/>, delegating to
/// <see cref="DocToolkit.HtmlToDocxConverter"/> with <see cref="DocToolkitOptions.AllowRemoteImageDownload"/>.
/// </summary>
internal sealed class HtmlToDocxConverterService : IHtmlToDocxConverter
{
    private readonly DocToolkitOptions _options;

    public HtmlToDocxConverterService(IOptions<DocToolkitOptions> options) => _options = options.Value;

    public Task<byte[]> ConvertAsync(string html, CancellationToken ct = default)
        => DocToolkit.HtmlToDocxConverter.ConvertAsync(html, _options.AllowRemoteImageDownload, ct);
}
