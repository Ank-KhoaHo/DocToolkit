using Microsoft.Extensions.Options;

namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>
/// Default <see cref="IHtmlToDocxConverter"/>, delegating to
/// <see cref="DocToolkit.HtmlToDocxConverter"/> with <see cref="DocToolkitOptions.AllowRemoteImageDownload"/>
/// and <see cref="DocToolkitOptions.RemoteImage"/>.
/// </summary>
internal sealed class HtmlToDocxConverterService : IHtmlToDocxConverter
{
    private readonly DocToolkitOptions _options;

    public HtmlToDocxConverterService(IOptions<DocToolkitOptions> options) => _options = options.Value;

    // Selecting the overload is the one decision this layer makes rather than delegating, and it
    // has to live here: the core API expresses "offline" and "bounded fetch" as two separate
    // overloads, while DI has a single options object configured once. The bool picks which
    // overload; RemoteImage supplies the bounds to the one that takes them. Passing `false` rather
    // than skipping the call keeps the offline path going through exactly the same core method it
    // always did.
    public Task<byte[]> ConvertAsync(string html, CancellationToken ct = default)
        => _options.AllowRemoteImageDownload
            ? DocToolkit.HtmlToDocxConverter.ConvertAsync(html, _options.RemoteImage, ct)
            : DocToolkit.HtmlToDocxConverter.ConvertAsync(html, false, ct);

    public Task ConvertAsync(string html, Stream destination, CancellationToken ct = default)
        => _options.AllowRemoteImageDownload
            ? DocToolkit.HtmlToDocxConverter.ConvertAsync(html, _options.RemoteImage, destination, ct)
            : DocToolkit.HtmlToDocxConverter.ConvertAsync(html, false, destination, ct);
}
