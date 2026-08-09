using Microsoft.Extensions.Options;

namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>
/// Default <see cref="IHtmlToPdfConverter"/>, delegating to
/// <see cref="DocToolkit.HtmlToPdfConverter"/> with <see cref="DocToolkitOptions.AllowRemoteImageDownload"/>
/// and <see cref="DocToolkitOptions.RemoteImage"/>.
/// </summary>
internal sealed class HtmlToPdfConverterService : IHtmlToPdfConverter
{
    private readonly IOptionsMonitor<DocToolkitOptions> _options;

    // IOptionsMonitor, not IOptions, and CurrentValue read PER CALL rather than captured in the
    // constructor. These services are singletons, so `IOptions<T>.Value` is resolved once for the
    // lifetime of the container - which means a configuration reload silently does nothing.
    //
    // That matters more here than the usual "nice to have reload" argument, because the option in
    // question is the ONLY switch that lets this library open a socket. Turning
    // AllowRemoteImageDownload off in configuration - as an incident response, say - looked like it
    // worked and did not take effect until the process restarted. Nothing reported that.
    //
    // Reading CurrentValue per call costs a volatile field read; it is not a per-conversion
    // allocation and does not need caching.
    public HtmlToPdfConverterService(IOptionsMonitor<DocToolkitOptions> options) => _options = options;

    // Same overload selection as HtmlToDocxConverterService, and for the same reason - see the
    // comment there. The fetch itself happens in the HTML stage either way; this converter only
    // composes that with the DOCX-to-PDF render.
    public Task<byte[]> ConvertAsync(string html, CancellationToken ct = default)
        => _options.CurrentValue.AllowRemoteImageDownload
            ? DocToolkit.HtmlToPdfConverter.ConvertAsync(html, _options.CurrentValue.RemoteImage, ct)
            : DocToolkit.HtmlToPdfConverter.ConvertAsync(html, false, ct);

    public Task ConvertAsync(string html, Stream destination, CancellationToken ct = default)
        => _options.CurrentValue.AllowRemoteImageDownload
            ? DocToolkit.HtmlToPdfConverter.ConvertAsync(html, _options.CurrentValue.RemoteImage, destination, ct)
            : DocToolkit.HtmlToPdfConverter.ConvertAsync(html, false, destination, ct);

    public Task<byte[]> ConvertAsync(string html, DocToolkit.PageSetup page, CancellationToken ct = default)
        => DocToolkit.HtmlToPdfConverter.ConvertAsync(html, page, ct);

    public Task ConvertAsync(string html, DocToolkit.PageSetup page, Stream destination, CancellationToken ct = default)
        => DocToolkit.HtmlToPdfConverter.ConvertAsync(html, page, destination, ct);
}
