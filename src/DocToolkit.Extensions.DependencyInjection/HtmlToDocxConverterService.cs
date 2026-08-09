using Microsoft.Extensions.Options;

namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>
/// Default <see cref="IHtmlToDocxConverter"/>, delegating to
/// <see cref="DocToolkit.HtmlToDocxConverter"/> with <see cref="DocToolkitOptions.AllowRemoteImageDownload"/>
/// and <see cref="DocToolkitOptions.RemoteImage"/>.
/// </summary>
internal sealed class HtmlToDocxConverterService : IHtmlToDocxConverter
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
    public HtmlToDocxConverterService(IOptionsMonitor<DocToolkitOptions> options) => _options = options;

    // Selecting the overload is the one decision this layer makes rather than delegating, and it
    // has to live here: the core API expresses "offline" and "bounded fetch" as two separate
    // overloads, while DI has a single options object configured once. The bool picks which
    // overload; RemoteImage supplies the bounds to the one that takes them. Passing `false` rather
    // than skipping the call keeps the offline path going through exactly the same core method it
    // always did.
    // Routed through the page overload so DocToolkitOptions.Page applies to a call that did
    // not name one. Before the core grew a (page, options) overload this could not be done:
    // the remote-image path always laid out on A4, so enabling downloads would have silently
    // discarded the configured paper.
    public Task<byte[]> ConvertAsync(string html, CancellationToken ct = default)
        => ConvertAsync(html, _options.CurrentValue.Page, ct);

    public Task ConvertAsync(string html, Stream destination, CancellationToken ct = default)
        => ConvertAsync(html, _options.CurrentValue.Page, destination, ct);

    // The page argument beats DocToolkitOptions.Page, and the remote-image setting still
    // applies - which it did NOT before: naming a page used to delegate to the offline core
    // overload, quietly opting the call back out of fetching a consumer had enabled.
    public Task<byte[]> ConvertAsync(string html, DocToolkit.PageSetup page, CancellationToken ct = default)
        => _options.CurrentValue.AllowRemoteImageDownload
            ? DocToolkit.HtmlToDocxConverter.ConvertAsync(html, page, _options.CurrentValue.RemoteImage, ct)
            : DocToolkit.HtmlToDocxConverter.ConvertAsync(html, page, ct);

    public Task ConvertAsync(string html, DocToolkit.PageSetup page, Stream destination, CancellationToken ct = default)
        => _options.CurrentValue.AllowRemoteImageDownload
            ? DocToolkit.HtmlToDocxConverter.ConvertAsync(html, page, _options.CurrentValue.RemoteImage, destination, ct)
            : DocToolkit.HtmlToDocxConverter.ConvertAsync(html, page, destination, ct);
}
