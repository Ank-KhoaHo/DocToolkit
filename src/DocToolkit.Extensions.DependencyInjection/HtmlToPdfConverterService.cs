using Microsoft.Extensions.Options;

namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>
/// Default <see cref="IHtmlToPdfConverter"/>, delegating to
/// <see cref="DocToolkit.HtmlToPdfConverter"/> with <see cref="DocToolkitOptions.AllowRemoteImageDownload"/>,
/// <see cref="DocToolkitOptions.RemoteImage"/>, <see cref="DocToolkitOptions.Page"/> and
/// <see cref="DocToolkitOptions.Fonts"/> - all four on the same conversion.
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

    // ONE options object per call, composed here. This is the documented exception to the
    // pure-delegation rule: core expresses the axes as separate overloads while DI has a single
    // options object, and mapping between them is the one thing this layer decides.
    //
    // It reads CurrentValue EXACTLY ONCE per call. The previous shape read it two or three times
    // in a ternary, so a configuration reload landing mid-call could have taken the offline branch
    // while reading the fetching options - a race with no symptom.
    //
    // AllowRemoteImageDownload REMAINS THE ONLY SWITCH deciding whether anything is fetched. While
    // it is false, RemoteImage must not reach core, whatever it says. Inverting that would break
    // the offline premise, and a null RemoteImage is what opens no socket at all.
    private HtmlToPdfOptions Compose(DocToolkit.PageSetup? page)
    {
        var o = _options.CurrentValue;
        return new HtmlToPdfOptions
        {
            Page = page ?? o.Page,
            RemoteImage = o.AllowRemoteImageDownload ? o.RemoteImage : null,
            Fonts = o.Fonts,
        };
    }

    // Fonts reach this converter as of core 0.34.0, which is what A57 was about: before it, the
    // core had no signature carrying fonts alongside a page or remote images, so wiring them here
    // would have made them apply only when neither of the others was in play.
    public Task<byte[]> ConvertAsync(string html, CancellationToken ct = default)
        => DocToolkit.HtmlToPdfConverter.ConvertAsync(html, Compose(null), ct);

    public Task ConvertAsync(string html, Stream destination, CancellationToken ct = default)
        => DocToolkit.HtmlToPdfConverter.ConvertAsync(html, Compose(null), destination, ct);

    // The page argument beats DocToolkitOptions.Page, and the remote-image setting still applies -
    // which it did NOT before 0.31.0: naming a page used to delegate to the offline core overload,
    // quietly opting the call back out of fetching a consumer had enabled.
    public Task<byte[]> ConvertAsync(string html, DocToolkit.PageSetup page, CancellationToken ct = default)
        => DocToolkit.HtmlToPdfConverter.ConvertAsync(html, Compose(page), ct);

    public Task ConvertAsync(string html, DocToolkit.PageSetup page, Stream destination, CancellationToken ct = default)
        => DocToolkit.HtmlToPdfConverter.ConvertAsync(html, Compose(page), destination, ct);
}
