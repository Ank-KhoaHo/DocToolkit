namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>Options controlling the services registered by <see cref="ServiceCollectionExtensions.AddDocToolkit"/>.</summary>
public sealed class DocToolkitOptions
{
    /// <summary>
    /// When true, HTML-to-DOCX and HTML-to-PDF conversion download images referenced by absolute
    /// <c>http</c>/<c>https</c> URLs, bounded by <see cref="RemoteImage"/>. This issues outbound
    /// network requests. Because this is a process-wide setting rather than a per-call choice,
    /// enabling it opts every conversion in the application into fetching whatever URL the markup
    /// names - narrow that with <see cref="RemoteImage"/> if any caller converts untrusted HTML.
    /// Default: <c>false</c>.
    ///
    /// This is the only switch that decides whether anything is fetched: while it is <c>false</c>
    /// nothing is, no matter what <see cref="RemoteImage"/> says.
    ///
    /// An air-gapped environment no longer fails the conversion - a host that cannot be reached
    /// leaves that image out of the result, at a cost of up to
    /// <see cref="RemoteImageOptions.Timeout"/> per image.
    /// </summary>
    public bool AllowRemoteImageDownload { get; set; }

    /// <summary>
    /// Bounds applied to every image fetch, when - and only when -
    /// <see cref="AllowRemoteImageDownload"/> is <c>true</c>. Every default is already the
    /// restrictive one: loopback, private and link-local addresses are refused (including
    /// <c>169.254.169.254</c>, the cloud metadata endpoint), only <c>http</c> and <c>https</c> are
    /// spoken, redirects are not followed, and each fetch is capped at 10 seconds and 5 MB.
    ///
    /// Configured in place rather than assigned, so the restrictive defaults cannot be replaced
    /// wholesale by an object that missed one of them:
    /// <code>
    /// services.AddDocToolkit(o =>
    /// {
    ///     o.AllowRemoteImageDownload = true;
    ///     o.RemoteImage.Timeout = TimeSpan.FromSeconds(3);
    ///     o.RemoteImage.AllowedHosts.Add("cdn.example.com");
    /// });
    /// </code>
    ///
    /// <b>This is not a complete SSRF defence</b>; see <see cref="RemoteImageOptions"/> for the
    /// DNS-rebinding window it does not close.
    /// </summary>
    public RemoteImageOptions RemoteImage { get; } = new();
}
