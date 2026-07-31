namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>Options controlling the services registered by <see cref="ServiceCollectionExtensions.AddDocToolkit"/>.</summary>
public sealed class DocToolkitOptions
{
    /// <summary>
    /// When true, HTML-to-DOCX and HTML-to-PDF conversion download images referenced by absolute
    /// <c>http</c>/<c>https</c> URLs. This issues outbound network requests - do not enable it in
    /// an air-gapped environment. Because this is a process-wide setting rather than a per-call
    /// choice, enabling it opts every conversion in the application into fetching whatever URL the
    /// markup names - only enable it if no caller ever converts untrusted HTML. Default: <c>false</c>.
    /// </summary>
    public bool AllowRemoteImageDownload { get; set; }
}
