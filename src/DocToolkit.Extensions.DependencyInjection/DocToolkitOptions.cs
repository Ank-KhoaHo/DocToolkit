namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>Options controlling the services registered by <c>AddDocToolkit</c> (added in a later task of this plan).</summary>
public sealed class DocToolkitOptions
{
    /// <summary>
    /// When true, HTML-to-DOCX and HTML-to-PDF conversion download images referenced by absolute
    /// <c>http</c>/<c>https</c> URLs. This issues outbound network requests - do not enable it in
    /// an air-gapped environment. Default: <c>false</c>.
    /// </summary>
    public bool AllowRemoteImageDownload { get; set; }
}
