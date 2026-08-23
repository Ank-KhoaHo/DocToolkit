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

    /// <summary>
    /// The page every producer lays out on when a call does not name one. Default:
    /// <see cref="DocToolkit.PageSetup.A4"/> - what the static API already uses, so leaving this
    /// alone changes nothing.
    ///
    /// Paper is an application-wide fact rather than a per-call one: a service producing documents
    /// for US readers wants Letter on all of them, and repeating that at every call site is how one
    /// of them ends up missing it.
    /// <code>
    /// services.AddDocToolkit(o => o.Page = PageSetup.Letter);
    /// </code>
    ///
    /// An explicit argument still wins: <c>ConvertAsync(html, PageSetup.A4)</c> produces A4 whatever
    /// this says, because a call naming a page is answering a narrower question than configuration.
    ///
    /// Assigned rather than configured in place, unlike <see cref="RemoteImage"/>, because
    /// <see cref="DocToolkit.PageSetup"/> is immutable - there is nothing to configure in place and
    /// no restrictive default a wholesale replacement could quietly drop.
    /// </summary>
    /// <exception cref="ArgumentNullException">
    /// The value is null.
    ///
    /// Note WHERE this surfaces: the configure delegate runs when the options are first
    /// materialised, not when <c>AddDocToolkit</c> is called or the service resolved - so a
    /// null assigned here throws out of the first conversion rather than out of startup.
    /// Measured, after an earlier version of this documentation asserted the opposite.
    /// </exception>
    public PageSetup Page
    {
        get => _page;
        set
        {
            ArgumentNullException.ThrowIfNull(value);
            _page = value;
        }
    }

    private PageSetup _page = PageSetup.A4;

    /// <summary>
    /// Fonts supplied for characters the PDF renderer cannot otherwise encode, applied to every
    /// conversion this container performs. <see langword="null"/> - the default - supplies none.
    /// </summary>
    /// <remarks>
    /// <b>Configured once rather than passed per call, which is a decision this layer makes rather
    /// than a signature it copies.</b> The core API takes fonts per conversion; needing them is a
    /// property of the deployment, not of the document - somebody converting Cyrillic needs the font
    /// for every document, not for some. That is the same reasoning that turned the core's per-call
    /// <c>allowRemoteImageDownload</c> into <see cref="AllowRemoteImageDownload"/> here.
    ///
    /// <b>Assigned rather than configured in place</b>, unlike <see cref="RemoteImage"/>, and for the
    /// opposite reason: <see cref="PdfFontOptions"/> is immutable and carries no defaults that could
    /// be lost by replacing it wholesale. There is nothing to protect.
    ///
    /// <code>
    /// services.AddDocToolkit(o =>
    ///     o.Fonts = new PdfFontOptions("Noto Sans", File.ReadAllBytes("NotoSans-Regular.ttf"))
    ///                   .Add("Noto Sans CJK", File.ReadAllBytes("NotoSansCJK-Regular.ttf")));
    /// </code>
    ///
    /// <b>Supply fonts covering everything your documents use, not only the script that failed.</b>
    /// They REPLACE the host's own fallbacks rather than adding to them, so too few is worse than
    /// none - measured over 99 real documents, one font rendered 63 where none rendered 71 and four
    /// rendered 77. See <see cref="PdfFontOptions"/> for the whole of that.
    ///
    /// <b>Applies to every converter that renders a PDF</b> - <see cref="IDocxToPdfConverter"/> and
    /// <see cref="IHtmlToPdfConverter"/> alike. It reached only the first until core 0.34.0, because
    /// no core overload carried fonts alongside page setup and the remote-image settings; wiring it
    /// anyway would have applied fonts only when neither of the others was in play, and a setting
    /// that silently stops taking effect depending on unrelated configuration is worse than one that
    /// is documented as absent. <c>HtmlToPdfOptions</c> closed that, and this option now reaches both.
    /// </remarks>
    public PdfFontOptions? Fonts { get; set; }
}
