using System.Net;
using System.Net.Sockets;
using HtmlToOpenXml;
using HtmlToOpenXml.IO;

namespace DocToolkit;

/// <summary>
/// The only component in this library that opens a socket. It exists so that the remote-image
/// opt-in is bounded rather than unbounded: HtmlToOpenXml's own <c>DefaultWebRequest</c> speaks
/// <c>file://</c> as well as http, downloads through a process-wide static <c>HttpClient</c> whose
/// headers it mutates per request, and imposes no timeout, host restriction or size limit.
/// Supplying this instead means that component is never constructed.
/// </summary>
internal sealed class GuardedResourceLoader : IWebRequest
{
    // One client for the process. Per-conversion clients exhaust sockets; mutating shared client
    // state per request is the HtmlToOpenXml 3.5.0 bug being routed around. Neither happens when
    // every per-request value lives on the request or on a linked token.
    private static readonly HttpClient Client = new(new SocketsHttpHandler
    {
        AllowAutoRedirect = false,      // each hop would need re-validating; an unvalidated hop is
                                        // the standard way past the address check below
        UseCookies = false,
        UseProxy = false,
    })
    {
        Timeout = System.Threading.Timeout.InfiniteTimeSpan,   // per-request, via a linked token
    };

    private readonly RemoteImageOptions _options;

    public GuardedResourceLoader(RemoteImageOptions options) => _options = options;

    /// <summary>
    /// http and https only. Refusing at the protocol gate is what removes the <c>file://</c>
    /// local-disclosure path, rather than relying on a rendering-mode setting to keep meaning what
    /// it means today.
    /// </summary>
    public bool SupportsProtocol(string protocol) =>
        string.Equals(protocol, "http", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(protocol, "https", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Loopback, private, link-local and unique-local addresses, including their IPv4-mapped forms.
    /// <c>169.254.169.254</c> falls out of the link-local check.
    /// </summary>
    internal static bool IsBlockedAddress(IPAddress address)
    {
        if (address.IsIPv4MappedToIPv6)
            address = address.MapToIPv4();

        if (IPAddress.IsLoopback(address)) return true;

        if (address.AddressFamily == AddressFamily.InterNetwork)
        {
            var b = address.GetAddressBytes();
            return b[0] == 10                                     // 10.0.0.0/8
                || (b[0] == 172 && b[1] >= 16 && b[1] <= 31)      // 172.16.0.0/12
                || (b[0] == 192 && b[1] == 168)                   // 192.168.0.0/16
                || (b[0] == 169 && b[1] == 254)                   // 169.254.0.0/16 link-local
                || b[0] == 0                                      // 0.0.0.0/8
                || b[0] >= 224;                                   // multicast and reserved
        }

        return address.IsIPv6LinkLocal
            || address.IsIPv6SiteLocal
            || address.IsIPv6Multicast
            || (address.GetAddressBytes()[0] & 0xFE) == 0xFC;      // fc00::/7 unique-local
    }

    public async Task<Resource?> FetchAsync(Uri requestUri, CancellationToken cancellationToken = default)
    {
        if (!SupportsProtocol(requestUri.Scheme)) return null;

        if (_options.AllowedHosts.Count > 0 && !_options.AllowedHosts.Contains(requestUri.Host))
            return null;

        if (!_options.AllowPrivateAddresses && await IsBlockedHostAsync(requestUri.Host, cancellationToken)
                .ConfigureAwait(false))
            return null;

        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        timeout.CancelAfter(_options.Timeout);

        try
        {
            using var request = new HttpRequestMessage(HttpMethod.Get, requestUri);
            using var response = await Client
                .SendAsync(request, HttpCompletionOption.ResponseHeadersRead, timeout.Token)
                .ConfigureAwait(false);

            if (!response.IsSuccessStatusCode) return null;

            var body = await ReadCappedAsync(response, timeout.Token).ConfigureAwait(false);
            if (body is null) return null;

            return new Resource
            {
                Content = new MemoryStream(body),
                StatusCode = response.StatusCode,
            };
        }
        catch (Exception) when (!cancellationToken.IsCancellationRequested)
        {
            // A refused, timed-out or failed fetch skips the image; it never costs the caller
            // their document. The caller's own cancellation is not swallowed.
            return null;
        }
    }

    /// <summary>
    /// Reads at most <see cref="RemoteImageOptions.MaxBytesPerImage"/> bytes, counting as it goes.
    /// <c>Content-Length</c> is never trusted: a hostile server can declare 1 KB and send a
    /// gigabyte. Returns null if the cap is exceeded.
    /// </summary>
    private async Task<byte[]?> ReadCappedAsync(HttpResponseMessage response, CancellationToken ct)
    {
        await using var source = await response.Content.ReadAsStreamAsync(ct).ConfigureAwait(false);
        using var buffer = new MemoryStream();

        var chunk = new byte[8192];
        int read;
        while ((read = await source.ReadAsync(chunk, ct).ConfigureAwait(false)) > 0)
        {
            if (buffer.Length + read > _options.MaxBytesPerImage) return null;
            buffer.Write(chunk, 0, read);
        }

        return buffer.ToArray();
    }

    private static async Task<bool> IsBlockedHostAsync(string host, CancellationToken ct)
    {
        // A literal address needs no lookup - and must not get one, since resolving it would be a
        // second chance to be told something different.
        if (IPAddress.TryParse(host, out var literal)) return IsBlockedAddress(literal);

        try
        {
            var addresses = await Dns.GetHostAddressesAsync(host, ct).ConfigureAwait(false);
            // Any blocked address blocks the host: a name with both a public and a private A record
            // must not be reachable just because the public one was checked first.
            return addresses.Length == 0 || addresses.Any(IsBlockedAddress);
        }
        catch
        {
            return true;   // cannot resolve it, cannot vouch for it
        }
    }
}
