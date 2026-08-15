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

    // Read once, here, rather than off the caller's RemoteImageOptions on each fetch. That object
    // is mutable and stays theirs, so a conversion that read it per fetch would be reading values
    // Validate() has not vouched for: a Timeout mutated negative mid-conversion would make
    // CancelAfter below throw out of FetchAsync rather than skip the image, and a MaxBytesPerImage
    // mutated mid-read would move the cap under a read already in flight. Snapshotting makes the
    // loader immutable once constructed, so the whole question stops existing.
    private readonly TimeSpan _timeout;
    private readonly long _maxBytesPerImage;
    private readonly bool _allowPrivateAddresses;
    private readonly HashSet<string> _allowedHosts;

    public GuardedResourceLoader(RemoteImageOptions options)
    {
        _timeout = options.Timeout;
        _maxBytesPerImage = options.MaxBytesPerImage;
        _allowPrivateAddresses = options.AllowPrivateAddresses;

        // The comparer is restated rather than inherited: RemoteImageOptions.AllowedHosts is
        // documented as case-insensitive, and a copy that silently became ordinal would narrow
        // the allow-list without any test noticing.
        _allowedHosts = new HashSet<string>(options.AllowedHosts, StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>
    /// http and https only. Refusing at the protocol gate is what removes the <c>file://</c>
    /// local-disclosure path, rather than relying on a rendering-mode setting to keep meaning what
    /// it means today.
    /// </summary>
    public bool SupportsProtocol(string protocol) =>
        string.Equals(protocol, "http", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(protocol, "https", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Loopback, private, link-local and unique-local addresses, including their IPv4-mapped forms
    /// and the IPv6 transition mechanisms that carry an IPv4 address inside them: IPv4-mapped,
    /// NAT64 (both the RFC 6052 well-known prefix and the RFC 8215 local-use prefix), 6to4, Teredo,
    /// IPv4-translated and IPv4-compatible. <c>169.254.169.254</c> falls out of the link-local check.
    /// This is not every IPv6 transition mechanism that exists (ISATAP and 6rd are not covered);
    /// it is exactly the set <see cref="TryGetEmbeddedIPv4"/> unwraps.
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
                || (b[0] == 100 && b[1] >= 64 && b[1] <= 127)     // 100.64.0.0/10 CGNAT
                || b[0] == 0                                      // 0.0.0.0/8
                || b[0] >= 224;                                   // multicast and reserved
        }

        // IsIPv4MappedToIPv6 above only covers ::ffff:0:0/96. NAT64, 6to4 and the older
        // IPv4-compatible/IPv4-translated forms all carry a real IPv4 address inside an IPv6 one
        // through a different well-known prefix, and every one of them is a documented bypass for
        // an address check that only inspects the outer /96: on an IPv6-only host behind
        // DNS64/NAT64 (AWS IPv6-only subnets, many Kubernetes clusters, mobile carriers), a
        // hostname resolving to 64:ff9b::a9fe:a9fe reaches 169.254.169.254 unless the embedded
        // address is extracted and checked in its own right.
        if (TryGetEmbeddedIPv4(address, out var embedded))
            return IsBlockedAddress(embedded);

        // :: (IPAddress.IPv6Any) is not checked here: it is all-zero, so it always matches the
        // IPv4-compatible ::/96 branch above first and is blocked via the embedded 0.0.0.0 - a
        // direct check here would never run. IsBlockedAddress_RefusesEveryPrivateForm("::") pins
        // that this still blocks it, just through the embedded-address path.
        return address.IsIPv6LinkLocal
            || address.IsIPv6SiteLocal
            || address.IsIPv6Multicast
            || (address.GetAddressBytes()[0] & 0xFE) == 0xFC;      // fc00::/7 unique-local
    }

    /// <summary>
    /// Extracts the IPv4 address embedded in <paramref name="address"/> if it uses one of the
    /// well-known IPv6 transition prefixes, so the caller can classify the real address rather than
    /// the IPv6 wrapper around it.
    /// </summary>
    private static bool TryGetEmbeddedIPv4(IPAddress address, out IPAddress embedded)
    {
        var b = address.GetAddressBytes();

        // NAT64 well-known prefix, RFC 6052: 64:ff9b::/96.
        if (b[0] == 0x00 && b[1] == 0x64 && b[2] == 0xFF && b[3] == 0x9B && IsZero(b, 4, 8))
        {
            embedded = ToIPv4(b, 12);
            return true;
        }

        // NAT64 local-use prefix, RFC 8215: 64:ff9b:1::/48 - for operators who run NAT64 inside
        // their own network rather than through the well-known /96 prefix above. A /48 prefix
        // leaves only 16 bits of address space before the embedded v4 address, so RFC 6052 §2.2
        // splits it around a "u" byte at position 8: bytes 6-7 carry the top two v4 octets, byte 8
        // is reserved, bytes 9-10 carry the bottom two. On an IPv6-only host whose NAT64 uses this
        // prefix instead of the global one, attacker-controlled DNS returning an address under it
        // reaches the same private ranges once unwrapped - 64:ff9b:1:a9fe:a9:fe00:: is this
        // encoding of 169.254.169.254.
        if (b[0] == 0x00 && b[1] == 0x64 && b[2] == 0xFF && b[3] == 0x9B && b[4] == 0x00 && b[5] == 0x01)
        {
            embedded = new IPAddress(new[] { b[6], b[7], b[9], b[10] });
            return true;
        }

        // 6to4, RFC 3056: 2002::/16, with the embedded v4 address at bytes 2-5.
        if (b[0] == 0x20 && b[1] == 0x02)
        {
            embedded = ToIPv4(b, 2);
            return true;
        }

        // Teredo, RFC 4380: 2001::/32. Bytes 12-15 carry the client's public IPv4 address,
        // obfuscated by XOR with 0xffffffff (so it does not appear in plaintext to routers along
        // the way) - unwrapping it means XOR-ing back rather than reading it directly the way the
        // other forms here do. 2001::5601:5601 decodes to 169.254.169.254 this way.
        if (b[0] == 0x20 && b[1] == 0x01 && b[2] == 0x00 && b[3] == 0x00)
        {
            embedded = new IPAddress(new[]
            {
                (byte)(b[12] ^ 0xFF), (byte)(b[13] ^ 0xFF), (byte)(b[14] ^ 0xFF), (byte)(b[15] ^ 0xFF),
            });
            return true;
        }

        // IPv4-translated (SIIT), historically written ::ffff:0:a.b.c.d: the first 8 bytes are
        // zero, followed by ffff and a zero group, then the embedded address. Distinct from the
        // IPv4-mapped ::ffff:a.b.c.d form handled by IsIPv4MappedToIPv6 above - that one has no
        // extra zero group before the address.
        if (IsZero(b, 0, 8) && b[8] == 0xFF && b[9] == 0xFF && b[10] == 0x00 && b[11] == 0x00)
        {
            embedded = ToIPv4(b, 12);
            return true;
        }

        // IPv4-compatible, RFC 4291 (deprecated but still parseable): ::/96.
        if (IsZero(b, 0, 12))
        {
            embedded = ToIPv4(b, 12);
            return true;
        }

        embedded = IPAddress.Any;
        return false;
    }

    private static bool IsZero(byte[] bytes, int offset, int count)
    {
        for (var i = offset; i < offset + count; i++)
        {
            if (bytes[i] != 0) return false;
        }

        return true;
    }

    private static IPAddress ToIPv4(byte[] ipv6Bytes, int offset) =>
        new(new[] { ipv6Bytes[offset], ipv6Bytes[offset + 1], ipv6Bytes[offset + 2], ipv6Bytes[offset + 3] });

    /// <summary>
    /// Any blocked address blocks the host: a name with both a public and a private A record must
    /// not be reachable just because the public one was checked first. Split out as a pure function
    /// so it can be tested directly with a crafted address list, without needing a DNS response
    /// that mixes public and private records to exist somewhere.
    /// </summary>
    internal static bool IsBlockedAddresses(IReadOnlyList<IPAddress> addresses) =>
        addresses.Count == 0 || addresses.Any(IsBlockedAddress);

    public async Task<Resource?> FetchAsync(Uri requestUri, CancellationToken cancellationToken = default)
    {
        // Host, never the URL: a query string routinely carries a signed token, and telemetry
        // leaves the machine. See the Telemetry class remarks.
        var host = requestUri.Host;
        using var activity = Telemetry.Source.StartActivity("RemoteImage.Fetch");
        activity?.SetTag("server.address", host);
        activity?.SetTag("url.scheme", requestUri.Scheme);
        // Port alongside host. Standard OpenTelemetry attribute, carries nothing sensitive
        // (a query string can hold a token; a port number cannot), and it is what tells two
        // services on the same host apart in a trace.
        activity?.SetTag("server.port", requestUri.Port);

        if (!SupportsProtocol(requestUri.Scheme))
        {
            Telemetry.RecordOutcome(activity, Telemetry.Outcomes.SchemeRefused, host);
            return null;
        }

        if (_allowedHosts.Count > 0 && !_allowedHosts.Contains(requestUri.Host))
        {
            Telemetry.RecordOutcome(activity, Telemetry.Outcomes.HostNotAllowed, host);
            return null;
        }

        // Created before the DNS lookup, not after, so Timeout bounds resolution too: an attacker
        // black-holing their own authoritative nameserver would otherwise stall every fetch for the
        // OS resolver timeout (5-12 s) regardless of this setting, and HtmlToOpenXml calls
        // FetchAsync twice per <img>, concurrently, which multiplies the cost.
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        timeout.CancelAfter(_timeout);

        try
        {
            if (!_allowPrivateAddresses &&
                await IsBlockedHostAsync(requestUri.Host, cancellationToken, timeout.Token)
                    .ConfigureAwait(false))
            {
                Telemetry.RecordOutcome(activity, Telemetry.Outcomes.BlockedAddress, host);
                return null;
            }

            using var request = new HttpRequestMessage(HttpMethod.Get, requestUri);
            using var response = await Client
                .SendAsync(request, HttpCompletionOption.ResponseHeadersRead, timeout.Token)
                .ConfigureAwait(false);

            if (!response.IsSuccessStatusCode)
            {
                activity?.SetTag("http.response.status_code", (int)response.StatusCode);
                Telemetry.RecordOutcome(activity, Telemetry.Outcomes.HttpError, host);
                return null;
            }

            var body = await ReadCappedAsync(response, timeout.Token).ConfigureAwait(false);
            if (body is null)
            {
                // ReadCappedAsync returns null for exactly one reason: the body went past
                // MaxBytesPerImage and was abandoned.
                Telemetry.RecordOutcome(activity, Telemetry.Outcomes.TooLarge, host);
                return null;
            }

            activity?.SetTag("doctoolkit.remote_image.bytes", body.Length);
            Telemetry.RemoteImageBytes.Record(
                body.Length, new KeyValuePair<string, object?>("server.address", host));
            Telemetry.RecordOutcome(activity, Telemetry.Outcomes.Ok, host);

            return new Resource
            {
                Content = new MemoryStream(body),
                StatusCode = response.StatusCode,
            };
        }
        catch (Exception ex) when (!cancellationToken.IsCancellationRequested)
        {
            // A refused, timed-out or failed fetch skips the image; it never costs the caller
            // their document. The caller's own cancellation is not swallowed.
            //
            // This is the case telemetry exists for: on an air-gapped host EVERY remote image lands
            // here, silently, and the document still succeeds. Without a span or a counter there is
            // nothing at all to tell a consumer their images never arrived.
            activity?.SetTag("exception.type", ex.GetType().FullName);
            Telemetry.RecordOutcome(activity, Telemetry.Outcomes.Failed, host);
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
        // Two awaits hide in the `await using` declaration form: the one that produces the stream,
        // and the DISPOSAL. Only the first can carry ConfigureAwait there, and the second is the
        // one that runs on the caller's context. Scoped with a block so both are configured.
        var source = await response.Content.ReadAsStreamAsync(ct).ConfigureAwait(false);
        await using (source.ConfigureAwait(false))
        {
            using var buffer = new MemoryStream();

            var chunk = new byte[8192];
            int read;
            while ((read = await source.ReadAsync(chunk, ct).ConfigureAwait(false)) > 0)
            {
                if (buffer.Length + read > _maxBytesPerImage) return null;
                buffer.Write(chunk, 0, read);
            }

            return buffer.ToArray();
        }
    }

    /// <param name="host">The host to resolve, or a literal address.</param>
    /// <param name="cancellationToken">
    /// The caller's own token, unlinked. Used only to tell the caller's cancellation apart from
    /// <paramref name="timeoutToken"/> firing - never passed to <c>Dns.GetHostAddressesAsync</c>
    /// directly, since that would leave resolution unbounded by <see cref="RemoteImageOptions.Timeout"/>.
    /// </param>
    /// <param name="timeoutToken">The linked, timeout-bearing token that actually bounds the lookup.</param>
    internal static async Task<bool> IsBlockedHostAsync(
        string host, CancellationToken cancellationToken, CancellationToken timeoutToken)
    {
        // A literal address needs no lookup - and must not get one, since resolving it would be a
        // second chance to be told something different.
        if (IPAddress.TryParse(host, out var literal)) return IsBlockedAddress(literal);

        try
        {
            var addresses = await Dns.GetHostAddressesAsync(host, timeoutToken).ConfigureAwait(false);
            return IsBlockedAddresses(addresses);
        }
        catch (Exception) when (!cancellationToken.IsCancellationRequested)
        {
            // Cannot resolve it, cannot vouch for it - whether that is a genuine resolution
            // failure or Timeout firing during the lookup. Either way this is "skip the image",
            // not "the caller cancelled": that case does not match this filter, so it propagates
            // out uncaught instead of being reported as an unresolvable host.
            return true;
        }
    }
}
