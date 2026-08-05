namespace DocToolkit;

/// <summary>
/// Bounds the one code path in this library that opens a socket: fetching images named by absolute
/// URLs during HTML conversion. Every default here is the restrictive one, so
/// <c>new RemoteImageOptions()</c> is safe to pass without reading this class.
///
/// <b>This is not a complete SSRF defence.</b> Host addresses are resolved and checked, then
/// resolved again by the HTTP stack when it connects — a DNS entry that changes between those two
/// moments defeats the check. It stops the ordinary cases, a literal metadata address or a
/// hard-coded internal hostname, and raises the cost of the rest; a service converting genuinely
/// untrusted HTML should also be egress-filtered at the network layer.
/// </summary>
public sealed class RemoteImageOptions
{
    /// <summary>How long a single image fetch may take. Default 10 seconds.</summary>
    public TimeSpan Timeout { get; set; } = TimeSpan.FromSeconds(10);

    /// <summary>
    /// The most bytes accepted from a single image. Default 5 MB. Enforced while reading the
    /// response body, not taken from <c>Content-Length</c>, which a hostile server can understate.
    /// </summary>
    public long MaxBytesPerImage { get; set; } = 5 * 1024 * 1024;

    /// <summary>
    /// Hosts that may be fetched from, compared case-insensitively. Empty — the default — means any
    /// host that also passes the address check; it does <b>not</b> disable that check.
    /// </summary>
    public ISet<string> AllowedHosts { get; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// Whether loopback, private and link-local addresses may be fetched. Default <c>false</c>,
    /// which is what keeps <c>169.254.169.254</c> — the cloud metadata endpoint — and internal
    /// services out of reach of attacker-supplied markup.
    /// </summary>
    public bool AllowPrivateAddresses { get; set; }

    /// <summary>Throws if any value is unusable. Called once, before conversion begins.</summary>
    internal void Validate()
    {
        if (Timeout <= TimeSpan.Zero)
            throw new ArgumentOutOfRangeException(
                nameof(Timeout), Timeout, "Timeout must be greater than zero.");

        if (MaxBytesPerImage <= 0)
            throw new ArgumentOutOfRangeException(
                nameof(MaxBytesPerImage), MaxBytesPerImage,
                "MaxBytesPerImage must be greater than zero.");

        if (AllowedHosts.Any(string.IsNullOrWhiteSpace))
            throw new ArgumentException("AllowedHosts contains a blank entry.", nameof(AllowedHosts));
    }
}
