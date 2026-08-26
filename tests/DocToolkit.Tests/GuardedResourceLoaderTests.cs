using System.Diagnostics;
using System.Net;
using System.Net.Sockets;
using System.Text;
using Xunit;
using Xunit.Abstractions;

namespace DocToolkit.Tests;

public class GuardedResourceLoaderTests
{
    private readonly ITestOutputHelper _output;

    public GuardedResourceLoaderTests(ITestOutputHelper output) => _output = output;

    [Theory]
    [InlineData("127.0.0.1")]        // loopback
    [InlineData("::1")]              // loopback, v6
    [InlineData("::")]               // unspecified - Socket.ConnectAsync refuses it anyway, but
                                     // defence in depth should not depend on that
    [InlineData("10.0.0.5")]         // RFC1918
    [InlineData("10.0.0.0")]         // RFC1918, exact low edge - a narrowed range must fail here
    [InlineData("10.255.255.255")]   // RFC1918, exact high edge
    [InlineData("172.16.0.1")]       // RFC1918
    [InlineData("172.16.0.0")]       // RFC1918, exact low edge
    [InlineData("172.31.255.255")]   // RFC1918, exact high edge
    [InlineData("192.168.1.1")]      // RFC1918
    [InlineData("192.168.0.0")]      // RFC1918, exact low edge
    [InlineData("192.168.255.255")]  // RFC1918, exact high edge
    [InlineData("169.254.169.254")]  // link-local - the cloud metadata endpoint
    [InlineData("100.64.0.1")]       // CGNAT (RFC 6598)
    [InlineData("100.64.0.0")]       // CGNAT, exact low edge
    [InlineData("100.127.255.255")]  // CGNAT, exact high edge
    [InlineData("224.0.0.0")]        // multicast, exact low edge - catches ">= 224" narrowed to ">"
    [InlineData("fe80::1")]          // link-local, v6
    [InlineData("fc00::1")]          // unique-local, v6
    [InlineData("::ffff:10.0.0.5")]  // v4-mapped RFC1918 - the obvious bypass
    // NAT64 (RFC 6052), 6to4 (RFC 3056), IPv4-translated and IPv4-compatible: every one of these
    // wraps a real IPv4 address in an IPv6 prefix that IsIPv4MappedToIPv6 does not recognise, so
    // an address check that only unwraps ::ffff:0:0/96 waves all of these through. On an IPv6-only
    // host behind DNS64/NAT64 (AWS IPv6-only subnets, many Kubernetes clusters, mobile carriers), a
    // hostname resolving to 64:ff9b::a9fe:a9fe reaches the cloud metadata endpoint unless these are
    // unwrapped too.
    [InlineData("64:ff9b::a9fe:a9fe")]   // NAT64 carrying 169.254.169.254
    [InlineData("64:ff9b::7f00:1")]      // NAT64 carrying 127.0.0.1
    [InlineData("2002:a00:5::")]         // 6to4 carrying 10.0.0.5
    [InlineData("::10.0.0.5")]           // IPv4-compatible carrying 10.0.0.5
    [InlineData("::ffff:0:10.0.0.5")]    // IPv4-translated carrying 10.0.0.5
    // RFC 8215's local-use NAT64 prefix (64:ff9b:1::/48) and Teredo (2001::/32, RFC 4380) are two
    // more IPv6 transition mechanisms that carry a real IPv4 address, missed until now. An
    // IPv6-only host whose NAT64 deployment uses the local-use prefix instead of the well-known
    // 64:ff9b::/96 one, or a Windows host with Teredo enabled, is reachable through either of
    // these unless the embedded address is unwrapped the same way as the others above.
    [InlineData("64:ff9b:1:a9fe:a9:fe00::")]   // RFC 8215 local-use NAT64 carrying 169.254.169.254
    [InlineData("2001::5601:5601")]            // Teredo carrying 169.254.169.254, obfuscated
    public void IsBlockedAddress_RefusesEveryPrivateForm(string address)
    {
        Assert.True(GuardedResourceLoader.IsBlockedAddress(IPAddress.Parse(address)));
    }

    [Theory]
    [InlineData("93.184.216.34")]
    [InlineData("2606:2800:220:1:248:1893:25c8:1946")]
    // Boundary data for every range above: each of these sits one address outside a blocked
    // range. Without them, widening 172.16-31 to all of 172, or "b[0] >= 224" to "b[0] >= 192",
    // leaves every theory in this file green - these are what actually discriminate the edges.
    [InlineData("172.15.0.1")]        // just below RFC1918 172.16.0.0/12
    [InlineData("172.32.0.1")]        // just above RFC1918 172.16.0.0/12
    [InlineData("11.0.0.1")]          // just above RFC1918 10.0.0.0/8
    [InlineData("192.167.0.1")]       // just below RFC1918 192.168.0.0/16
    [InlineData("223.255.255.255")]   // just below the multicast/reserved b[0] >= 224 cutoff
    [InlineData("100.63.255.255")]    // just below CGNAT 100.64.0.0/10
    [InlineData("100.128.0.0")]       // just above CGNAT 100.64.0.0/10
    [InlineData("9.255.255.255")]     // just below RFC1918 10.0.0.0/8
    [InlineData("169.253.0.1")]       // just below link-local 169.254.0.0/16
    [InlineData("169.255.0.1")]       // just above link-local 169.254.0.0/16
    // Boundary data for the embedded-IPv4 extraction itself: each of these carries a *public*
    // address inside an IPv6 transition prefix. Without them, collapsing TryGetEmbeddedIPv4's
    // per-prefix extraction into "block all of 2002::/16" or "block all of ::/96" - discarding the
    // embedded address and blocking the whole prefix instead - leaves every theory above green,
    // because none of them assert that a *public* embedded address is let through.
    [InlineData("2002:5db8:d822::")]         // 6to4 carrying the public 93.184.216.34
    [InlineData("64:ff9b::5db8:d822")]       // NAT64 carrying the public 93.184.216.34
    [InlineData("::93.184.216.34")]          // IPv4-compatible carrying the public 93.184.216.34
    public void IsBlockedAddress_AllowsPublicAddresses(string address)
    {
        Assert.False(GuardedResourceLoader.IsBlockedAddress(IPAddress.Parse(address)));
    }

    [Theory]
    [InlineData("file")]
    [InlineData("ftp")]
    [InlineData("data")]
    [InlineData("gopher")]
    public void SupportsProtocol_RefusesEverythingButHttpAndHttps(string protocol)
    {
        var loader = new GuardedResourceLoader(new RemoteImageOptions());

        Assert.False(loader.SupportsProtocol(protocol));
    }

    [Theory]
    [InlineData("http")]
    [InlineData("https")]
    public void SupportsProtocol_AcceptsHttpAndHttps(string protocol)
    {
        var loader = new GuardedResourceLoader(new RemoteImageOptions());

        Assert.True(loader.SupportsProtocol(protocol));
    }

    // =========================================================================================
    // IsBlockedAddresses - the "any blocked address blocks the whole host" rule, pulled out as a
    // pure function specifically so it can be tested with a crafted address list. A live DNS
    // response that mixes a public and a private A record is not something to depend on existing
    // anywhere, let alone offline.
    // =========================================================================================

    [Fact]
    public void IsBlockedAddresses_BlocksWhenAnyAddressIsBlocked_PrivateFirst()
    {
        var addresses = new[] { IPAddress.Parse("10.0.0.5"), IPAddress.Parse("93.184.216.34") };

        Assert.True(GuardedResourceLoader.IsBlockedAddresses(addresses));
    }

    [Fact]
    public void IsBlockedAddresses_BlocksWhenAnyAddressIsBlocked_PublicFirst()
    {
        // The public record must not be enough just because it happened to be checked first.
        var addresses = new[] { IPAddress.Parse("93.184.216.34"), IPAddress.Parse("10.0.0.5") };

        Assert.True(GuardedResourceLoader.IsBlockedAddresses(addresses));
    }

    [Fact]
    public void IsBlockedAddresses_AllowsWhenEveryAddressIsPublic()
    {
        var addresses = new[] { IPAddress.Parse("93.184.216.34"), IPAddress.Parse("1.1.1.1") };

        Assert.False(GuardedResourceLoader.IsBlockedAddresses(addresses));
    }

    [Fact]
    public void IsBlockedAddresses_BlocksAnEmptyList()
    {
        // No A record at all is exactly as unusable as an unresolvable host.
        Assert.True(GuardedResourceLoader.IsBlockedAddresses(Array.Empty<IPAddress>()));
    }

    // =========================================================================================
    // FetchAsync - pinned here in isolation, against the loader directly. RemoteImageGuardTests
    // covers the same guard end to end through HtmlToDocxConverter; these tests exist alongside it
    // because a converter-level assertion cannot tell "FetchAsync refused it" apart from "the
    // conversion never asked", and several invariants below (the exact byte-cap boundary, caller
    // cancellation during the DNS phase) have no observable end-to-end signal at all. Each test
    // pins one invariant against a real loopback socket rather than a mock, reusing LoopbackProbe
    // with a custom responder rather than a second raw-socket HTTP server.
    //
    // One property is deliberately NOT pinned here: that DNS resolution itself happens inside
    // Timeout (the linked token is created and CancelAfter'd before IsBlockedHostAsync is called,
    // specifically so a black-holed nameserver cannot stall past Timeout - see FetchAsync's own
    // comment). Proving that against a loopback fixture would need a fake DNS resolver that can be
    // told to hang for longer than Timeout, and this loader calls the real System.Net.Dns
    // statically with no seam to substitute one - there is no way to control real resolution
    // latency deterministically from a test, and a "real slow resolver" test would be exactly the
    // flaky, environment-dependent kind of test this file exists to avoid (see the finding fixed
    // above about FetchAsync_UnresolvableHost_IsBlocked). This is a structural gap, left
    // documented rather than faked: if DNS resolution is ever moved outside the linked token, the
    // suite stays green.
    // =========================================================================================

    [Fact]
    public async Task FetchAsync_AllowlistDoesNotWaiveTheAddressCheck()
    {
        using var probe = new LoopbackProbe(_output);
        var options = new RemoteImageOptions();
        options.AllowedHosts.Add("127.0.0.1");   // explicitly allowlisted...
        // ...but AllowPrivateAddresses stays false, the default. The allowlist narrows which
        // hosts may be fetched; it must never widen what counts as a safe address.
        var loader = new GuardedResourceLoader(options);

        var result = await loader.FetchAsync(new Uri($"{probe.BaseUrl}/x.png"));

        Assert.Null(result);
        Assert.False(await probe.WaitForConnectionAsync(TimeSpan.FromMilliseconds(500)),
            "AllowedHosts must not waive the address check: a name with a public and a private " +
            "A record, or an operator who allowlists a literal private address by mistake, must " +
            "still be blocked before a socket opens - this is the DNS-rebinding hole otherwise.");
    }

    [Fact]
    public async Task FetchAsync_ByteCap_RejectsOneByteOverTheLimit_WithNoContentLengthToConsult()
    {
        const int cap = 16;

        // No Content-Length header at all - the body is delimited purely by the connection
        // closing (HTTP/1.1 RFC 7230 §3.3.3 case 7). If the cap were driven by a declared length
        // instead of the bytes actually read, there would be nothing here to drive it with, and
        // the whole cap+1 bytes would go through.
        using var probe = new LoopbackProbe(_output, FixedBodyResponder(cap + 1));
        var loader = new GuardedResourceLoader(
            new RemoteImageOptions { AllowPrivateAddresses = true, MaxBytesPerImage = cap });

        var result = await loader.FetchAsync(new Uri($"{probe.BaseUrl}/x.bin"));

        Assert.Null(result);
    }

    [Fact]
    public async Task FetchAsync_ByteCap_AcceptsExactlyTheLimit()
    {
        const int cap = 16;
        using var probe = new LoopbackProbe(_output, FixedBodyResponder(cap));
        var loader = new GuardedResourceLoader(
            new RemoteImageOptions { AllowPrivateAddresses = true, MaxBytesPerImage = cap });

        var result = await loader.FetchAsync(new Uri($"{probe.BaseUrl}/x.bin"));

        Assert.NotNull(result);
        Assert.Equal(cap, result!.Content.Length);
    }

    // Chunked transfer encoding is the third framing HTTP/1.1 offers, and the one a streaming
    // attacker actually reaches for: the response declares no length at all and the server keeps
    // emitting chunks for as long as the client keeps reading. The two tests above cover the other
    // two framings (declared Content-Length, and close-delimited), so without these the cap was
    // reasoned about on this path rather than measured - ReadCappedAsync checks inside the read
    // loop and HttpClient dechunks above the stream it reads, so chunked bodies traverse the
    // identical loop, but "traverses the identical loop" is an argument, not a test.
    //
    // Note what the accepting case asserts: exactly `cap` bytes back from a body whose *wire*
    // form is larger (chunk-size lines and CRLFs between every chunk). A cap accidentally counting
    // wire bytes instead of decoded ones would reject that body, so the pair also pins which of
    // the two is being measured.

    [Fact]
    public async Task FetchAsync_ByteCap_RejectsAChunkedBodyThatCrossesTheLimit()
    {
        const int cap = 16;

        // Four 8-byte chunks: 32 bytes total, but no single chunk is over the cap and no single
        // read is either. Only the running total crosses it, which is what a cap checked per-read
        // rather than per-total would miss entirely.
        using var probe = new LoopbackProbe(_output, ChunkedBodyResponder(chunkSize: 8, chunkCount: 4));
        var loader = new GuardedResourceLoader(
            new RemoteImageOptions { AllowPrivateAddresses = true, MaxBytesPerImage = cap });

        var result = await loader.FetchAsync(new Uri($"{probe.BaseUrl}/x.bin"));

        Assert.Null(result);
    }

    [Fact]
    public async Task FetchAsync_ByteCap_AcceptsAChunkedBodyWithinTheLimit()
    {
        const int cap = 16;

        // The counterweight: two 8-byte chunks land exactly on the cap. Without this, a guard that
        // simply refused every chunked response would pass the rejection test above.
        using var probe = new LoopbackProbe(_output, ChunkedBodyResponder(chunkSize: 8, chunkCount: 2));
        var loader = new GuardedResourceLoader(
            new RemoteImageOptions { AllowPrivateAddresses = true, MaxBytesPerImage = cap });

        var result = await loader.FetchAsync(new Uri($"{probe.BaseUrl}/x.bin"));

        Assert.NotNull(result);
        Assert.Equal(cap, result!.Content.Length);
    }

    private const int TimingAttempts = 3;

    /// <summary>
    /// <b>Best-of-N, and this is the THIRD member of a family that has now needed it three times.</b>
    ///
    /// This was a single-shot wall-clock assertion and it failed CI on 2026-08-26 at <b>7.63 s</b>
    /// against a 5 s ceiling — on a pull request that changed <c>DocxEditor.BlockText</c> and
    /// nothing this test can reach. The identical check <i>passed</i> on a parallel run of the same
    /// commit, which is the tell.
    ///
    /// Its two siblings — <c>AirGapGuardTests.MeasureStallAsync</c> and
    /// <c>RemoteImageGuardTests.SlowResponse_TimesOut_RatherThanHangingTheConversion</c> — were
    /// each given this defence after flaking the same way. This one was simply missed, and
    /// <c>CLAUDE.md</c> already carried the reasoning that applies to it verbatim.
    ///
    /// <b>Taking the SMALLEST elapsed is valid because the noise is one-sided.</b> CPU contention on
    /// a 2-4 core runner can only make an attempt slower, never faster — xunit runs collections in
    /// parallel and this suite has PDF renders beside it. The defect being detected costs
    /// <see cref="SlowDripResponder"/>'s full drip on <i>every</i> attempt, so it survives the
    /// minimum. That is measured from the responder rather than assumed: it writes 50 chunks
    /// 200 ms apart, so an unbounded read takes <b>~10 s</b> every time against a correct read's
    /// ~0.3 s. The ceiling sits between them at 5 s.
    ///
    /// <b>Do NOT "fix" a future failure here by raising the ceiling.</b> 5 s is already 16x the
    /// 300 ms <c>Timeout</c> being asserted, and raising it walks toward the ~10 s that IS the
    /// defect. The headroom is not the problem.
    ///
    /// <b>Verified by sabotage — and the OBVIOUS sabotage does not test the timing half at all.</b>
    /// Removing the bound entirely (<c>CancelAfter(InfiniteTimeSpan)</c>) lets the drip finish, so
    /// the fetch <i>succeeds</i> and <c>Assert.Null</c> fires — the test goes red for a reason that
    /// has nothing to do with elapsed time, which is the "green for the wrong reason" failure this
    /// repository keeps recording, wearing red instead.
    ///
    /// The sabotage that actually exercises this loop keeps the bound and makes it <i>late</i>:
    /// <c>CancelAfter(TimeSpan.FromSeconds(8))</c>. The fetch is still refused, so
    /// <c>Assert.Null</c> passes, and only the timing assertion can see the defect. Measured — all
    /// three attempts came in at 8.00-8.01 s, so it survives the minimum exactly as the one-sided
    /// argument requires:
    ///
    /// <code>
    /// attempt 1/3: 8.01 s (best 8.01 s, ceiling 5 s)
    /// attempt 2/3: 8.01 s (best 8.01 s, ceiling 5 s)
    /// attempt 3/3: 8.00 s (best 8.00 s, ceiling 5 s)
    /// </code>
    /// </summary>
    [Fact]
    public async Task FetchAsync_Timeout_BoundsASlowDripBody()
    {
        var ceiling = TimeSpan.FromSeconds(5);
        var best = TimeSpan.MaxValue;

        for (var attempt = 1; attempt <= TimingAttempts; attempt++)
        {
            // A fresh probe per attempt rather than reusing one: nothing here promises the
            // responder serves a second request, and a silently-refused reconnect would look
            // exactly like a fast read.
            using var probe = new LoopbackProbe(_output, SlowDripResponder);
            var loader = new GuardedResourceLoader(new RemoteImageOptions
            {
                AllowPrivateAddresses = true,
                Timeout = TimeSpan.FromMilliseconds(300),
            });

            var stopwatch = Stopwatch.StartNew();
            var result = await loader.FetchAsync(new Uri($"{probe.BaseUrl}/slow.bin"));
            stopwatch.Stop();

            // Asserted on EVERY attempt, not only the fastest. This is the actual guarantee - the
            // fetch is refused - and it is not a timing claim, so a retry must never launder it.
            Assert.Null(result);

            if (stopwatch.Elapsed < best) best = stopwatch.Elapsed;
            _output.WriteLine(
                $"attempt {attempt}/{TimingAttempts}: {stopwatch.Elapsed.TotalSeconds:0.00} s "
                + $"(best {best.TotalSeconds:0.00} s, ceiling {ceiling.TotalSeconds:0.#} s)");

            // A first attempt under the ceiling has answered the question; retrying would triple
            // the cost of the common case for nothing.
            if (best < ceiling) break;
        }

        Assert.True(best < ceiling,
            $"Best of {TimingAttempts} attempts was {best.TotalSeconds:0.00} s against a slow-drip "
            + "body with a 300 ms Timeout - Timeout is not bounding the read.");
    }

    [Fact]
    public async Task FetchAsync_CallerCancellation_PropagatesDuringTheHttpPhase()
    {
        using var probe = new LoopbackProbe(_output, SlowDripResponder);
        var loader = new GuardedResourceLoader(new RemoteImageOptions
        {
            AllowPrivateAddresses = true,
            Timeout = TimeSpan.FromSeconds(30),   // long enough that Timeout cannot be what fires
        });
        using var cts = new CancellationTokenSource(TimeSpan.FromMilliseconds(200));

        await Assert.ThrowsAnyAsync<OperationCanceledException>(
            () => loader.FetchAsync(new Uri($"{probe.BaseUrl}/slow.bin"), cts.Token));
    }

    [Fact]
    public async Task FetchAsync_CallerCancellation_PropagatesDuringTheDnsPhase()
    {
        // Regression test for the finding: a pre-cancelled token against a non-literal host takes
        // the DNS path in IsBlockedHostAsync, whose bare catch used to swallow
        // OperationCanceledException and report the host as merely unresolvable - a *completed*
        // null result instead of the cancellation propagating, which would leave a cancelled
        // conversion looking like a successful one with images silently missing.
        var loader = new GuardedResourceLoader(new RemoteImageOptions());
        using var cts = new CancellationTokenSource();
        cts.Cancel();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(
            () => loader.FetchAsync(new Uri("http://example.org/x.png"), cts.Token));
    }

    [Fact]
    public async Task IsBlockedHostAsync_UnresolvableHost_ReturnsTrue()
    {
        // Asserted directly against IsBlockedHostAsync, not through FetchAsync. The previous
        // version of this test called FetchAsync and asserted only Assert.Null(result) - which is
        // satisfied identically by "blocked before connecting" and by "resolution returned
        // something, SendAsync then failed for an unrelated reason, and FetchAsync's outer catch
        // returned null anyway". Measured: flipping IsBlockedHostAsync's catch from fail-closed
        // (return true) to fail-open (return false) and rebuilding left all 51 tests in the suite
        // passing, including this one - the invariant was invisible end-to-end. Asserting the
        // return value of IsBlockedHostAsync itself closes that gap, and as a side effect this
        // test can never make a genuine outbound HTTP request: the function under test only
        // resolves DNS, it never opens a socket to whatever address comes back.
        //
        // ".invalid" is reserved by RFC 2606 specifically to never resolve.
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(5));

        var stopwatch = Stopwatch.StartNew();
        var result = await GuardedResourceLoader.IsBlockedHostAsync(
            "definitely-does-not-exist.invalid", CancellationToken.None, timeout.Token);
        stopwatch.Stop();

        Assert.True(result);
        _output.WriteLine($"Gave up on an unresolvable host in {stopwatch.Elapsed.TotalSeconds:0.00} s.");
    }

    /// <summary>
    /// Writes a response with no <c>Content-Length</c> header, exactly <paramref name="length"/>
    /// bytes of body, then closes the connection - the framing HTTP/1.1 uses when the length is
    /// not declared up front.
    /// </summary>
    private static Func<NetworkStream, string, CancellationToken, Task> FixedBodyResponder(int length) =>
        async (stream, _, ct) =>
        {
            await stream.WriteAsync(Encoding.ASCII.GetBytes(
                "HTTP/1.1 200 OK\r\nContent-Type: application/octet-stream\r\nConnection: close\r\n\r\n"),
                ct);
            await stream.WriteAsync(new byte[length], ct);
            await stream.FlushAsync(ct);
        };

    /// <summary>
    /// Writes a <c>Transfer-Encoding: chunked</c> response of <paramref name="chunkCount"/> chunks
    /// of <paramref name="chunkSize"/> bytes each, terminated by the zero-length chunk. Written by
    /// hand at the socket rather than through any HTTP server library, for the same reason
    /// <see cref="LoopbackProbe"/> exists at all: a framing bug in a server abstraction would
    /// otherwise be indistinguishable from the behaviour under test.
    /// </summary>
    private static Func<NetworkStream, string, CancellationToken, Task> ChunkedBodyResponder(
        int chunkSize, int chunkCount) =>
        async (stream, _, ct) =>
        {
            // No Content-Length: chunked framing supplies the length per chunk instead, so there
            // is nothing here a length-driven cap could consult even if it wanted to.
            await stream.WriteAsync(Encoding.ASCII.GetBytes(
                "HTTP/1.1 200 OK\r\nContent-Type: application/octet-stream\r\n" +
                "Transfer-Encoding: chunked\r\nConnection: close\r\n\r\n"), ct);

            for (var i = 0; i < chunkCount; i++)
            {
                // Chunk size is hex, per RFC 7230 section 4.1 - "10" here would mean 16 bytes.
                await stream.WriteAsync(Encoding.ASCII.GetBytes($"{chunkSize:x}\r\n"), ct);
                await stream.WriteAsync(new byte[chunkSize], ct);
                await stream.WriteAsync(Encoding.ASCII.GetBytes("\r\n"), ct);
                await stream.FlushAsync(ct);
            }

            await stream.WriteAsync(Encoding.ASCII.GetBytes("0\r\n\r\n"), ct);
            await stream.FlushAsync(ct);
        };

    /// <summary>
    /// Sends headers immediately, then one byte at a time with a delay between each - long enough
    /// in total (up to 10 s) that a correctly-bounded <c>Timeout</c> or a caller's own
    /// cancellation cuts the read off well before the drip finishes, but capped so a test failure
    /// here cannot hang the run if cancellation is not observed for some reason.
    /// </summary>
    private static async Task SlowDripResponder(NetworkStream stream, string requestLine, CancellationToken ct)
    {
        await stream.WriteAsync(Encoding.ASCII.GetBytes(
            "HTTP/1.1 200 OK\r\nContent-Type: application/octet-stream\r\nConnection: close\r\n\r\n"), ct);

        using var bound = CancellationTokenSource.CreateLinkedTokenSource(ct);
        bound.CancelAfter(TimeSpan.FromSeconds(10));
        try
        {
            for (var i = 0; i < 50; i++)
            {
                await Task.Delay(TimeSpan.FromMilliseconds(200), bound.Token);
                await stream.WriteAsync(new byte[] { 0 }, bound.Token);
                await stream.FlushAsync(bound.Token);
            }
        }
        catch (OperationCanceledException)
        {
            // The client gave up - Timeout, its own cancellation, or the probe tearing down.
            // Either way there is nothing left to serve.
        }
    }
}
