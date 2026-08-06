using System.Diagnostics;
using System.Net.Sockets;
using System.Text;
using Xunit;
using Xunit.Abstractions;

namespace DocToolkit.Tests;

/// <summary>
/// Proves the remote-image opt-in is bounded <b>end to end, over a real socket</b> - not just at
/// the address-classifier level, which <see cref="GuardedResourceLoaderTests"/> already covers in
/// isolation. Every test here asserts on <see cref="LoopbackProbe"/>'s connection count or on the
/// converted document's own content, never on a mock, for the same reason
/// <c>AirGapGuardTests</c> counts sockets rather than mocking an HTTP client: the assertion then
/// holds regardless of which internal component does the fetching.
///
/// A zero-connection assertion on its own proves nothing - it passes identically whether the guard
/// works or the probe is broken - so <see cref="Loopback_IsFetched_WithAllowPrivateAddresses"/> is
/// the load-bearing counterweight to
/// <see cref="Loopback_IsRefusedByDefault"/>: together they show the block is real. See this
/// task's report for the mutation proof that confirms it.
/// </summary>
public class RemoteImageGuardTests
{
    private readonly ITestOutputHelper _output;

    public RemoteImageGuardTests(ITestOutputHelper output) => _output = output;

    // =========================================================================================
    // 1. file:// - the sharpest fix in this change. HtmlToOpenXml's own DefaultWebRequest speaks
    // file:// as well as http/https, which would let an <img src="file:///etc/passwd"> read the
    // host's own disk into the document. GuardedResourceLoader refuses it at the protocol gate
    // (SupportsProtocol only accepts http/https), before any address check or fetch is attempted,
    // so this holds even with downloads fully enabled via RemoteImageOptions.
    // =========================================================================================

    [Fact]
    public async Task Fetch_RefusesFileProtocol_EvenWithDownloadsEnabled()
    {
        var html = """
            <p>
              <img src="file:///etc/passwd" alt="unix path" />
              <img src="file:///C:/Windows/win.ini" alt="windows path" />
            </p>
            """;

        byte[] docx = Array.Empty<byte>();
        await RemoteDownloadGate.RunAsync(async () =>
            docx = await HtmlToDocxConverter.ConvertAsync(html, new RemoteImageOptions()));

        Assert.NotEmpty(docx);
        Assert.Equal(0, DocxFixtures.Read(docx, main => main.ImageParts.Count()));
    }

    // =========================================================================================
    // 2 & 3. Loopback blocked by default, fetched once explicitly allowed. These two are a pair:
    // (2) alone cannot tell "the guard refused it" apart from "the probe never saw anything for
    // some unrelated reason", and (3) is what rules the second explanation out. See this task's
    // report for the direct mutation proof against IsBlockedAddress.
    // =========================================================================================

    [Fact]
    public async Task Loopback_IsRefusedByDefault()
    {
        using var probe = new LoopbackProbe(_output);

        await RemoteDownloadGate.RunAsync(() => HtmlToDocxConverter.ConvertAsync(
            $"""<p><img src="{probe.BaseUrl}/logo.bmp" alt="logo" /></p>""",
            new RemoteImageOptions()));

        await probe.AssertSilentAsync("RemoteImageGuardTests.Loopback_IsRefusedByDefault");
    }

    [Fact]
    public async Task Loopback_IsFetched_WithAllowPrivateAddresses()
    {
        using var probe = new LoopbackProbe(_output);

        await RemoteDownloadGate.RunAsync(() => HtmlToDocxConverter.ConvertAsync(
            $"""<p><img src="{probe.BaseUrl}/logo.bmp" alt="logo" /></p>""",
            new RemoteImageOptions { AllowPrivateAddresses = true }));

        Assert.True(await probe.WaitForConnectionAsync(TimeSpan.FromSeconds(10)),
            "AllowPrivateAddresses = true made no outbound connection to the loopback probe - " +
            "either the guard is refusing more than it should, or the probe cannot detect a " +
            "fetch, which would make the zero-connection assertion above vacuous.");
    }

    // =========================================================================================
    // 4. AllowedHosts narrows which hosts may be fetched; it must never widen what counts as a
    // safe address (that invariant is already pinned directly against IsBlockedAddresses in
    // GuardedResourceLoaderTests). Both probes bind to 127.0.0.1, so AllowPrivateAddresses = true
    // is set in both phases purely to get past the address check - what actually discriminates
    // "outside" from "inside" here is membership in AllowedHosts, not the address.
    // =========================================================================================

    [Fact]
    public async Task AllowedHosts_RefusesOutsideHost_FetchesInsideHost()
    {
        using (var outside = new LoopbackProbe(_output))
        {
            var options = new RemoteImageOptions { AllowPrivateAddresses = true };
            options.AllowedHosts.Add("images.example.invalid");   // never matches 127.0.0.1

            await RemoteDownloadGate.RunAsync(() => HtmlToDocxConverter.ConvertAsync(
                $"""<p><img src="{outside.BaseUrl}/logo.bmp" alt="logo" /></p>""", options));

            await outside.AssertSilentAsync(
                "RemoteImageGuardTests.AllowedHosts (host outside the allow-list)");
        }

        using (var inside = new LoopbackProbe(_output))
        {
            var options = new RemoteImageOptions { AllowPrivateAddresses = true };
            options.AllowedHosts.Add("127.0.0.1");

            await RemoteDownloadGate.RunAsync(() => HtmlToDocxConverter.ConvertAsync(
                $"""<p><img src="{inside.BaseUrl}/logo.bmp" alt="logo" /></p>""", options));

            Assert.True(await inside.WaitForConnectionAsync(TimeSpan.FromSeconds(10)),
                "A host inside a non-empty AllowedHosts made no outbound connection.");
        }
    }

    // =========================================================================================
    // 5. The byte cap must be enforced by counting bytes actually read off the stream, never by
    // trusting Content-Length - GuardedResourceLoader's ReadCappedAsync never even looks at that
    // header. Measured directly against a real HttpClient/SocketsHttpHandler pair: when a
    // Content-Length header is present, HttpClient's own HTTP/1.1 framing delivers *exactly* that
    // many bytes to the caller and silently discards anything the server writes afterwards, no
    // matter how much more there is - so a header that understates a small body can never be used
    // to smuggle extra bytes past a byte-counting reader. The header here therefore declares a
    // value that is itself already past the cap (4 KB against a 1 KB cap) while still understating
    // the far larger amount the server actually attempts to write (64 KB) - "understating" the
    // true size, exactly as a lying server would, while still giving the cap something real to
    // trip over as GuardedResourceLoader reads the framed body incrementally.
    // =========================================================================================

    [Fact]
    public async Task OversizedResponse_IsAborted_DespiteAnUnderstatedContentLength()
    {
        const int cap = 1024;
        const int declaredLength = 4096;
        const int actualLength = 65536;

        using var probe = new LoopbackProbe(
            _output, LyingContentLengthResponder(declaredLength, actualLength));
        var options = new RemoteImageOptions { AllowPrivateAddresses = true, MaxBytesPerImage = cap };

        byte[] docx = Array.Empty<byte>();
        await RemoteDownloadGate.RunAsync(async () =>
            docx = await HtmlToDocxConverter.ConvertAsync(
                $"""<p><img src="{probe.BaseUrl}/big.bin" alt="big" /></p>""", options));

        Assert.NotEmpty(docx);
        Assert.Equal(0, DocxFixtures.Read(docx, main => main.ImageParts.Count()));
    }

    // =========================================================================================
    // 6. A slow response must cost at most Timeout, not hang the conversion. The drip is slow
    // enough (up to 10 s of total drip time) that a Timeout that is not actually bounding the read
    // would make this test itself run long enough to notice, well inside the suite's own default
    // per-test timeout.
    // =========================================================================================

    [Fact]
    public async Task SlowResponse_TimesOut_RatherThanHangingTheConversion()
    {
        using var probe = new LoopbackProbe(_output, SlowDripResponder);
        var options = new RemoteImageOptions
        {
            AllowPrivateAddresses = true,
            Timeout = TimeSpan.FromMilliseconds(300),
        };

        var stopwatch = Stopwatch.StartNew();
        byte[] docx = Array.Empty<byte>();
        await RemoteDownloadGate.RunAsync(async () =>
            docx = await HtmlToDocxConverter.ConvertAsync(
                $"""<p><img src="{probe.BaseUrl}/slow.bin" alt="slow" /></p>""", options));
        stopwatch.Stop();

        Assert.NotEmpty(docx);
        Assert.True(stopwatch.Elapsed < TimeSpan.FromSeconds(5),
            $"Took {stopwatch.Elapsed.TotalSeconds:0.00} s against a slow-drip body with a " +
            "300 ms Timeout - Timeout is not bounding the conversion.");
    }

    // =========================================================================================
    // 7. Options validation must reach the caller through the converter, not just through
    // RemoteImageOptions.Validate() called directly (already pinned in RemoteImageOptionsTests).
    // =========================================================================================

    [Fact]
    public async Task ConvertAsync_NullOptions_ThrowsArgumentNullException()
    {
        await Assert.ThrowsAsync<ArgumentNullException>(
            () => HtmlToDocxConverter.ConvertAsync("<p>x</p>", (RemoteImageOptions)null!));
    }

    [Fact]
    public async Task ConvertAsync_ZeroTimeout_ThrowsArgumentOutOfRangeException()
    {
        var options = new RemoteImageOptions { Timeout = TimeSpan.Zero };

        await Assert.ThrowsAsync<ArgumentOutOfRangeException>(
            () => HtmlToDocxConverter.ConvertAsync("<p>x</p>", options));
    }

    // =========================================================================================
    // Responders. LoopbackProbe (see AirGapGuardTests.cs) already accepts an arbitrary responder
    // delegate via its constructor, so no change to that shared class is needed for either of
    // these - they are just two more delegates, the same shape GuardedResourceLoaderTests already
    // supplies its own of.
    // =========================================================================================

    /// <summary>
    /// Declares <paramref name="declaredLength"/> in <c>Content-Length</c> - understating the
    /// truth - then writes <paramref name="actualLength"/> raw bytes regardless. A compliant
    /// HTTP/1.1 client reads exactly the declared count and discards the rest, which is the
    /// behaviour this test's own comment measures and relies on.
    /// </summary>
    private static Func<NetworkStream, string, CancellationToken, Task> LyingContentLengthResponder(
        int declaredLength, int actualLength) =>
        async (stream, _, ct) =>
        {
            await stream.WriteAsync(Encoding.ASCII.GetBytes(
                "HTTP/1.1 200 OK\r\nContent-Type: application/octet-stream\r\n" +
                $"Content-Length: {declaredLength}\r\nConnection: close\r\n\r\n"), ct);
            await stream.WriteAsync(new byte[actualLength], ct);
            await stream.FlushAsync(ct);
        };

    /// <summary>
    /// Sends headers immediately, then one byte at a time with a delay between each - long enough
    /// in total (up to 10 s) that a correctly-bounded <c>Timeout</c> cuts the read off well before
    /// the drip finishes, but capped so a failure to observe the timeout cannot hang the run
    /// indefinitely.
    /// </summary>
    private static async Task SlowDripResponder(NetworkStream stream, string requestLine, CancellationToken ct)
    {
        await stream.WriteAsync(Encoding.ASCII.GetBytes(
            "HTTP/1.1 200 OK\r\nContent-Type: image/bmp\r\nConnection: close\r\n\r\n"), ct);

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
            // The client gave up - Timeout, or the probe tearing down. Nothing left to serve.
        }
    }
}
