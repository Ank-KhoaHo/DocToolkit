using System.Diagnostics;
using System.Net;
using System.Net.Sockets;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocToolkit;
using Xunit;

namespace DocToolkit.Tests;

public class HtmlToDocxConverterTests
{
    private const string Html = """
        <h1>Quarterly Report</h1>
        <p>Revenue was <strong>up 12%</strong> and costs were <em>flat</em>.</p>
        <table border="1"><tr><th>Region</th><th>Total</th></tr>
        <tr><td>North</td><td>1200</td></tr></table>
        <ul><li>First</li><li>Second</li></ul>
        <p><a href="https://example.com/report">Full report</a></p>
        """;

    [Fact]
    public async Task ConvertAsync_ProducesAValidDocxPackage()
    {
        var bytes = await HtmlToDocxConverter.ConvertAsync(Html);

        Assert.NotEmpty(bytes);
        // A .docx is a ZIP: it must start with the local file header magic "PK\x03\x04".
        Assert.Equal(new byte[] { 0x50, 0x4B, 0x03, 0x04 }, bytes.Take(4).ToArray());
    }

    [Fact]
    public async Task ConvertAsync_PreservesStructureAndFormatting()
    {
        var bytes = await HtmlToDocxConverter.ConvertAsync(Html);

        using var ms = new MemoryStream(bytes);
        using var doc = WordprocessingDocument.Open(ms, false);
        var body = doc.MainDocumentPart!.Document!.Body!;

        Assert.True(body.Descendants<Paragraph>().Count() >= 4);
        Assert.Single(body.Descendants<Table>());
        Assert.Equal(2, body.Descendants<TableRow>().Count());
        Assert.NotEmpty(body.Descendants<Bold>());
        Assert.NotEmpty(body.Descendants<Italic>());
        Assert.NotEmpty(body.Descendants<Hyperlink>());
        Assert.Contains("Quarterly Report", body.InnerText);
    }

    [Fact]
    public async Task ConvertAsync_RejectsNullHtml()
    {
        await Assert.ThrowsAsync<ArgumentNullException>(
            () => HtmlToDocxConverter.ConvertAsync(null!));
    }

    // ---------------------------------------------------------------------------------------
    // I-10: the CancellationToken now reaches HtmlToOpenXml's parser, so a long conversion can
    // actually be abandoned mid-flight rather than only at the entry check.
    // ---------------------------------------------------------------------------------------

    /// <summary>
    /// The token must reach HtmlToOpenXml's parser, not merely the guard on the way in. Proved by
    /// cancelling at a point that is INSIDE the parse by construction rather than by timing.
    ///
    /// The image is the first element, so the parser fetches it before converting the 60,000
    /// paragraphs that follow. The probe cancels the moment that fetch arrives. Cancellation is
    /// therefore requested while the parser demonstrably still has most of its work left, with no
    /// clock involved anywhere.
    ///
    /// WHY NOT A TIMER, which is what this test used to do: `cts.CancelAfter(150ms)` schedules its
    /// callback on the thread pool, and xunit runs test collections in parallel, so on a loaded
    /// machine the callback can fire long after the parse has finished - the call then returns a
    /// document and the assertion fails with "no exception was thrown". That is exactly how it
    /// broke: it passed 5/5 run alone and failed inside the full suite on a Windows runner,
    /// blocking a release PR. Backlog B10 removed an earlier timing assertion but left the race,
    /// and was marked done while this remained.
    ///
    /// The assertion itself is unchanged and is still the whole point. ConvertAsync checks the
    /// token once on entry - which cannot fire here, because the token is live when the call starts
    /// - and then hands it to ParseBody, with no check afterwards. So a token that never reached
    /// the parser would let this run to completion and return a document.
    /// </summary>
    [Fact]
    public async Task ConvertAsync_CancelsPartwayThroughALongParse()
    {
        using var cts = new CancellationTokenSource();
        using var probe = new LoopbackProbe(onContact: cts.Cancel);

        var sb = new StringBuilder();
        sb.Append("<img src=\"").Append(probe.Url).Append("\" />");
        for (var i = 0; i < 60_000; i++)
            sb.Append("<p>Paragraph ").Append(i).Append(" of the stress fixture.</p>");

        // AllowPrivateAddresses is the documented escape hatch for a loopback probe: the guard
        // blocks loopback by default and would otherwise refuse this listener, so the fetch - and
        // with it the cancellation - would never happen. Same reason AirGapGuardTests sets it.
        await Assert.ThrowsAnyAsync<OperationCanceledException>(
            () => HtmlToDocxConverter.ConvertAsync(
                sb.ToString(),
                new RemoteImageOptions { AllowPrivateAddresses = true },
                cts.Token));
    }

    [Fact]
    public async Task ConvertAsync_ThrowsForAnAlreadyCancelledToken()
    {
        using var cts = new CancellationTokenSource();
        await cts.CancelAsync();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(
            () => HtmlToDocxConverter.ConvertAsync(Html, cts.Token));
    }

    // ---------------------------------------------------------------------------------------
    // I-9: a byte[] in / byte[] out API must not silently make outbound HTTP requests.
    // ---------------------------------------------------------------------------------------

    [Fact]
    public async Task ConvertAsync_DoesNotFetchRemoteImagesByDefault()
    {
        using var probe = new LoopbackProbe();

        var bytes = await HtmlToDocxConverter.ConvertAsync(
            $"""<p>Report <img src="{probe.Url}" alt="logo" /></p>""");

        Assert.NotEmpty(bytes);
        Assert.False(await probe.WasContactedAsync(),
            "HtmlToDocxConverter made an outbound request for a remote <img> - that is an SSRF " +
            "reach and an unbounded hang for a caller who only handed over a string.");
    }

    [Fact]
    public async Task ConvertAsync_FetchesRemoteImagesOnlyWhenExplicitlyAllowed()
    {
        using var probe = new LoopbackProbe();

        // Gated: this is not the only opt-in test in the suite, and HtmlToOpenXml's ParseBody is
        // not proven safe to run concurrently with itself.
        await RemoteDownloadGate.RunAsync(async () =>
        {
            try
            {
                // GuardedResourceLoader blocks loopback, private and link-local addresses by
                // default (RemoteImageOptions.AllowPrivateAddresses is false), which would refuse
                // this listener too. AllowPrivateAddresses = true is the escape hatch this test
                // depends on to reach it.
                await HtmlToDocxConverter.ConvertAsync(
                    $"""<p>Report <img src="{probe.Url}" alt="logo" /></p>""",
                    new RemoteImageOptions { AllowPrivateAddresses = true });
            }
            catch (DocumentConversionException)
            {
                // A guarded fetch can still fail for other reasons (a malformed response, a
                // cancelled read). What is under test here is that the opt-in flag reaches the
                // image processing mode at all - the outbound connection below proves that either
                // way, and it is one more reason the default is no-network.
            }
        });

        Assert.True(await probe.WasContactedAsync(),
            "The opt-in flag did not reach HtmlToOpenXml's image processing mode.");
    }

    // ---------------------------------------------------------------------------------------
    // Known gap: nothing here pins that allowRemoteImageDownload: true still constructs a
    // GuardedResourceLoader rather than passing null (i.e. behaving like false/Offline).
    //
    // The test above necessarily drives its probe through the RemoteImageOptions overload with
    // AllowPrivateAddresses = true, because GuardedResourceLoader refuses loopback by design and
    // the bool overload can never set that flag - so it cannot be used to prove the bool path
    // reaches the network. Checked directly: BuildPackageAsync(html, null, ct) and
    // BuildPackageAsync(html, new RemoteImageOptions(), ct) produce byte-identical output for an
    // <img> the fetch cannot reach (confirmed against a blocked loopback address, where both
    // paths resolve to "image skipped" without any socket involved) - so there is no black-box
    // signal in the produced document that tells "true mapped to null" and "true mapped to
    // GuardedResourceLoader, then correctly refused" apart. Proving reachability the way the test
    // above does would require a real routable address for the bool path too, which is an
    // external dependency this suite deliberately does not take on elsewhere, and exposing an
    // internal seam to observe the mapping directly is out of scope for a docs/tests-only pass.
    // If a future edit ever changed `allowRemoteImageDownload ? new RemoteImageOptions() : null`
    // in HtmlToDocxConverter/HtmlToPdfConverter to unconditionally pass null, nothing in this
    // suite would catch it.
    // ---------------------------------------------------------------------------------------

    /// <summary>
    /// A loopback TCP socket standing in for a remote image host. It answers the question "did
    /// anything connect?" and serves a real 1x1 bitmap, so the opt-in path converts cleanly
    /// instead of failing on a broken download.
    /// </summary>
    private sealed class LoopbackProbe : IDisposable
    {
        private readonly TcpListener _listener;
        private readonly Action? _onContact;
        private readonly TaskCompletionSource<bool> _contacted =
            new(TaskCreationOptions.RunContinuationsAsynchronously);

        /// <summary>
        /// Starts a loopback listener that serves a one-pixel bitmap to whatever fetches it, and
        /// records that it was contacted.
        /// </summary>
        /// <param name="onContact">
        /// Runs the instant the parser fetches the image, before this probe answers. That is a
        /// deterministic point INSIDE HtmlToOpenXml's work, which is what
        /// <see cref="ConvertAsync_CancelsPartwayThroughALongParse"/> needs and what a wall-clock
        /// timer could not provide.
        /// </param>
        public LoopbackProbe(Action? onContact = null)
        {
            _onContact = onContact;
            _listener = new TcpListener(IPAddress.Loopback, 0);
            _listener.Start();
            Url = $"http://127.0.0.1:{((IPEndPoint)_listener.LocalEndpoint).Port}/logo.bmp";
            _ = AcceptAsync();
        }

        public string Url { get; }

        /// <summary>
        /// True if something connected. The conversion has already finished by the time this is
        /// called, so the short grace period only covers handing the accepted socket over.
        /// </summary>
        public async Task<bool> WasContactedAsync()
            => await Task.WhenAny(_contacted.Task, Task.Delay(500)) == _contacted.Task;

        private async Task AcceptAsync()
        {
            try
            {
                using var client = await _listener.AcceptTcpClientAsync();
                _contacted.TrySetResult(true);
                _onContact?.Invoke();

                var stream = client.GetStream();
                await DrainRequestAsync(stream);

                var image = OnePixelBitmap();
                var header = Encoding.ASCII.GetBytes(
                    "HTTP/1.1 200 OK\r\nContent-Type: image/bmp\r\n" +
                    $"Content-Length: {image.Length}\r\nConnection: close\r\n\r\n");

                await stream.WriteAsync(header);
                await stream.WriteAsync(image);
                await stream.FlushAsync();
            }
            catch (Exception)
            {
                // The listener is torn down at the end of the test; nothing left to answer.
            }
        }

        private static async Task DrainRequestAsync(NetworkStream stream)
        {
            var buffer = new byte[2048];
            var seen = new StringBuilder();
            while (!seen.ToString().Contains("\r\n\r\n"))
            {
                var read = await stream.ReadAsync(buffer);
                if (read == 0) return;
                seen.Append(Encoding.ASCII.GetString(buffer, 0, read));
            }
        }

        /// <summary>
        /// A 1x1 24-bit BMP built by hand: 14-byte file header, 40-byte BITMAPINFOHEADER, one
        /// padded pixel. Hand-rolled rather than committed as an asset so the fixture is readable,
        /// and BMP rather than PNG so no compressor is involved.
        /// </summary>
        private static byte[] OnePixelBitmap()
        {
            var bmp = new byte[58];

            // Disposed, though BinaryWriter writes through per call rather than accumulating like
            // StreamWriter, so nothing is lost without it today. Not worth depending on that.
            using var w = new BinaryWriter(new MemoryStream(bmp));
            w.Write((byte)'B'); w.Write((byte)'M');
            w.Write(58);            // file size
            w.Write(0);             // reserved
            w.Write(54);            // pixel data offset
            w.Write(40);            // BITMAPINFOHEADER size
            w.Write(1);             // width
            w.Write(1);             // height
            w.Write((short)1);      // planes
            w.Write((short)24);     // bits per pixel
            w.Write(0);             // BI_RGB, uncompressed
            w.Write(4);             // image byte size
            w.Write(2835);          // horizontal pixels per metre
            w.Write(2835);          // vertical pixels per metre
            w.Write(0);             // palette colours used
            w.Write(0);             // important colours
            w.Write(new byte[] { 0x00, 0x00, 0xFF, 0x00 }); // one red pixel + row padding
            return bmp;
        }

        public void Dispose() => _listener.Stop();
    }
}
