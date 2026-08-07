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

    [Fact]
    public async Task ConvertAsync_CancelsPartwayThroughALongParse()
    {
        // Sized so the parse takes the best part of a second (measured ~0.8-1.8 s), so a token
        // cancelled after 150 ms is guaranteed to land inside HtmlToOpenXml's work rather than at
        // the guard on the way in.
        var sb = new StringBuilder();
        for (var i = 0; i < 60_000; i++)
            sb.Append("<p>Paragraph ").Append(i).Append(" of the stress fixture.</p>");

        using var cts = new CancellationTokenSource();
        cts.CancelAfter(TimeSpan.FromMilliseconds(150));

        // The exception alone proves what this test exists to prove. ConvertAsync checks the token
        // once on entry - which cannot fire here, since the token is still live when the call
        // starts - and then hands it to ParseBody. There is no check after the parse. So if the
        // token were NOT reaching the parser, this call would run to completion and return a
        // document: no exception at all, and this assertion fails.
        //
        // An earlier version also asserted the whole thing finished inside 600 ms. That added
        // nothing the above does not already establish, and made the test depend on how loaded the
        // CI runner happened to be - it failed three times in one day on pull requests that changed
        // only Markdown and a sample. See backlog B10.
        await Assert.ThrowsAnyAsync<OperationCanceledException>(
            () => HtmlToDocxConverter.ConvertAsync(sb.ToString(), cts.Token));
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
        private readonly TaskCompletionSource<bool> _contacted =
            new(TaskCreationOptions.RunContinuationsAsynchronously);

        public LoopbackProbe()
        {
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
