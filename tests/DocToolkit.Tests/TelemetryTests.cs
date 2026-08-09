using System.Diagnostics;
using System.Diagnostics.Metrics;
using Xunit;
using Xunit.Abstractions;

namespace DocToolkit.Tests;

/// <summary>
/// The remote-image fetch telemetry.
///
/// Driven through the real <see cref="GuardedResourceLoader"/> against a real loopback listener,
/// not by calling the instruments directly — an assertion that <c>Telemetry.RecordOutcome</c> sets a
/// tag would pass while nothing in the fetch path ever called it.
///
/// The point of this telemetry is the case a caller cannot otherwise see: a refused fetch is
/// deliberately <b>silent</b>, so the image is skipped and the document still succeeds. On an
/// air-gapped host every remote image lands there.
/// </summary>
public class TelemetryTests : IDisposable
{
    private readonly ITestOutputHelper _output;
    private readonly List<Activity> _captured = new();
    private readonly ActivityListener _listener;

    public TelemetryTests(ITestOutputHelper output)
    {
        _output = output;

        // AllDataAndRecorded, not PropagationData: without it Activity.Current is created but tags
        // are dropped, and every SetTag assertion below would fail against working code.
        _listener = new ActivityListener
        {
            ShouldListenTo = source => source.Name == DocToolkitTelemetry.ActivitySourceName,
            Sample = (ref ActivityCreationOptions<ActivityContext> _) =>
                ActivitySamplingResult.AllDataAndRecorded,
            ActivityStopped = activity =>
            {
                lock (_captured) _captured.Add(activity);
            },
        };

        ActivitySource.AddActivityListener(_listener);
    }

    public void Dispose() => _listener.Dispose();

    private static string? Outcome(Activity activity) =>
        activity.GetTagItem("doctoolkit.remote_image.outcome") as string;

    /// <summary>
    /// The one span recorded for <paramref name="host"/>.
    ///
    /// <b>Filtered by host rather than simply asserted to be the only span.</b> An
    /// <see cref="ActivityListener"/> is PROCESS-WIDE: it sees every activity from every source it
    /// matches, including ones raised by <c>AirGapGuardTests</c> and <c>RemoteImageGuardTests</c>
    /// running in parallel in another collection. An unfiltered <c>Assert.Single</c> therefore fails
    /// intermittently, for a reason with nothing to do with the code under test — it did exactly
    /// that on the first full run, on net8.0 only, capturing a <c>blocked_address</c> span from a
    /// neighbouring test. Each test below uses a distinct host so it can select its own.
    /// </summary>
    private Activity Single(string host)
    {
        lock (_captured)
        {
            _output.WriteLine("captured: " + string.Join(
                ", ", _captured.Select(a => $"{a.GetTagItem("server.address")}={Outcome(a)}")));

            return Assert.Single(
                _captured.Where(a => (a.GetTagItem("server.address") as string) == host).ToList());
        }
    }

    /// <summary>
    /// The one span for <paramref name="port"/>.
    ///
    /// Host is not enough for the probe-driven tests below: a <c>LoopbackProbe</c> is always on
    /// 127.0.0.1, which every other loopback suite also uses. The PORT is OS-assigned per probe
    /// instance, so it is the only thing that separates two spans in a process-wide listener.
    /// </summary>
    private Activity SingleOnPort(int port)
    {
        lock (_captured)
        {
            return Assert.Single(
                _captured.Where(a => (a.GetTagItem("server.port") as int?) == port).ToList());
        }
    }

    // The listener has to be able to see a span before any assertion about one means anything -
    // the same reasoning as AirGapGuardTests' positive probe test.
    [Fact]
    public async Task TheListenerCapturesASpan()
    {
        var loader = new GuardedResourceLoader(new RemoteImageOptions());

        await loader.FetchAsync(new Uri("ftp://example.invalid/x.png"));

        Assert.NotEmpty(_captured);
        Assert.Equal("RemoteImage.Fetch", Single("example.invalid").OperationName);
    }

    [Fact]
    public async Task ANonHttpSchemeIsReportedAsRefused()
    {
        var loader = new GuardedResourceLoader(new RemoteImageOptions());

        await loader.FetchAsync(new Uri("ftp://scheme-refused.test/x.png"));

        Assert.Equal("scheme_refused", Outcome(Single("scheme-refused.test")));
    }

    // The case the whole feature exists for: a private address is refused deep inside the fetch,
    // the image is skipped, and the conversion still succeeds. Nothing else tells the caller.
    [Fact]
    public async Task ABlockedPrivateAddressIsReportedWithItsHost()
    {
        var loader = new GuardedResourceLoader(new RemoteImageOptions());

        var resource = await loader.FetchAsync(new Uri("http://192.168.77.77/logo.png"));

        Assert.Null(resource);
        var activity = Single("192.168.77.77");
        Assert.Equal("blocked_address", Outcome(activity));
        Assert.Equal("192.168.77.77", activity.GetTagItem("server.address"));
        Assert.Equal(ActivityStatusCode.Error, activity.Status);
    }

    [Fact]
    public async Task AHostOutsideTheAllowListIsReportedAsSuch()
    {
        // AllowedHosts is get-only and mutated in place - a settable property would let a
        // caller drop in an options object that lost a restrictive default.
        var options = new RemoteImageOptions();
        options.AllowedHosts.Add("images.example.com");
        var loader = new GuardedResourceLoader(options);

        await loader.FetchAsync(new Uri("http://elsewhere.example.com/x.png"));

        Assert.Equal("host_not_allowed", Outcome(Single("elsewhere.example.com")));
    }

    // An unroutable host is the air-gapped case: every fetch fails this way, silently.
    [Fact]
    public async Task AnUnreachableHostIsReportedAsFailed()
    {
        var loader = new GuardedResourceLoader(new RemoteImageOptions
        {
            AllowPrivateAddresses = true,
            Timeout = TimeSpan.FromMilliseconds(300),
        });

        await loader.FetchAsync(new Uri("http://192.0.2.77/x.png"));

        var activity = Single("192.0.2.77");
        Assert.Equal("failed", Outcome(activity));
        Assert.NotNull(activity.GetTagItem("exception.type"));
    }

    // A URL carries credentials in its query string often enough that recording one is a real way
    // to leak a token into an observability vendor. Only the host is ever recorded.
    [Fact]
    public async Task NoTagCarriesTheFullUrlOrItsQueryString()
    {
        var loader = new GuardedResourceLoader(new RemoteImageOptions());

        await loader.FetchAsync(new Uri("http://10.77.77.77/logo.png?sig=SUPERSECRETTOKEN"));

        var activity = Single("10.77.77.77");
        foreach (var tag in activity.TagObjects)
        {
            Assert.DoesNotContain("SUPERSECRETTOKEN", tag.Value?.ToString() ?? "", StringComparison.Ordinal);
            Assert.DoesNotContain("logo.png", tag.Value?.ToString() ?? "", StringComparison.Ordinal);
        }

        Assert.Equal("10.77.77.77", activity.GetTagItem("server.address"));

        // server.port is recorded and is deliberately not covered by the assertion above: a port
        // number cannot carry a credential the way a query string can.
        Assert.Equal(80, activity.GetTagItem("server.port"));
    }

    // Costing nothing without a listener is the claim that lets this exist in a package whose
    // premise is what it does NOT drag in. StartActivity returns null; the code must not care.
    [Fact]
    public async Task WithNoListenerTheFetchStillBehavesIdentically()
    {
        _listener.Dispose();

        var loader = new GuardedResourceLoader(new RemoteImageOptions());
        var resource = await loader.FetchAsync(new Uri("http://192.168.66.66/x.png"));

        Assert.Null(resource);
        lock (_captured)
        {
            Assert.DoesNotContain(
                _captured, a => (a.GetTagItem("server.address") as string) == "192.168.66.66");
        }
    }

    // The remaining fetch outcomes. These were added with the telemetry and never exercised - a
    // successful fetch, a non-success status, and a body past MaxBytesPerImage - which mutation
    // testing reported as surviving mutants in the recording calls themselves.
    //
    // Each needs a server that behaves a particular way, which is why they live here with the
    // LoopbackProbe rather than beside the address-classification tests.

    private static async Task NotFoundResponder(
        System.Net.Sockets.NetworkStream stream, string requestLine, CancellationToken ct)
    {
        await stream.WriteAsync(System.Text.Encoding.ASCII.GetBytes(
            "HTTP/1.1 404 Not Found\r\nContent-Length: 0\r\nConnection: close\r\n\r\n"), ct);
        await stream.FlushAsync(ct);
    }

    private static async Task OversizedResponder(
        System.Net.Sockets.NetworkStream stream, string requestLine, CancellationToken ct)
    {
        var body = new byte[64 * 1024];
        await stream.WriteAsync(System.Text.Encoding.ASCII.GetBytes(
            "HTTP/1.1 200 OK\r\nContent-Type: image/bmp\r\n"
            + $"Content-Length: {body.Length}\r\nConnection: close\r\n\r\n"), ct);
        await stream.WriteAsync(body, ct);
        await stream.FlushAsync(ct);
    }

    [Fact]
    public async Task ASuccessfulFetchIsReportedAsOkWithItsSize()
    {
        using var probe = new LoopbackProbe(_output);
        var loader = new GuardedResourceLoader(
            new RemoteImageOptions { AllowPrivateAddresses = true });

        var resource = await loader.FetchAsync(new Uri($"{probe.BaseUrl}/logo.bmp"));

        Assert.NotNull(resource);
        var activity = SingleOnPort(probe.Port);
        Assert.Equal("ok", Outcome(activity));
        Assert.Equal(ActivityStatusCode.Unset, activity.Status);

        // The size tag is what makes the span useful for "why is my document 40 MB"; without it a
        // successful fetch is indistinguishable from any other.
        var bytes = Assert.IsType<int>(activity.GetTagItem("doctoolkit.remote_image.bytes"));
        Assert.True(bytes > 0, "A successful fetch recorded no size.");
    }

    [Fact]
    public async Task ANonSuccessStatusIsReportedWithItsCode()
    {
        using var probe = new LoopbackProbe(_output, NotFoundResponder);
        var loader = new GuardedResourceLoader(
            new RemoteImageOptions { AllowPrivateAddresses = true });

        var resource = await loader.FetchAsync(new Uri($"{probe.BaseUrl}/missing.bmp"));

        Assert.Null(resource);
        var activity = SingleOnPort(probe.Port);
        Assert.Equal("http_error", Outcome(activity));
        Assert.Equal(404, activity.GetTagItem("http.response.status_code"));
    }

    // Distinguishing "too large" from a plain failure is the point: one means raise
    // MaxBytesPerImage, the other means the host is unreachable. A single "failed" would leave a
    // consumer guessing.
    [Fact]
    public async Task ABodyPastTheCapIsReportedAsTooLarge()
    {
        using var probe = new LoopbackProbe(_output, OversizedResponder);
        var loader = new GuardedResourceLoader(new RemoteImageOptions
        {
            AllowPrivateAddresses = true,
            MaxBytesPerImage = 1024,
        });

        var resource = await loader.FetchAsync(new Uri($"{probe.BaseUrl}/huge.bmp"));

        Assert.Null(resource);
        Assert.Equal("too_large", Outcome(SingleOnPort(probe.Port)));
    }


    // =============================================================================================
    // The two recording calls themselves, rather than the outcome they record.
    // =============================================================================================
    //
    // Mutation testing found four survivors here: deleting the url.scheme tag, blanking its key,
    // deleting the RemoteImageBytes.Record call, and blanking the server.address key on it. Every
    // existing test asserted the OUTCOME, which those mutants leave untouched - so the fetch could
    // stop reporting its scheme, or stop recording a metric at all, and the suite stayed green.

    /// <summary>
    /// The scheme is on the span because "was this http or https" is the first question asked of a
    /// fetch that misbehaved, and the URL itself is deliberately never recorded - so without this
    /// tag there is nothing to answer it with.
    /// </summary>
    [Fact]
    public async Task ASuccessfulFetchTagsTheSchemeItUsed()
    {
        using var probe = new LoopbackProbe(_output);
        var loader = new GuardedResourceLoader(
            new RemoteImageOptions { AllowPrivateAddresses = true });

        await loader.FetchAsync(new Uri($"{probe.BaseUrl}/logo.bmp"));

        Assert.Equal("http", SingleOnPort(probe.Port).GetTagItem("url.scheme"));
    }

    // A body length no other test in this repository serves, so a measurement carrying it can only
    // have come from this test - which matters because MeterListener, like ActivityListener, is
    // PROCESS-WIDE and every loopback suite reports server.address=127.0.0.1.
    private const int DistinctiveBodyLength = 4242;

    private static async Task DistinctiveSizeResponder(
        System.Net.Sockets.NetworkStream stream, string requestLine, CancellationToken ct)
    {
        var body = new byte[DistinctiveBodyLength];
        await stream.WriteAsync(System.Text.Encoding.ASCII.GetBytes(
            "HTTP/1.1 200 OK\r\nContent-Type: image/bmp\r\n"
            + $"Content-Length: {body.Length}\r\nConnection: close\r\n\r\n"), ct);
        await stream.WriteAsync(body, ct);
        await stream.FlushAsync(ct);
    }

    /// <summary>
    /// The size histogram is a METRIC, not a span tag, so no ActivityListener can see it. Nothing
    /// in this suite subscribed to the meter at all, which is why deleting the Record call survived.
    /// </summary>
    [Fact]
    public async Task ASuccessfulFetchRecordsItsSizeAgainstTheHostOnTheMeter()
    {
        var measurements = new List<(long Value, string? Host)>();

        using var meterListener = new MeterListener();
        meterListener.InstrumentPublished = (instrument, listener) =>
        {
            if (instrument.Meter.Name == DocToolkitTelemetry.MeterName
                && instrument.Name == "doctoolkit.remote_image.bytes")
            {
                listener.EnableMeasurementEvents(instrument);
            }
        };
        meterListener.SetMeasurementEventCallback<long>((_, value, tags, _) =>
        {
            string? host = null;
            foreach (var tag in tags)
            {
                if (tag.Key == "server.address") host = tag.Value as string;
            }

            lock (measurements) measurements.Add((value, host));
        });
        meterListener.Start();

        using var probe = new LoopbackProbe(_output, DistinctiveSizeResponder);
        var loader = new GuardedResourceLoader(
            new RemoteImageOptions { AllowPrivateAddresses = true });

        var resource = await loader.FetchAsync(new Uri($"{probe.BaseUrl}/logo.bmp"));
        Assert.NotNull(resource);

        lock (measurements)
        {
            // Contains rather than Single: the listener is process-wide, so a concurrent suite may
            // add its own measurements. It cannot add one of THIS length.
            Assert.Contains(measurements, m => m.Value == DistinctiveBodyLength && m.Host == "127.0.0.1");
        }
    }
}
