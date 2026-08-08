using System.Diagnostics;
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
}
