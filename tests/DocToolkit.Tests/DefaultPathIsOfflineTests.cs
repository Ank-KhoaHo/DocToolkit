using System.Diagnostics;
using Xunit;

namespace DocToolkit.Tests;

/// <summary>
/// Proves the DEFAULT conversion overloads use the offline resource loader, by asserting they emit
/// no fetch telemetry at all.
///
/// <b>Why the air-gap suite does not already cover this.</b> `AirGapGuardTests` proves that nothing
/// reaches its loopback probe. That is a different claim, and mutation testing found the gap:
/// flipping `allowRemoteImageDownload: false` to `true` in these overloads' delegation
/// <b>survives that entire suite</b>. With `true` the GUARDED loader runs, and its defaults block
/// loopback and private addresses — so the probe sees zero connections either way. "Nothing reached
/// our probe" is not "the offline loader was used", and `CLAUDE.md` says which loader is passed "is
/// the whole of the offline/online decision".
///
/// <b>The discriminator needs no network.</b> `GuardedResourceLoader.FetchAsync` starts a
/// `RemoteImage.Fetch` activity <i>before</i> deciding, so it emits telemetry even when it refuses.
/// `OfflineResourceLoader` refuses at `SupportsProtocol` and is never asked to fetch, so it emits
/// nothing. Absence of an activity therefore distinguishes the two loaders exactly.
///
/// <b>An ActivityListener is PROCESS-WIDE</b>, and that broke `TelemetryTests` twice — it sees
/// activities raised by suites running in parallel. Every assertion here filters on a
/// `server.address` no other test uses, for that reason. Do not reuse an address from another file.
/// </summary>
public class DefaultPathIsOfflineTests : IDisposable
{
    // TEST-NET-2 (RFC 5737) and a private address, neither used by any other suite. The filter is
    // what keeps a parallel suite's activity from being read as this test's.
    private const string DefaultPathHost = "198.51.100.181";
    private const string OptInHost = "198.51.100.182";

    private readonly List<Activity> _captured = [];
    private readonly ActivityListener _listener;

    public DefaultPathIsOfflineTests()
    {
        _listener = new ActivityListener
        {
            ShouldListenTo = source => source.Name == DocToolkitTelemetry.ActivitySourceName,
            Sample = (ref ActivityCreationOptions<ActivityContext> _) =>
                ActivitySamplingResult.AllDataAndRecorded,
            ActivityStopped = activity => { lock (_captured) _captured.Add(activity); },
        };

        ActivitySource.AddActivityListener(_listener);
    }

    public void Dispose() => _listener.Dispose();

    private List<Activity> For(string host)
    {
        lock (_captured)
            return _captured
                .Where(a => (a.GetTagItem("server.address") as string) == host)
                .ToList();
    }

    private static string HtmlNaming(string host) =>
        $"<p>text</p><img src=\"http://{host}/logo.png\" alt=\"remote\">";

    // ---- THE POSITIVE CONTROL, and it has to come first --------------------------------------

    [Fact]
    public async Task TheOptInOverload_EmitsFetchTelemetry_SoAnAbsenceBelowMeansSomething()
    {
        // Without this, every assertion in this file could be passing because the listener is
        // broken, the source name changed, or the markup never reaches a loader at all. It is the
        // same control the air-gap suites carry, for the same reason.
        //
        // AllowPrivateAddresses stays FALSE: the guarded loader starts its activity before it
        // decides, so a refusal still emits one. That is what makes this fast and offline.
        await HtmlToDocxConverter.ConvertAsync(
            HtmlNaming(OptInHost), new RemoteImageOptions(), CancellationToken.None);

        var seen = For(OptInHost);
        Assert.NotEmpty(seen);
        Assert.All(seen, a => Assert.Equal("RemoteImage.Fetch", a.OperationName));
    }

    // ---- The mutant this file exists to kill -------------------------------------------------

    [Fact]
    public async Task TheDefaultByteArrayOverload_NeverConsultsAFetchingLoader()
    {
        await HtmlToDocxConverter.ConvertAsync(HtmlNaming(DefaultPathHost), CancellationToken.None);

        Assert.Empty(For(DefaultPathHost));
    }

    [Fact]
    public async Task TheDefaultStreamOverload_NeverConsultsAFetchingLoader()
    {
        // The overload whose `allowRemoteImageDownload: false` mutation survived the air-gap suite.
        using var destination = new MemoryStream();

        await HtmlToDocxConverter.ConvertAsync(
            HtmlNaming(DefaultPathHost), destination, CancellationToken.None);

        Assert.Empty(For(DefaultPathHost));
        Assert.True(destination.Length > 0, "the conversion still has to produce a document");
    }

    [Fact]
    public async Task TheDefaultPdfOverload_NeverConsultsAFetchingLoader()
    {
        // HTML → PDF pivots through DOCX, so it shares the loader decision. Covered here because
        // sharing a code path is not the same as sharing an assertion.
        await HtmlToPdfConverter.ConvertAsync(HtmlNaming(DefaultPathHost), CancellationToken.None);

        Assert.Empty(For(DefaultPathHost));
    }
}
