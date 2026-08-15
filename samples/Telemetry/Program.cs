using System.Diagnostics;
using System.Diagnostics.Metrics;
using DocToolkit;

Console.WriteLine("Telemetry");
Console.WriteLine("=========");

// --- Only one thing here is instrumented ------------------------------------------------------
// DocToolkitTelemetry exposes an ActivitySource name and a Meter name, both "Ank.DocToolkit". Only
// GuardedResourceLoader.FetchAsync - the opt-in remote-image fetch - is instrumented, and that is
// the whole of the scope. Every other call in this library is synchronous, in-process and throws a
// typed exception on failure, so a caller learns as much by timing their own call. The fetch path
// is different: it is the only network access this library ever makes, the allow-or-refuse
// decision happens inside HtmlToOpenXml's pipeline where the caller cannot see it, and a refused
// fetch is deliberately SILENT - the image is skipped and the document still succeeds. Seeing that
// means actually triggering a fetch, which needs the opt-in (RemoteImageOptions) and something to
// fetch from - both arranged below.

// --- Subscribing needs no OpenTelemetry package -------------------------------------------------
// ActivitySource and Meter are System.Diagnostics(.Metrics) types in the shared framework, not
// something the OpenTelemetry package defines. The listeners below are a hand-written, miniature
// version of exactly what OpenTelemetry's SDK attaches when you call
// .WithTracing(t => t.AddSource(...)) / .WithMetrics(m => m.AddMeter(...)) - written by hand here
// only so this sample can stay free of that package, for the same reason DocToolkitTelemetry can
// exist in a package whose whole premise is four constraints on its own dependency graph.

var spans = new List<Activity>();
using var activityListener = new ActivityListener
{
    ShouldListenTo = source => source.Name == DocToolkitTelemetry.ActivitySourceName,
    Sample = (ref ActivityCreationOptions<ActivityContext> _) => ActivitySamplingResult.AllDataAndRecorded,
    ActivityStopped = activity => { lock (spans) spans.Add(activity); },
};
ActivitySource.AddActivityListener(activityListener);

var counts = new List<(string Outcome, string? Host)>();
var recordedBytes = new List<long>();
using var meterListener = new MeterListener();
meterListener.InstrumentPublished = (instrument, listener) =>
{
    if (instrument.Meter.Name == DocToolkitTelemetry.MeterName) listener.EnableMeasurementEvents(instrument);
};
meterListener.SetMeasurementEventCallback<long>((instrument, value, tags, _) =>
{
    if (instrument.Name == "doctoolkit.remote_image.fetches")
    {
        string? outcome = null, host = null;
        foreach (var tag in tags)
        {
            if (tag.Key == "doctoolkit.remote_image.outcome") outcome = tag.Value as string;
            if (tag.Key == "server.address") host = tag.Value as string;
        }
        lock (counts) counts.Add((outcome ?? "(none)", host));
    }
    else if (instrument.Name == "doctoolkit.remote_image.bytes")
    {
        lock (recordedBytes) recordedBytes.Add(value);
    }
});
meterListener.Start();

// Prints only the spans captured since the last call - each attempt below reports its own.
var reportedUpTo = 0;
void ReportNewSpans()
{
    List<Activity> fresh;
    lock (spans)
    {
        fresh = spans.Skip(reportedUpTo).ToList();
        reportedUpTo = spans.Count;
    }

    foreach (var span in fresh)
    {
        var outcome = span.GetTagItem("doctoolkit.remote_image.outcome");
        var host = span.GetTagItem("server.address");
        var port = span.GetTagItem("server.port");
        var bytes = span.GetTagItem("doctoolkit.remote_image.bytes");
        Console.WriteLine(
            $"  span   outcome={outcome,-15} host={host}:{port}" +
            (bytes is not null ? $" bytes={bytes}" : ""));
    }
}

// --- Something to fetch from, without reaching outside this process ----------------------------
// A real "ok" outcome needs a real HTTP response. Rather than reach the actual internet - which
// would make this sample non-deterministic, occasionally slow, and a quiet argument that network
// access is normal here - it brings its own tiny loopback origin, reachable only at 127.0.0.1 and
// only for the life of this process.

// The same 137-byte PNG samples/DocxImages uses, so this sample carries no binary asset either.
const string LogoBase64 =
    "iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAIAAAAlC+aJAAAAUElEQVR42u3PQQkAAAgEsOtlFCsZ2gi+hcEKLNXzWgQEBA" +
    "QEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQErsACwGghD5ay/wAAAAAASUVORK5CYII=";

using var origin = new LoopbackImageServer(Convert.FromBase64String(LogoBase64));
Console.WriteLine($"\nLoopback origin: {origin.BaseUrl} (this process, this machine only)");

// The query string carries a fake signed token, the way a real image CDN URL often does, so the
// "host only, never the URL" claim below is something you can check against this output rather
// than something you have to take on faith.
var requestUrl = $"{origin.BaseUrl}/logo.png?sig=SUPERSECRETTOKEN";
var html = $"<p>Before</p><p><img src=\"{requestUrl}\"></p><p>After</p>";

// --- Attempt 1: opted in, pointed at our own loopback origin ------------------------------------
// AllowPrivateAddresses has to be true for this one to reach anything at all - loopback is blocked
// by default, which is exactly what attempt 2 below demonstrates.

Console.WriteLine("\nAttempt 1 - RemoteImageOptions { AllowPrivateAddresses = true }, pointed at the origin above");
Console.WriteLine($"  Requested with: {requestUrl}");

var allowed = new RemoteImageOptions { AllowPrivateAddresses = true, Timeout = TimeSpan.FromSeconds(5) };
byte[] withImage = await HtmlToDocxConverter.ConvertAsync(html, allowed);

Console.WriteLine($"  Origin saw    : {origin.Connections} connection(s)");
ReportNewSpans();
Console.WriteLine($"  Document      : {withImage.Length:N0} bytes, image embedded");
Console.WriteLine("  Notice the span's host has no path, no query string and no token in it - only " +
    "'server.address' and 'server.port' are ever recorded, which is what keeps a signed URL out of " +
    "wherever this telemetry ends up.");

// --- Attempt 2: default options, same origin - refused before any connection --------------------
// Passing a RemoteImageOptions at all is what opts in; passing one with every default left in
// place is what refuses this. AllowPrivateAddresses defaults to false, so the address check
// refuses our own loopback origin before a socket is even opened - no listener needed to observe
// that, only to observe that nothing reached it.

Console.WriteLine("\nAttempt 2 - RemoteImageOptions with every default left in place, same origin");
Console.WriteLine($"  Requested with: {requestUrl}");

var beforeAttempt2 = origin.Connections;
var refused = new RemoteImageOptions();
byte[] withoutImage = await HtmlToDocxConverter.ConvertAsync(html, refused);

Console.WriteLine($"  Origin saw    : {origin.Connections - beforeAttempt2} new connection(s) (refused before any socket opened)");
ReportNewSpans();
Console.WriteLine($"  Document      : {withoutImage.Length:N0} bytes, image silently absent");
Console.WriteLine($"  Text survived : {DocxEditor.ExtractText(withoutImage).Replace("\n", " / ")}");

// --- Attempt 3: allowed to reach private addresses, but not this host ---------------------------
// AllowPrivateAddresses is true this time, so the address check alone would let this through - but
// the allow-list is checked first, and 127.0.0.1 was never added to it. Same silent skip, a
// different reason, and still no connection to the origin.

Console.WriteLine("\nAttempt 3 - AllowPrivateAddresses = true, but 127.0.0.1 is not on AllowedHosts");
Console.WriteLine($"  Requested with: {requestUrl}");

var beforeAttempt3 = origin.Connections;
var notOnList = new RemoteImageOptions { AllowPrivateAddresses = true };
notOnList.AllowedHosts.Add("cdn.contoso.example");
byte[] alsoWithoutImage = await HtmlToDocxConverter.ConvertAsync(html, notOnList);

Console.WriteLine($"  Origin saw    : {origin.Connections - beforeAttempt3} new connection(s) (host not on the allow-list)");
ReportNewSpans();
Console.WriteLine($"  Document      : {alsoWithoutImage.Length:N0} bytes, image silently absent");

// --- What a real OpenTelemetry pipeline would show ----------------------------------------------

Console.WriteLine("\nMeter counts (doctoolkit.remote_image.fetches, by outcome):");
foreach (var group in counts.GroupBy(c => c.Outcome).OrderBy(g => g.Key))
    Console.WriteLine($"  {group.Key,-15} x{group.Count()}");

Console.WriteLine($"\nMeter measurements (doctoolkit.remote_image.bytes): {string.Join(", ", recordedBytes)}");

// Measured, not assumed: attempts 2 and 3 each reported two spans for one <img>, the successful
// attempt reported one. GuardedResourceLoader.FetchAsync's own doc comment warns that HtmlToOpenXml
// can call it more than once per image, concurrently - this is that showing up in the counters,
// and it is a reason to read this meter as "fetch attempts", not "images requested".
Console.WriteLine("\nNotice attempts 2 and 3 each produced TWO spans for one <img> tag, while the " +
    "successful attempt produced one - GuardedResourceLoader.FetchAsync can be called more than " +
    "once per image. Read the counter above as fetch ATTEMPTS, not images requested.");

Console.WriteLine("\nA real consumer gets all of this from two lines, no ActivityListener required:");
Console.WriteLine("  builder.Services.AddOpenTelemetry()");
Console.WriteLine("      .WithTracing(t => t.AddSource(DocToolkitTelemetry.ActivitySourceName))");
Console.WriteLine("      .WithMetrics(m => m.AddMeter(DocToolkitTelemetry.MeterName));");

Console.WriteLine("\nDone.");
