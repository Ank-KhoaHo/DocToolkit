using System.Diagnostics;
using System.Diagnostics.Metrics;

namespace DocToolkit;

/// <summary>
/// The names to subscribe to for DocToolkit's telemetry.
///
/// <code>
/// builder.Services.AddOpenTelemetry()
///     .WithTracing(t => t.AddSource(DocToolkitTelemetry.ActivitySourceName))
///     .WithMetrics(m => m.AddMeter(DocToolkitTelemetry.MeterName));
/// </code>
///
/// <b>Only the opt-in remote-image fetch is instrumented</b>, and that is a deliberate scope rather
/// than a first instalment. Every other call in this package is one synchronous, in-process,
/// stateless operation that throws a typed exception on failure — a caller can time and log around
/// it and learn everything a span would tell them. The fetch path is the exception: it is the only
/// place this library reaches the network, the decision to allow or refuse a host happens deep
/// inside HtmlToOpenXml's pipeline, and a refused fetch is deliberately <i>silent</i> — the image is
/// skipped and the document still succeeds. Without this, a consumer who enabled remote images had
/// no way to find out that an image never arrived, or why.
/// </summary>
public static class DocToolkitTelemetry
{
    /// <summary>The <see cref="System.Diagnostics.ActivitySource"/> name. Matches the package id.</summary>
    public const string ActivitySourceName = "Ank.DocToolkit";

    /// <summary>The <see cref="System.Diagnostics.Metrics.Meter"/> name. Matches the package id.</summary>
    public const string MeterName = "Ank.DocToolkit";
}

/// <summary>
/// The instruments themselves.
///
/// <b>Nothing here records a full URL — only the host.</b> A query string routinely carries a
/// signed token or an API key, and telemetry is exported off the machine and retained; a span
/// attribute is one of the easier ways to leak a credential into somebody's observability vendor.
/// The host is what makes a refusal diagnosable, and is the whole of what is needed for that.
///
/// Both types cost nothing when nobody is listening: <c>ActivitySource.StartActivity</c>
/// returns null with no listener, and a <see cref="Counter{T}"/> with no collector does not
/// allocate. Neither needs a package reference — both live in the shared framework on the targeted
/// frameworks, so this adds nothing to the resolved dependency graph, which is the only reason it
/// can exist in a package whose premise is four constraints on that graph.
/// </summary>
internal static class Telemetry
{
    private static readonly string Version =
        typeof(Telemetry).Assembly.GetName().Version?.ToString() ?? "0.0.0";

    internal static readonly ActivitySource Source =
        new(DocToolkitTelemetry.ActivitySourceName, Version);

    private static readonly Meter Meter = new(DocToolkitTelemetry.MeterName, Version);

    /// <summary>
    /// One per attempted remote-image fetch, tagged with its outcome. The counter rather than the
    /// span is what answers "are my images arriving?" cheaply, since a consumer may sample traces.
    /// </summary>
    internal static readonly Counter<long> RemoteImageFetches =
        Meter.CreateCounter<long>(
            "doctoolkit.remote_image.fetches",
            unit: "{fetch}",
            description: "Remote image fetches attempted, by outcome.");

    /// <summary>Bytes actually read, for fetches that produced an image.</summary>
    internal static readonly Histogram<long> RemoteImageBytes =
        Meter.CreateHistogram<long>(
            "doctoolkit.remote_image.bytes",
            unit: "By",
            description: "Size of successfully fetched remote images.");

    /// <summary>
    /// Records the outcome on both the span and the counter, so a caller gets the same answer
    /// whether they sample traces or scrape metrics.
    /// </summary>
    internal static void RecordOutcome(Activity? activity, string outcome, string host)
    {
        activity?.SetTag("doctoolkit.remote_image.outcome", outcome);
        if (outcome != Outcomes.Ok) activity?.SetStatus(ActivityStatusCode.Error, outcome);

        RemoteImageFetches.Add(
            1,
            new KeyValuePair<string, object?>("doctoolkit.remote_image.outcome", outcome),
            new KeyValuePair<string, object?>("server.address", host));
    }

    /// <summary>
    /// The closed set of fetch outcomes. Named constants rather than literals so a metric consumer's
    /// dashboard cannot be broken by a typo, and so the set is greppable when one is added.
    /// </summary>
    internal static class Outcomes
    {
        /// <summary>An image was fetched and is being embedded.</summary>
        internal const string Ok = "ok";

        /// <summary>The URL was not http or https.</summary>
        internal const string SchemeRefused = "scheme_refused";

        /// <summary>An allow-list was configured and this host was not on it.</summary>
        internal const string HostNotAllowed = "host_not_allowed";

        /// <summary>The host resolved to a loopback, private, link-local or CGNAT address.</summary>
        internal const string BlockedAddress = "blocked_address";

        /// <summary>The server answered, with a non-success status.</summary>
        internal const string HttpError = "http_error";

        /// <summary>The body exceeded <c>RemoteImageOptions.MaxBytesPerImage</c> and was abandoned.</summary>
        internal const string TooLarge = "too_large";

        /// <summary>Timed out, unreachable, or the transport failed. The common air-gapped case.</summary>
        internal const string Failed = "failed";
    }
}
