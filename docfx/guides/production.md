# Running in production

The things that stop being theoretical once this is serving traffic: memory, containers, fonts,
trimming, and knowing when a document came out wrong.

## Streaming

Every conversion has a `Stream` pair alongside its `byte[]` form.

[!code-csharp[](../../samples/LargeFileStreaming/Program.cs#stream)]

**They are not a memory optimisation, and it is worth being clear about that.** Measured, the same
edit costs about the same either way. A `.docx` or `.xlsx` is a ZIP, and a ZIP's central directory
is at the **end** of the file — you cannot process the first entry until you have seen the last
byte. Peak memory is dominated by the OOXML object model, not by how the bytes arrived.

What they actually buy you is that `source` may be **forward-only and non-seekable**: an HTTP
request body, a network stream, a pipe. Those cannot be handed to a ZIP reader directly, and the
alternative — `ReadAllBytes` into a `byte[]` first — is exactly what the `Stream` overload does for
you, correctly, honouring your `CancellationToken` while it does it.

The sample above proves it with a stream that *refuses* to seek, so a rewind throws rather than
quietly succeeding on a `MemoryStream` and misleading everyone.

> [!IMPORTANT]
> **DocToolkit never closes a stream you hand it.** The stream is yours. That is what lets
> `destination` be a response body you go on writing to — and equally what makes disposal your job.

If your documents are large enough that memory is a real constraint, the lever is concurrency —
how many conversions you allow at once — not which overload you call.

## Containers

The library needs no native binaries, so a plain .NET runtime image is enough. There is nothing to
`apt-get install` and no LibreOffice layer.

```dockerfile
FROM mcr.microsoft.com/dotnet/runtime:8.0
```

> [!WARNING]
> Use `runtime`, not `runtime-deps`. The `runtime-deps` image carries the *native dependencies* of
> the .NET runtime but not the runtime itself — it is for self-contained publishes. A
> framework-dependent app copied into it fails at startup, not at build time.

The [Container sample](https://github.com/Ank-KhoaHo/DocToolkit/tree/main/samples/Container) is
built **and run** in CI with `--network none`, which is how the offline claim stays true rather
than remaining an assertion in a README.

### Fonts in containers

A slim image has no fonts installed, and that changes the PDF you get:

| Where | What happens | Size of the same invoice |
|---|---|---|
| A dev machine with system fonts | The font is **embedded** in the PDF | ~167 KB, carrying Arial-Regular and Arial-Bold |
| A slim container, no fonts | Falls back to the **base-14 standard fonts** (Helvetica) | ~1.5 KB, embedding nothing |

**Both are valid PDFs and both render.** Arial and Helvetica are metric-compatible, so line breaks
do not move. But the glyphs are not identical, and the files are nowhere near the same size — so a
PDF built in your container will never be byte-identical to one built on your laptop.

The practical rules: do not assert on PDF byte size or hash across environments, and if you need a
specific typeface, install it in the image.

## Trimming

Both assemblies are marked trimmable, so a trimmed publish keeps working and gets smaller. CI
proves it on every pull request: it trim-publishes a probe application, asserts that **no trim
warning belongs to DocToolkit**, and then *runs* the trimmed binary. A library that publishes
warning-free but crashes on first call would not pass.

That said, the underlying OOXML libraries use reflection in places. If you trim aggressively and
something fails, check the trim warnings from *those* assemblies before suspecting this one.

## Nothing reaches the network unless you ask

No conversion opens a socket by default. `<link rel="stylesheet">` is not fetched. `<img src>`
pointing at a URL is dropped, not downloaded. This is not a configuration default that could drift
— it is the absence of an outbound path except through one guarded loader.

When you do opt in, per call or through `DocToolkitOptions`, the guard is an allow-list checked
before any connection is attempted, followed by an address check that refuses loopback,
link-local, and every private range. See
[Images the HTML points at](html-to-word-and-pdf.md#images-the-html-points-at).

## Telemetry

One `ActivitySource` and one `Meter`, both named `Ank.DocToolkit`, exposed as constants on
@DocToolkit.DocToolkitTelemetry:

```csharp
builder.Services.AddOpenTelemetry()
    .WithTracing(t => t.AddSource(DocToolkitTelemetry.ActivitySourceName))
    .WithMetrics(m => m.AddMeter(DocToolkitTelemetry.MeterName));
```

This adds **no packages** — `ActivitySource` and `Meter` are in the shared framework on both target
frameworks — and costs nothing when nobody subscribes.

**Only the opt-in remote-image fetch is instrumented**, deliberately. Every other call is one
synchronous, in-process, stateless operation that throws a typed exception when it fails; you can
time and log around it and learn everything a span would have told you.

The fetch path is genuinely different, and this is the reason the instrumentation exists: it is the
only place the library touches the network, the allow-or-refuse decision happens deep inside
HtmlToOpenXml's pipeline where you cannot see it, and **a refused fetch is silent** — the image is
skipped and your document still succeeds. On an air-gapped host every remote image lands there.
Without telemetry, nothing tells you an image never arrived, or why.

| Instrument | What it records |
|---|---|
| `doctoolkit.remote_image.fetches` | Attempts by outcome: `ok`, `scheme_refused`, `host_not_allowed`, `blocked_address`, `http_error`, `too_large`, `failed` |
| `doctoolkit.remote_image.bytes` | The size of images that actually arrived |

> [!IMPORTANT]
> **Only the host is recorded, never the URL.** A query string routinely carries a signed token,
> and telemetry leaves the machine and is retained.

If images are silently missing from your output, `host_not_allowed` and `blocked_address` are the
two counters that tell you why.

## What this library will not do

Worth knowing before you design around it:

- **PDF fidelity is bounded, and unsupported features are dropped silently.** Charts, conditional
  formatting and some shape effects are omitted with no warning channel. The PDF is valid.
- **HTML → PDF goes through Word**, so CSS layout — flexbox, grid, floats, absolute positioning —
  does not survive. Text, headings, tables, lists, inline styling and images do.
- **DOCX → HTML returns a full document**, not a fragment. Extract the body with a parser.
- **Headers and footers are one line each.** Set on `PageSetup` — a single aligned line of text
  and page-number fields per header or footer. One running header and footer per document, plus
  an optional distinct first page — no per-section headers, and no odd/even (mirrored) variants.
- **One page setup per document.** No mixed portrait-and-landscape sections.

The full list, with the reasoning behind each, is in the
[README](https://github.com/Ank-KhoaHo/DocToolkit#known-limitations).
