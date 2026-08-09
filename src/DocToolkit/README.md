# DocToolkit

Generating Word, Excel, PowerPoint or PDF files from .NET usually means one of these:

- `System.Drawing.Common` throwing **`PlatformNotSupportedException`** the first time your container
  runs on Linux — it restored fine, it built fine, and it fails at runtime;
- a library whose licence is not permissive for commercial use — **EPPlus** is Polyform
  Noncommercial, **Syncfusion**'s community licence is revenue-gated, and **Spire** and
  **IronPDF** are commercial products;
- installing **LibreOffice** or a headless **Chromium** into your image to render a PDF, and owning
  a few hundred MB and a CVE feed for the privilege;
- or discovering the package downloads fonts or images at runtime, on a machine with no route out.

**DocToolkit exists because all four are avoidable.** Convert HTML to DOCX and PDF, render XLSX and
PPTX to PDF, turn DOCX into HTML or Markdown, and open/edit DOCX, XLSX and PPTX — from .NET, with:

| | |
|---|---|
| **Permissive licences only** | MIT / Apache-2.0. No revenue thresholds, no per-seat fees, nothing to read twice. |
| **No native binaries** | `dotnet restore` is the whole install. No browser, no LibreOffice, no Office interop. |
| **Runs everywhere .NET does** | The full suite runs in CI on Linux, Windows, macOS and **arm64** — measured on each, not inferred. |
| **No runtime network I/O** | Nothing opens a socket by default. Proven by 37 air-gap tests. |

All four are properties of the *resolved dependency graph*, so CI re-checks every one on every push.

## Offline by default — safe in air-gapped environments

**No method on DocToolkit's public API opens a network connection.** Not for images, not for
stylesheets, not for fonts, not for linked pictures or external workbook references. Once the
package is restored, DocToolkit never needs the network again. That default is unchanged and still
proven by 37 dedicated tests — see below.

There is exactly one way to change that, and you have to ask for it by name:

```csharp
// The ONLY API family that makes an outbound request: downloads and embeds images the markup
// names. It still succeeds in an air-gapped environment - a host that will not answer just leaves
// that image out of the result, after a per-image timeout, rather than failing the conversion.
byte[] docx = await HtmlToDocxConverter.ConvertAsync(html, allowRemoteImageDownload: true, ct);
byte[] pdf  = await HtmlToPdfConverter.ConvertAsync(html, allowRemoteImageDownload: true, ct);

// RemoteImageOptions bounds that opt-in instead of leaving it wide open. Every default here is
// already the restrictive one, so `new RemoteImageOptions()` is far narrower than the bool form.
byte[] bounded = await HtmlToDocxConverter.ConvertAsync(html, new RemoteImageOptions(), ct);
```

**The opt-in is now bounded, not just present.** Every fetch it makes is subject to fixed limits:

- **http and https only** — never `file://`, which would otherwise read the host's own disk.
- **No redirects are followed** — each hop would need re-validating against the whole policy below;
  an unvalidated hop is the standard way past an address check like this one.
- **Loopback, private and link-local addresses are blocked by default**, including
  `169.254.169.254` (the cloud metadata endpoint), unless
  `RemoteImageOptions.AllowPrivateAddresses` is set `true`.
- **A 10-second timeout and a 5 MB cap per image**, both configurable, and the cap is enforced by
  counting bytes actually read off the stream — never by trusting a `Content-Length` header, which
  a hostile server can understate.

**Those limits are per image, not per document.** A document naming many remote images has no
aggregate ceiling of its own: at the defaults, peak memory lands near 240 MB whatever the image
count (fetches run concurrently, and buffering one costs roughly three times the cap), and images
on hosts that never answer cost about ten seconds each, several at a time. Neither is unbounded —
your `CancellationToken` is honoured throughout, and `Timeout` and `MaxBytesPerImage` are yours to
lower — but if you convert documents of unknown size, bound them with a deadline rather than
assuming the per-image caps do it for you.

**This is not a complete SSRF defence.** A host's address is resolved and checked, then resolved
again by the HTTP stack when it actually connects — a DNS answer that changes between those two
moments defeats the check. It stops the ordinary cases (a literal metadata address, a hard-coded
internal hostname) and raises the cost of the rest. A service that converts genuinely untrusted
HTML should also be egress-filtered at the network layer, not rely on this alone.

Everything else — `ConvertAsync(html)`, `ConvertToFileAsync`, `DocxToPdfConverter`, `DocxEditor`,
`WorkbookEditor`, `PresentationEditor` — is offline, unconditionally.

This is enforced, not merely intended. The test suite starts a real TCP listener on loopback,
feeds every public API markup that names it as an `<img src>`, a `<link rel="stylesheet">`, a CSS
`@import`, a `background-image`, an `<a href>`, an externally linked DOCX picture, an external
XLSX workbook link and more, and requires the accepted-connection count to be **exactly zero**. A
companion test points the same APIs at an unroutable address (TEST-NET-3) and requires them to
return promptly rather than stall on a connect timeout. A further suite proves the opt-in itself
over a real socket, not a mock: `file://` is refused even with downloads enabled, loopback is
refused by default and only reached with `AllowPrivateAddresses = true`, a host outside a non-empty
`AllowedHosts` is refused while one inside it is fetched, an oversized or slow response is aborted
rather than trusted, and invalid options are rejected by the converter itself.

`dotnet restore` is the one step that still needs a package feed. `THIRD-PARTY-NOTICES.txt` lists
the full dependency closure with resolved versions, so it can be mirrored onto an internal feed;
every entry is a plain managed assembly with no native payload and no post-restore download.

## Memory, and why there is no size limit

This library will not refuse a large document. It edits and converts documents; rejecting a big one
would be a defect, not a safeguard. What it will do is use memory proportional to the document's
*expanded* form, which is far larger than the file.

Measured 2026-08-08 on a 1.9 MB `.xlsx` of 40,000 rows x 8 columns:

| operation | peak managed heap held | relative to the file |
|---|---|---|
| `ReadSheet` | 120 MB | 64x |
| `SetCell` | 233 MB | 124x |
| `SetCellAsync` (`Stream`) | 238 MB | 127x |

The multiplier comes from the OOXML object model, not from copying bytes around: a spreadsheet cell
that occupies a few bytes compressed becomes a live object with a type, a style reference and a
parent chain. So **size a container from the expanded cost, not from the file size** — roughly two
orders of magnitude for a dense spreadsheet, less for text-heavy documents.

**The `Stream` overloads are not a memory optimisation**, and the table shows it: 238 MB against
233 MB for the same edit. They exist so a caller can hand over a source that is forward-only and
non-seekable, such as an HTTP request body, without materialising a `byte[]` first. If you need to
bound memory, bound concurrency — the per-call cost is what it is.

## Install

```bash
dotnet add package Ank.DocToolkit
```

Targets `net8.0` and `net10.0` — two LTS releases, one public API surface.

Verified in CI on Linux (x64 and **arm64**), Windows and macOS.

## Usage

```csharp
using DocToolkit;

// HTML -> DOCX
byte[] docx = await HtmlToDocxConverter.ConvertAsync("<h1>Invoice</h1><p>Total: 18,100.00</p>");

// HTML -> PDF (pivots through DOCX internally)
byte[] pdf = await HtmlToPdfConverter.ConvertAsync("<h1>Invoice</h1><p>Total: 18,100.00</p>");

// DOCX -> PDF
byte[] rendered = DocxToPdfConverter.Convert(docx);

// XLSX -> PDF and PPTX -> PDF. A deck renders one page per slide, at its own slide
// geometry rather than being letterboxed onto a paper size.
byte[] sheetPdf = XlsxToPdfConverter.Convert(xlsx);
byte[] deckPdf  = PptxToPdfConverter.Convert(pptx);

// DOCX -> HTML / Markdown, keeping the structure ExtractText throws away: a heading stays
// a heading, a table stays a table. Images come back as data: URIs, so the output is
// self-contained.
string html     = DocxToHtmlConverter.Convert(docx);
string markdown = DocxToMarkdownConverter.Convert(docx);

// Page size, orientation and margins. Generated documents are A4 with one-inch margins
// unless you say otherwise.
byte[] landscape = await HtmlToPdfConverter.ConvertAsync(
    html, PageSetup.A4.Landscape().WithMargins(36));

// Build a DOCX from data rather than markup - headings, paragraphs, tables and images.
// There is no HTML to escape here, so a value containing '<' cannot corrupt the
// document's structure, and the same blocks produce the same content on every machine.
var blocks = new[]
{
    DocxBlock.Heading("Quarterly Report", 1),
    DocxBlock.Paragraph("Revenue rose 12% against a flat cost base."),
    DocxBlock.Table(
        new[] { "Region", "Revenue" },
        new[] { new object[] { "EMEA", 1200 }, new object[] { "APAC", 980 } }),

    // altText becomes the drawing's descr - what a screen reader announces. Omit it for a
    // purely decorative image; omitted means the attribute is absent, not filled with a
    // placeholder that would be read out as though it described the picture.
    DocxBlock.Image(logoBytes, widthPoints: 120, altText: "Contoso logo"),
};
byte[] report = DocxEditor.Create(blocks);

// Straight to a stream or a file, without materialising the byte[]
await DocxEditor.CreateAsync(blocks, responseBody);
await DocxEditor.CreateToFileAsync(blocks, "report.docx");

// Fill a DOCX template - body, headers, footers, footnotes, endnotes and text boxes
byte[] filled = DocxEditor.ReplaceText(docx, new Dictionary<string, string>
{
    ["{{customer}}"] = "Contoso Ltd",
});
string text = DocxEditor.ExtractText(filled);                                 // body only
string all  = DocxEditor.ExtractText(filled, includeHeadersAndFooters: true);

// Spreadsheets
byte[] xlsx = WorkbookEditor.Create("Sales", new[]
{
    new object?[] { "Region", "Total" },
    new object?[] { "North", 1200 },
});
string cell = WorkbookEditor.ReadCell(xlsx, "Sales", "B2");
byte[] updated = WorkbookEditor.SetCell(xlsx, "Sales", "B2", 1500);

// Read a workbook you were handed, without knowing its shape in advance
IReadOnlyList<string> sheets = WorkbookEditor.SheetNames(xlsx);              // tab order, hidden included
IReadOnlyList<IReadOnlyList<string>> grid = WorkbookEditor.ReadSheet(xlsx, "Sales");
string topLeft = grid[0][0];    // anchored at A1, padded rectangular, blanks are ""

// Or build a workbook with several sheets, from data, one worksheet each - XlsxFormula is a
// cell value, usable here, in AppendRows rows below, and as SetCell's value argument
byte[] workbook = WorkbookEditor.Create(new[]
{
    XlsxSheet.Named("Sales",   new[] { new object?[] { "Region", "Total" }, new object?[] { "EMEA", 1200 } }),
    XlsxSheet.Named("Summary", new[] { new object?[] { "Grand total", XlsxFormula.From("SUM(Sales!B2:B2)") } }),
});

// Straight to a stream or a file, without materialising the byte[]
using var workbookStream = new MemoryStream();
await WorkbookEditor.CreateAsync(new[] { XlsxSheet.Named("Sales", rows) }, workbookStream);
await WorkbookEditor.CreateToFileAsync(new[] { XlsxSheet.Named("Sales", rows) }, "workbook.xlsx");

// Append rows after a sheet's last used row - every other sheet, and all existing formatting,
// is left as it was
byte[] appended = WorkbookEditor.AppendRows(xlsx, "Sales", new[] { new object?[] { "APAC", 980 } });

// Straight to a stream or a file, without materialising the byte[]
using var appendedStream = new MemoryStream();
await WorkbookEditor.AppendRowsAsync(xlsxStream, "Sales", new[] { new object?[] { "APAC", 980 } }, appendedStream);
await WorkbookEditor.AppendRowsAsync("workbook.xlsx", "workbook.xlsx", "Sales", new[] { new object?[] { "APAC", 980 } });

// Presentations
byte[] pptx = File.ReadAllBytes("deck.pptx");
int slides = PresentationEditor.SlideCount(pptx);
IReadOnlyList<string> slideText = PresentationEditor.ExtractText(pptx);   // in deck order
byte[] editedPptx = PresentationEditor.ReplaceText(pptx, new Dictionary<string, string>
{
    ["{{title}}"] = "Q3 Results",
});

// Or build a deck from data - titles and bullets, no template needed
var deckSlides = new[]
{
    PptxSlide.Titled("Q3 Results", "Revenue up 12%", "Costs flat"),
    PptxSlide.Titled("Outlook", "Hiring 3 engineers"),
};
byte[] deck = PresentationEditor.Create(deckSlides);

// Straight to a stream or a file, without materialising the byte[]
using var deckStream = new MemoryStream();
await PresentationEditor.CreateAsync(deckSlides, deckStream);
await PresentationEditor.CreateToFileAsync(deckSlides, "deck.pptx");

// Work directly with files - no ReadAllBytes/WriteAllBytes dance
var customer = new Dictionary<string, string> { ["{{customer}}"] = "Contoso Ltd" };
await DocxEditor.ReplaceTextAsync("invoice-template.docx", "invoice.docx", customer);
string invoiceText = await DocxEditor.ExtractTextAsync("invoice.docx");

// Input and output may be the same file
await DocxEditor.ReplaceTextAsync("invoice.docx", "invoice.docx", customer);
```

**Formulas carry no cached value.** A cell written with `XlsxFormula` holds the formula and nothing
else. Excel recalculates when it opens the file, and `ReadCell`/`ReadSheet` compute the value on
read — but a third-party reader that only reads cached values, such as openpyxl with
`data_only=True`, sees an empty cell until Excel has opened and saved the file. A formula that
cannot be evaluated reads back as its Excel error string (`#DIV/0!`, `#NAME?`, `#REF!`) rather than
throwing.

## Repeating table rows

A table row whose cells contain `{{item.Field}}` placeholders repeats once per record — invoice
line items, timesheet entries, order lines:

| Description | Qty | Total |
|---|---|---|
| `{{item.Desc}}` | `{{item.Qty}}` | `{{item.Total}}` |

```csharp
byte[] filled = DocxEditor.FillRows(docx, "item", new[]
{
    new Dictionary<string, string> { ["Desc"] = "Widget", ["Qty"] = "2", ["Total"] = "19.98" },
    new Dictionary<string, string> { ["Desc"] = "Gadget", ["Qty"] = "5", ["Total"] = "45.00" },
});

// then the document-level scalars
filled = DocxEditor.ReplaceText(filled, new() { ["{{customer}}"] = "Contoso Ltd" });
```

Every clone keeps the template row's formatting, shading and borders, and a hyperlink inside a cell
survives with its target intact.

**Row keys are bare field names** (`Desc`), while `ReplaceText` keys are full placeholders
(`{{customer}}`) — the collection name is already an argument here, so repeating it in every key of
every record would duplicate it many times over.

A placeholder with no matching key becomes empty rather than staying visible. Placeholders for other
prefixes are untouched, so a second call fills a second table. An empty list removes the template
row, and removes the table if that row was its only one.

## Images

A text placeholder becomes an inline image — a logo, a signature, a QR code:

```csharp
byte[] withLogo = DocxEditor.ReplaceImage(docx, "{{logo}}", File.ReadAllBytes("logo.png"));

// or at a chosen width; the height scales to keep the aspect ratio
byte[] signed = DocxEditor.ReplaceImage(withLogo, "{{signature}}", sigBytes, widthPoints: 90);
```

**PNG and JPEG**, identified by their own magic bytes rather than a filename. Omit the size and the
image's intrinsic dimensions are read from its header at 96 DPI; give one dimension and the other
scales; give both and it is stretched to fit.

Works in the body, headers, footers, footnotes and endnotes — a logo usually belongs in a header,
and the image is attached to that header's own part so Word resolves it correctly.

Only the placeholder text is removed: `Signed: {{signature}} (authorised)` becomes `Signed: `, the
image, then ` (authorised)`, with the surrounding runs keeping their formatting.

## Placeholder replacement

`DocxEditor.ReplaceText` and `PresentationEditor.ReplaceText` substitute against the
concatenated text of each paragraph, because Word and PowerPoint routinely split a single
visible word across several runs — a per-run `string.Replace` would miss `{{name}}` whenever
it straddles a boundary.

The result is spliced back into only the runs a match actually overlaps, so:

- runs outside a match keep their text **and their formatting**;
- hyperlinks and text boxes are left alone unless they contain a placeholder themselves;
- when a placeholder does straddle runs, the value lands in the run holding its first
  character and inherits that run's formatting.

Keys are matched in one left-to-right pass, longest key first at any given offset, so a
substituted value is never rescanned for further placeholders.

## How the no-network guarantee is built

`HtmlToOpenXml`, the HTML parser underneath, defaults to downloading every image it sees, and its
resource loader also speaks `file://` — so left alone it would give every caller of a
`byte[]`-in/`byte[]`-out API an SSRF reach, a read of the host's disk, and an unbounded hang.
DocToolkit shuts that off in two independent places on the default path:

1. **Image processing** is set to `EmbedDataUriOnly`. Only `data:` URI images are embedded;
   `http`, `https` and `file` sources are skipped.
2. **The resource loader is replaced** with one that supports no protocol and fetches nothing.
   The component capable of making a request is never constructed, so the guarantee does not rest
   on what a future release decides `EmbedDataUriOnly` means. It also keeps the default path away
   from HtmlToOpenXml 3.5.0's process-wide static `HttpClient`, which is not thread-safe.

Self-contained documents still convert in full: `data:` URI images are decoded by the parser and
never go through the loader.

The other converters and editors need no such handling — `DocumentFormat.OpenXml`, `ClosedXML`
and `OfficeIMO` do not resolve external relationships, external workbook links or remote fonts.
That is asserted, not assumed; see above.

## Known limitations

Things this package deliberately does not do, or does only partly — listed because the alternative
is that you find out by reading the source.

| Limitation | Detail |
|---|---|
| **PDF fidelity is bounded, and unsupported features drop silently** | Charts, conditional formatting and some shape effects are omitted rather than reported; there is no warning channel. The output is a valid PDF either way. |
| **HTML → PDF goes through DOCX** | So fidelity is bounded by what HtmlToOpenXml maps into WordprocessingML, not by what a browser would render. Complex CSS layout — flexbox, grid, floats, absolute positioning — does not survive. Text, headings, tables, lists, inline styling and images do. |
| **No external stylesheets** | `<link rel="stylesheet">` is not fetched, by design. Inline `<style>` and `style=` are honoured. |
| **No headers or footers on generated documents** | `DocxEditor.Create` and `HtmlToDocxConverter` produce a body. `ReplaceText` *does* reach into the headers and footers of a document you supply. |
| **One page setup per document** | `PageSetup` applies to the whole document; multiple sections with different paper is not supported. |
| **DOCX → HTML returns a full document, not a fragment** | Extract the body with a parser if you are embedding it. |
| **Formulas carry no cached value** | Excel recalculates on open and `ReadCell`/`ReadSheet` evaluate on read, but a reader that only reads cached values sees an empty cell until Excel has opened and saved the file. |
| **Memory scales with the document, not the file** | Peak is dominated by the OOXML object model — measured ~120 MB for `ReadSheet` and ~233 MB for `SetCell` on a 1.9 MB, 40,000-row workbook. See above; the `Stream` overloads are not cheaper. |
| **Below 1.0.0, permanently** | Anything may change in a minor version. |

## Telemetry

One `ActivitySource` and one `Meter`, both named `Ank.DocToolkit`:

```csharp
builder.Services.AddOpenTelemetry()
    .WithTracing(t => t.AddSource(DocToolkitTelemetry.ActivitySourceName))
    .WithMetrics(m => m.AddMeter(DocToolkitTelemetry.MeterName));
```

**Only the opt-in remote-image fetch is instrumented**, deliberately — it is the only place this
library touches the network, and a refused fetch is otherwise *silent*: the image is skipped and
your document still succeeds. On an air-gapped host that is every remote image, so without this
there was nothing to tell you an image never arrived, or why.

`doctoolkit.remote_image.fetches` counts attempts by outcome (`ok`, `scheme_refused`,
`host_not_allowed`, `blocked_address`, `http_error`, `too_large`, `failed`). **Only the host is
recorded, never the URL** — query strings carry tokens. It adds no packages and costs nothing when
nobody subscribes.

## Errors

Every public method reports failure as `DocumentConversionException`, with the underlying
library exception as `InnerException`. Bad arguments (null, empty, blank) still surface as
`ArgumentNullException`/`ArgumentException`, and a cancelled `CancellationToken` as
`OperationCanceledException`.

## Why HTML to PDF goes through DOCX

No permissively-licensed, NuGet-only library renders HTML to PDF on Linux: the only free
renderers are browsers, and a browser is a native binary. Pivoting through DOCX keeps the
whole chain pure managed.

## Licence

MIT. See `THIRD-PARTY-NOTICES.txt` for dependency attribution — in particular the pinned
`SixLabors.Fonts 1.0.0`, which is the last Apache-2.0 release of that package.
