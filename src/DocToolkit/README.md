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

## Measured against the alternatives

Every number below was measured on 2026-08-09 by adding the package to an empty `net8.0` console
app and building it. **Reproduce any row in under a minute** — that is the point of publishing the
method rather than a conclusion:

```bash
dotnet new console && dotnet add package <name> && dotnet build -c Release
find bin -path '*runtimes*' \( -name '*.so' -o -name '*.dylib' \) | wc -l
du -sm bin/Release/net8.0/runtimes
```

| package | native `.so`/`.dylib` in build output | `runtimes/` | licence in NuGet metadata |
|---|---|---|---|
| **Ank.DocToolkit** | **0** | **0 MB** | **`MIT`**, as an SPDX expression |
| EPPlus | 0 | 1 MB | ships `license.md` — read it |
| QuestPDF | 10 | 83 MB | ships `LICENSE.md` — read it |
| NPOI | 12 | 416 MB | ships `OSMFEULA.txt` — read it |
| ShapeCrawler | 19 | 664 MB | none declared |

**Two things this table deliberately does not do.** It does not tell you what those licences say —
they are linked from each package's own page and are the authors' to describe, not ours. And it does
not claim native payload is everyone's problem: **EPPlus carries essentially none**, so if that is
your only concern it is not a reason to switch. Where the payload does appear it comes from
SkiaSharp and Magick.NET, pulled in transitively for image rendering.

What the table is for is the case where **both** columns matter at once — a Linux container that has
to stay small, with a licence you can clear without asking anybody. That combination is the whole
reason this package exists, and all four constraints are re-checked by CI on every push.

## Offline by default — safe in air-gapped environments

**No method on DocToolkit's public API opens a network connection.** Not for images, not for
stylesheets, not for fonts, not for linked pictures or external workbook references. Once the
package is restored, DocToolkit never needs the network again. That default is unchanged and still
proven by 37 dedicated tests — see below.

There is exactly one way to change that, and you have to ask for it by name:

<!-- BEGIN SNIPPET: readme-remote-images -->

```csharp
// The ONLY API family that makes an outbound request: downloads and embeds images the markup
// names. It still succeeds in an air-gapped environment - a host that will not answer just leaves
// that image out of the result, after a per-image timeout, rather than failing the conversion.
byte[] docx = await HtmlToDocxConverter.ConvertAsync(html, allowRemoteImageDownload: true, ct);
byte[] pdf = await HtmlToPdfConverter.ConvertAsync(html, allowRemoteImageDownload: true, ct);

// RemoteImageOptions bounds that opt-in instead of leaving it wide open. Every default here is
// already the restrictive one, so `new RemoteImageOptions()` is far narrower than the bool form.
byte[] bounded = await HtmlToDocxConverter.ConvertAsync(html, new RemoteImageOptions(), ct);
```

<!-- END SNIPPET -->

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

**Trim-safe and native-AOT compatible**, and both claims are earned the same way: CI publishes a
probe application — trimmed in one job, native-AOT in the other — and *runs* it, asserting every
capability's result. Neither trimming nor AOT fails at publish time when it fails; it removes a
type something resolves by name and the app throws, or silently produces an empty document, in
production. One caveat that is a dependency's rather than ours: **ClosedXML emits `IL2090`** under
both, and will appear in your publish output.

📖 **[Guides](https://ank-khoaho.github.io/DocToolkit/guides/getting-started.html)** — getting
started, HTML conversion and page setup, Word templates, spreadsheets and presentations,
dependency injection, and running in production ·
🔎 **[API reference](https://ank-khoaho.github.io/DocToolkit/)**

## Usage

**This is one connected walkthrough of the whole surface, not a script that compiles as pasted** —
variables such as `logoBytes` and `chartPngBytes` stand in for data you already have.

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

// ...and back the other way, completing the round trip.
//
// Nothing here reaches the network or the disk: a remote image reference becomes a
// hyperlink rather than a fetch, a local file reference is refused, and data: images are
// inlined. That is asserted against a live loopback listener, with a positive control.
byte[] fromMarkdown = MarkdownToDocxConverter.Convert(markdown);

// Markdown -> PDF pivots through DOCX, the same way HTML -> PDF does, and inherits the
// offline guarantee above unchanged because it performs no conversion of its own.
byte[] markdownPdf = MarkdownToPdfConverter.Convert(markdown);

// Legacy Word 97-2003 binary .doc -> DOCX, for the files sitting on an old share drive.
// Reading is unconditional; converting REFUSES by default when the .doc holds pictures,
// drawings or form fields, because a .docx cannot carry them and a silently incomplete
// document is worse than an exception. In practice any .doc with a table is such a file.
string docText = DocToDocxConverter.ExtractText(legacyDoc);           // never refuses
byte[] converted = DocToDocxConverter.Convert(legacyDoc,
    new LegacyDocOptions { AllowContentLoss = true });                // opt in deliberately

// ConvertWithReport returns the same bytes and names exactly what was dropped.
ConversionResult<byte[]> docReport = DocToDocxConverter.ConvertWithReport(legacyDoc,
    new LegacyDocOptions { AllowContentLoss = true });

// XLSX -> CSV and XLSX -> HTML, one named sheet at a time.
//
// Cell text is CULTURE-INVARIANT in both - numbers use a dot, dates are ISO 8601 - which is
// a correctness requirement rather than a preference: a decimal comma on a German machine
// would collide with the CSV delimiter itself. Note this differs from WorkbookEditor.ReadSheet,
// which follows your culture because its result is data you inspect rather than a file you
// hand to something else. A formula exports its computed VALUE.
//
// The HTML is a <table> FRAGMENT, not a whole document (the opposite of DocxToHtmlConverter):
// a sheet is part of a page rather than a page. Every cell is escaped.
string sheetCsv  = XlsxToCsvConverter.Convert(xlsx, "Sales");
string sheetHtml = XlsxToHtmlConverter.Convert(xlsx, "Sales");

// Make a generated sheet readable: XlsxFormat.Report is a bold header row, that row frozen,
// and columns auto-fitted. Add a number format per column if you want one.
//
// This set is deliberately SMALL and closed. Cell styling is an open-ended surface - fonts,
// borders, fills, conditional rules - and this package's premise is a narrow one it can
// guarantee. If you need more, use ClosedXML directly rather than through a thinner API.
byte[] report = WorkbookEditor.Format(xlsx, "Sales",
    XlsxFormat.Report.WithNumberFormat("B", "#,##0.00"));

// ...and, if you need to know what those conversions could NOT carry across, the same
// call with a report. ConversionResult<T> gives you the output plus a ConversionWarning
// list; each warning carries a Code, a Message and a ConversionLossKind saying how bad
// it was (None, Approximation, Omission, Failure).
//
// The plain Convert overloads above are unchanged and return exactly the same output -
// this is opt-in. A conversion that loses something still SUCCEEDS: you get the document
// and the warnings, and decide for yourself.
ConversionResult<string> report = DocxToHtmlConverter.ConvertWithReport(docx);
if (report.HasLoss)
{
    foreach (ConversionWarning w in report.Warnings)
        Console.WriteLine($"{w.Kind} {w.Code}: {w.Message}");
}
string faithful = report.Value;

// Page size, orientation and margins. Generated documents are A4 with one-inch margins
// unless you say otherwise.
byte[] landscape = await HtmlToPdfConverter.ConvertAsync(
    html, PageSetup.A4.Landscape().WithMargins(36));

// Paper and remote images together. Before 0.18.0 these were mutually exclusive:
// (html, page) always converted offline and (html, options) always laid out on A4,
// so asking for both silently dropped one of them.
byte[] branded = await HtmlToDocxConverter.ConvertAsync(html, PageSetup.Letter, remoteImages);

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

// Swap a placeholder box for an image - see Images below for what "position and size come
// from the template" means, and when this is refused rather than done silently
byte[] withChart = PresentationEditor.ReplaceImage(editedPptx, "{{chart}}", chartPngBytes);

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

<!-- BEGIN SNIPPET: readme-fill-rows -->

```csharp
byte[] filled = DocxEditor.FillRows(docx, "item", new[]
{
    new Dictionary<string, string> { ["Desc"] = "Widget", ["Qty"] = "2", ["Total"] = "19.98" },
    new Dictionary<string, string> { ["Desc"] = "Gadget", ["Qty"] = "5", ["Total"] = "45.00" },
});

// then the document-level scalars
filled = DocxEditor.ReplaceText(filled, new Dictionary<string, string> { ["{{customer}}"] = "Contoso Ltd" });
```

<!-- END SNIPPET -->

Every clone keeps the template row's formatting, shading and borders, and a hyperlink inside a cell
survives with its target intact.

**Row keys are bare field names** (`Desc`), while `ReplaceText` keys are full placeholders
(`{{customer}}`) — the collection name is already an argument here, so repeating it in every key of
every record would duplicate it many times over.

A placeholder with no matching key becomes empty rather than staying visible. Placeholders for other
prefixes are untouched, so a second call fills a second table. An empty list removes the template
row, and removes the table if that row was its only one.

Read one back to verify what actually landed, rather than trusting the fill succeeded silently:

<!-- BEGIN SNIPPET: readme-read-table -->

```csharp
int tables = DocxEditor.TableCount(filled);
IReadOnlyList<IReadOnlyList<string>> rows = DocxEditor.ReadTable(filled, 0);
// rows[0] is the header row: ["Description", "Qty", "Total"]
// rows[1] is: ["Widget", "2", "19.98"]
```

<!-- END SNIPPET -->

The index is 0-based — deliberately unlike `PdfEditor.ExtractPages`'s 1-based `firstPage`, which
numbers pages the way a reader does; a table index has no such reader-facing numbering. And a row
comes back with the shape it actually has: a horizontally merged cell means that row genuinely
holds fewer cells than its neighbours, and padding it out to a rectangle would invent data that is
not in the document.

## Images

A text placeholder becomes an inline image — a logo, a signature, a QR code:

<!-- BEGIN SNIPPET: readme-replace-image -->

```csharp
byte[] withLogo = DocxEditor.ReplaceImage(docx, "{{logo}}", File.ReadAllBytes("logo.png"));

// or at a chosen width; the height scales to keep the aspect ratio
byte[] signed = DocxEditor.ReplaceImage(withLogo, "{{signature}}", sigBytes, widthPoints: 90);
```

<!-- END SNIPPET -->

**PNG and JPEG**, identified by their own magic bytes rather than a filename. Omit the size and the
image's intrinsic dimensions are read from its header at 96 DPI; give one dimension and the other
scales; give both and it is stretched to fit.

Works in the body, headers, footers, footnotes and endnotes — a logo usually belongs in a header,
and the image is attached to that header's own part so Word resolves it correctly.

Only the placeholder text is removed: `Signed: {{signature}} (authorised)` becomes `Signed: `, the
image, then ` (authorised)`, with the surrounding runs keeping their formatting.

### PowerPoint

`PresentationEditor.ReplaceImage` swaps a whole shape rather than splicing into text, because a
PPTX picture is a positioned shape and not something inline in a text flow the way a DOCX image
is - which is also why it takes no size argument, unlike `DocxEditor.ReplaceImage` above:

<!-- BEGIN SNIPPET: readme-pptx-replace-image -->

```csharp
byte[] filled = PresentationEditor.ReplaceImage(pptx, "{{chart}}", File.ReadAllBytes("chart.png"));
```

<!-- END SNIPPET -->

Position and size come from the template, so there is nothing to pass: a designer draws a box in
PowerPoint where the chart belongs, and the image lands there, scaled to fit and centred.

The shape's text must be nothing but the placeholder. The whole shape is replaced, so a shape
reading `Chart: {{chart}} (Q3)` would lose the words around the placeholder - silently, and in a
document that is still schema-valid. That is refused instead, with a
`DocumentConversionException` naming the shape's actual text.

**PNG and JPEG**, identified by their own magic bytes rather than a filename, same as the DOCX
form above.

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

## PDF utilities

Operations on a PDF that already exists — the only part of this library that **reads** one rather
than writing it. Nothing here re-renders: pages move between documents as they are, so text, fonts
and images arrive unchanged and the converters' fidelity caveats do not apply.

<!-- BEGIN SNIPPET: readme-pdf-utilities -->

```csharp
int pages = PdfEditor.PageCount(pdf);

// Join several into one, in the order given.
byte[] bundle = PdfEditor.Merge([cover, invoice, terms]);

// And take a range back out. firstPage is 1-based, the way a reader numbers pages.
byte[] justTheInvoice = PdfEditor.ExtractPages(bundle, firstPage: 2, count: 1);

// Or keep everything except a range - the complement of ExtractPages.
byte[] withoutTheCover = PdfEditor.RemovePages(bundle, firstPage: 1, count: 1);

// Turn a page that came out sideways. Relative, so calling it twice leaves you at 180.
byte[] upright = PdfEditor.RotatePages(bundle, firstPage: 3, count: 1, degrees: 90);

// Put the pages in a different order - a permutation of every page, not a subset.
byte[] resequenced = PdfEditor.ReorderPages(bundle, [3, 1, 2]);

// Slot another document in. atPage is where its first page lands; PageCount + 1 appends.
byte[] withAppendix = PdfEditor.InsertPages(bundle, appendix, atPage: 2);

// Read a PDF's text back out, one string per page — pageText[0] is page 1.
IReadOnlyList<string> pageText = PdfEditor.ExtractText(bundle);
```

<!-- END SNIPPET -->

A scanned PDF has no text layer, so every page comes back as `""` — that is the file's actual
content, not a failure to extract, and OCR is out of scope here.

A PDF needing a **password to open** raises `DocumentConversionException`. One carrying only owner
permission flags — "no copying" and the like — is still read. That is measured behaviour and is
what the rest of the ecosystem does, but it is worth knowing rather than discovering.

### Passwords and permissions

`PdfEditor.Protect` encrypts a PDF with a `PdfProtection`; `PdfEditor.Unprotect` takes the
encryption off again so the operations above can work on it.

<!-- BEGIN SNIPPET: readme-pdf-protection -->

```csharp
// A password to OPEN the document. Without it the file cannot be read at all.
byte[] locked = PdfEditor.Protect(statement, new PdfProtection
{
    UserPassword = "s3cret",
    AllowPrinting = false,
});

// An OWNER password leaves the document readable and asks readers to honour the
// restrictions. It is not a lock - use UserPassword when content must not be read.
byte[] restricted = PdfEditor.Protect(statement, new PdfProtection
{
    OwnerPassword = "admin",
    AllowCopying = false,
});

// The other PdfEditor operations refuse an encrypted document, so unprotect it first.
// If the document has an owner password, that is the one required here.
byte[] opened = PdfEditor.Unprotect(locked, "s3cret");
```

<!-- END SNIPPET -->

**The two passwords are not interchangeable, and this is the usual mistake.** A **user password** is
required to open the document and is enforced by cryptography. An **owner password** leaves the file
readable by anyone and merely *asks* a reader to honour the permission flags — a cooperative reader
greys out printing, an uncooperative one need not. If the content must not be read, set
`UserPassword`.

Two consequences that are measured rather than assumed:

- **`Unprotect` needs the OWNER password when the document has one**, even if you also know the user
  password — removing protection is a modification, and the PDF format reserves that for the owner.
- **Every permission defaults to allowed**, so adding a password does not silently stop a document
  being printed.

`PdfEncryptionStrength.Aes128` is the default because every reader in service can open it.
`Aes256` is stronger but needs a PDF 2.0 reader (Acrobat X and later), which is a compatibility
decision rather than a "more is better" one.

Document information — what a file manager shows in its properties panel, and what a search
indexer reads — is a `PdfMetadata`:

<!-- BEGIN SNIPPET: readme-pdf-metadata -->

```csharp
byte[] stamped = PdfEditor.WithMetadata(bundle, new PdfMetadata
{
    Title = "Invoice INV-2026-0042",
    Author = "Contoso Ltd",
});

PdfMetadata info = PdfEditor.ReadMetadata(stamped);
```

<!-- END SNIPPET -->

Every `PdfMetadata` property is nullable, and **`null` means absent rather than blank** in both
directions. Reading, that lets you tell "no title" from "a title deliberately set to empty";
writing, a `null` property leaves what the document already had alone, so stamping a title does not
silently erase the author.

`Stream` overloads exist for `PageCount`, `Merge`, `ExtractPages`, `RemovePages`, `RotatePages`, `ReorderPages`, `InsertPages` and `ExtractText` — that is, for every operation here. Unreadable input raises
`DocumentConversionException`, like everything else here.

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

## Migrating

### 0.20.x to 0.21.0 - `DocxEditor.ExtractText` now separates blocks

Before 0.21.0 `ExtractText` returned the document's raw concatenated text, with **no separator
between blocks at all**. A heading `Title` followed by a paragraph `Body text.` came back as the
single token `TitleBody text.`, and adjacent table cells `A` and `B` came back as `AB`. Word
boundaries were lost, so anything that tokenised, indexed or diffed the result got fused words.

From 0.21.0 blocks are separated by `\n` and the cells of a table row by `\t` - which is what
Word's own *save as plain text* writes, and what this method already did between the body and any
headers or footers.

```csharp
// 0.20.x  ->  "TitleBody text."
// 0.21.0  ->  "Title\nBody text."
string text = DocxEditor.ExtractText(docx);
```

Substring checks such as `text.Contains("Title")` are unaffected. **Exact-match comparisons against
the old fused output will need updating** - that is the whole of the breaking change. If you need
the previous shape, `text.Replace("\n", "").Replace("\t", "")` reproduces it.

### 0.12.x to 0.13.0 - generated documents are now A4

Before 0.13.0 a generated DOCX stated **no page size at all**, so Word applied its Normal template -
US Letter on a US install, A4 on most others - and a generated PDF was always US Letter. The same
content therefore printed on different paper depending on who opened it.

From 0.13.0 every producer states its page setup explicitly and **defaults to A4** with one-inch
margins. To keep the previous PDF behaviour, pass `PageSetup.Letter`:

<!-- BEGIN SNIPPET: readme-html-to-pdf-page -->

```csharp
byte[] pdf = await HtmlToPdfConverter.ConvertAsync(html, PageSetup.Letter);
byte[] docx = DocxEditor.Create(blocks, PageSetup.Letter);
```

<!-- END SNIPPET -->

The change was deliberate and is the point of that work rather than a side effect - but it shipped
filed under *Added* in the changelog, which understated it. It is recorded here because a published
version can be unlisted but never edited.

### 0.15.0 to 0.16.0 - the extensions package needs a newer DI abstraction

No behaviour change, but a **floor you may have to satisfy**.
`Ank.DocToolkit.Extensions.DependencyInjection` now requires
`Microsoft.Extensions.DependencyInjection.Abstractions` **8.0.2**, up from 8.0.0.

It follows from 0.15.0: PDFsharp raised `Microsoft.Extensions.Logging.Abstractions` from 6.0.0 to
8.0.3 in the core package's graph, and 8.0.3 requires DI abstractions >= 8.0.2. If your application
pins 8.0.0 or 8.0.1 you will see NuGet report a package downgrade rather than resolve silently -
raise your reference, or remove the pin and let it float.

The core package `Ank.DocToolkit` is unaffected.

### 0.13.0 to 0.14.0 - verified on macOS and arm64

No behaviour change. CI now runs the full suite on macOS and Linux arm64 as well as Linux x64 and
Windows, so "runs everywhere .NET does" is measured on each rather than inferred.

## Headers and footers

Attach them to the `PageSetup`, and every producer honours them:

<!-- BEGIN SNIPPET: readme-page-setup -->

```csharp
var page = PageSetup.A4
    .WithHeader(DocxHeader.Text("Contoso Ltd"))
    .WithFooter(DocxHeader.Of(HeaderAlignment.Right,
        DocxHeaderSegment.Text("Page "), DocxHeaderSegment.PageNumber,
        DocxHeaderSegment.Text(" of "), DocxHeaderSegment.PageCount));

byte[] docx = DocxEditor.Create(blocks, page);
```

<!-- END SNIPPET -->

The page number is a real field, so "Page 3 of 12" is right on every page rather than frozen at
the moment the document was generated.

## Known limitations

Things this package deliberately does not do, or does only partly — listed because the alternative
is that you find out by reading the source.

| Limitation | Detail |
|---|---|
| **PDF fidelity is bounded, and unsupported features drop silently** | Charts, conditional formatting and some shape effects are omitted rather than reported; the PDF converters have no warning channel, because the renderer beneath them produces no report to surface. The output is a valid PDF either way. The DOCX → HTML and DOCX → Markdown exporters are the exception — see `ConvertWithReport`. |
| **PDF fonts depend on the machine doing the conversion** | Where a system font is available it is **embedded**: on a Windows dev box the same invoice produces a ~167 KB PDF carrying Arial-Regular and Arial-Bold. In a slim container with no fonts installed, nothing is embedded and the PDF falls back to the **base-14 standard fonts** (Helvetica), giving ~1.5 KB. **Both are valid and both render**, and Arial and Helvetica are metric-compatible so line breaks do not move — but the glyphs are not identical, so a PDF built in your container will not be byte-identical to one built on your laptop. Install fonts in the image if you need a specific face. |
| **HTML → PDF goes through DOCX** | So fidelity is bounded by what HtmlToOpenXml maps into WordprocessingML, not by what a browser would render. Complex CSS layout — flexbox, grid, floats, absolute positioning — does not survive. Text, headings, tables, lists, inline styling and images do. |
| **No external stylesheets** | `<link rel="stylesheet">` is not fetched, by design. Inline `<style>` and `style=` are honoured. |
| **Headers and footers are one line each** | A header or footer is a single aligned line of text and page-number fields, set on `PageSetup`. One running header and footer per document, plus an optional distinct first page — per-section headers and odd/even (mirrored) variants are not supported. |
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

MIT. See `THIRD-PARTY-NOTICES.txt` for dependency attribution — in particular `SixLabors.Fonts`,
pinned to an exact version on its 1.x line because 1.x is the last Apache-2.0 licensed line of
that package. The notices file records the version actually resolved.
