# DocToolkit

[![CI](https://github.com/Ank-KhoaHo/DocToolkit/actions/workflows/ci.yml/badge.svg)](https://github.com/Ank-KhoaHo/DocToolkit/actions/workflows/ci.yml)
[![codecov](https://codecov.io/gh/Ank-KhoaHo/DocToolkit/branch/main/graph/badge.svg)](https://codecov.io/gh/Ank-KhoaHo/DocToolkit)
[![NuGet](https://img.shields.io/nuget/v/Ank.DocToolkit.svg?label=Ank.DocToolkit)](https://www.nuget.org/packages/Ank.DocToolkit/)
[![NuGet Downloads](https://img.shields.io/nuget/dt/Ank.DocToolkit.svg)](https://www.nuget.org/packages/Ank.DocToolkit/)
[![NuGet](https://img.shields.io/nuget/v/Ank.DocToolkit.Extensions.DependencyInjection.svg?label=Ank.DocToolkit.Extensions.DependencyInjection)](https://www.nuget.org/packages/Ank.DocToolkit.Extensions.DependencyInjection/)

Convert **HTML and Markdown → DOCX and PDF**, export **DOCX → HTML/Markdown** and
**XLSX → CSV/HTML**, read text out of a **PDF**, and open/edit **DOCX, XLSX and PPTX**, from .NET.

**Pure managed. No native binaries, no browser, no LibreOffice, no Office interop.**
Works after `dotnet restore` alone, runs on Linux, macOS, Windows and arm64, and makes
**no network calls at runtime**.

```bash
dotnet add package Ank.DocToolkit
```

Targets `net8.0` and `net10.0`. MIT licensed.

📖 [Guides](https://ank-khoaho.github.io/DocToolkit/guides/getting-started.html) ·
🔎 [API reference](https://ank-khoaho.github.io/DocToolkit/) · 🗺️ [Roadmap](ROADMAP.md) · 📦
[package README](src/DocToolkit/README.md) · 📝 [CHANGELOG](CHANGELOG.md)

> **Contributing?** Branch from `main` and open a pull request back into it — `main` itself cannot
> be pushed directly. See [CONTRIBUTING.md](CONTRIBUTING.md) for the build, the commit-message
> rules, and the four constraints that will get a pull request rejected.

## Why this exists

Most .NET document stacks fail at least one of these. This one satisfies all four:

| Constraint | How |
|---|---|
| **Free for commercial use** | 30 dependencies: 28 MIT, 2 Apache-2.0. No revenue thresholds, no per-seat fees. |
| **NuGet only** | No Chromium download, no LibreOffice install, no native binaries. |
| **Targets `net8.0` and `net10.0`** | Two LTS targets, one public API surface. `net9.0` is deliberately absent — a `net9.0` app already consumes the `net8.0` build, so it would add no reach. `netstandard2.0` is deliberately absent too: every dependency supports it, but the bounded-fetch guarantee on remote images cannot be expressed there, and `DateOnly`/`TimeOnly` would make the API differ per target. |
| **Runs everywhere .NET does** | The full suite runs in CI on Linux, Windows, macOS and **arm64** (`ubuntu-24.04` x64, `windows-latest`, `macos-latest` Apple Silicon, `ubuntu-24.04-arm`). Not inferred from "pure managed" - measured on each. |
| **Works offline** | No runtime network I/O. Proven by 40 air-gap tests. |

All four are properties of the *resolved dependency graph*, so a single upstream bump can break
them silently — which is why CI re-checks every one on every push. That has happened once already
(see [Design notes](#design-notes)).

### Trimming and native AOT

Both packages are marked `IsTrimmable`, and the core package is marked `IsAotCompatible`.

That claim is checked the same way as the four above — CI trim-publishes an application over the
real dependency graph, then **runs it** and asserts every capability still works. A trim failure
does not appear at publish time: the trimmer removes a type something looks up by name, and the app
throws, or quietly produces an empty document, in production.

One caveat that is a dependency's rather than ours: **ClosedXML emits a trim warning**
(`IL2090`, in `DescribedEnumParser<T>`). Spreadsheet reading and writing work correctly in the
trimmed app CI runs, but the warning will appear in your publish output.

**Native AOT is claimed, and earned the same way.** `IsAotCompatible` is a strictly stronger
promise than `IsTrimmable` — it additionally forbids runtime code generation — and this README said
it was *not* claimed for as long as that was true. It changed on 2026-08-14 when a second CI job
started native-AOT-publishing the same probe and running it. Measured on that run: DOCX create and
read-back, placeholder replacement, PPTX, XLSX, HTML → DOCX, DOCX → PDF and PDF → text all correct,
including PdfPig's CMap handling and the font work, which are the two most reflection-dependent
paths in the graph.

The ClosedXML `IL2090` caveat above applies to an AOT publish too, for the same reason and with the
same outcome.

### Verifying what you downloaded

Every published `.nupkg` carries a signed **build provenance attestation** naming the workflow, the
commit and the runner that produced it. To check a package really came from this repository's CI
rather than someone's laptop:

```bash
gh attestation verify Ank.DocToolkit.<version>.nupkg --repo Ank-KhoaHo/DocToolkit
```

Attestation is produced immediately before the push, so the bytes that are verified are the bytes
that were published.

**The packages are not code-signed.** Authenticode signing needs a code-signing certificate this
project does not hold; provenance attestation answers "did this come from that source" without one,
which is the question a consumer of an open-source package usually has.

### The SBOM

Each release attaches a **CycloneDX SBOM** per package to its
[GitHub Release](https://github.com/Ank-KhoaHo/DocToolkit/releases) —
`DocToolkit.cdx.json` and `DocToolkit.Extensions.DependencyInjection.cdx.json`.

It is the machine-readable counterpart to `THIRD-PARTY-NOTICES.txt`: that file is for a human
meeting a licence question, this is for a scanner answering "am I exposed to CVE-X" without
resolving the graph itself. The core package's BOM lists 30 packages: 28 MIT, 2 Apache-2.0 — the
same closure the constraints table states above, derived independently.

(Phrased that way on purpose. `gen-third-party-notices.py` verifies README package counts against
the resolved graph, but only where they are written as "N packages: N MIT, N Apache-2.0". This
sentence used to read "29 components, 28 MIT and one Apache-2.0", which escaped the pattern on all
three counts — so it stayed stale through the release that added PdfPig while the check reported
success. A claim worth checking is worth phrasing the way the checker reads.)

**The SBOMs are attested too**, by the same provenance step as the packages. An SBOM that is merely
attached to a release is a file anyone with write access could replace; attested, it carries proof
it came from this workflow and this commit, which is what makes it worth reading.

```bash
gh attestation verify DocToolkit.cdx.json --repo Ank-KhoaHo/DocToolkit
```

**Its content is reproducible; the file is not.** Generation is run with `--no-serial-number`, so
the random per-run UUID is gone and exactly one field varies between two runs over an identical
graph — `metadata.timestamp`. To check a published SBOM against a regenerated one, drop it:

```bash
dotnet tool restore
dotnet dotnet-CycloneDX src/DocToolkit/DocToolkit.csproj -o . -fn regenerated.cdx.json   --json --no-serial-number --exclude-dev

diff <(jq 'del(.metadata.timestamp)' DocToolkit.cdx.json)      <(jq 'del(.metadata.timestamp)' regenerated.cdx.json)
```

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

## Usage

**This is one connected walkthrough, not a script that compiles as pasted** — variables such as
`lineItems` and `logoBytes` stand in for data you already have.

```csharp
using DocToolkit;

byte[] docx = await HtmlToDocxConverter.ConvertAsync("<h1>Invoice</h1><p>Total: 18,100.00</p>");
byte[] pdf  = await HtmlToPdfConverter.ConvertAsync("<h1>Invoice</h1>");  // pivots through DOCX
byte[] rendered = DocxToPdfConverter.Convert(docx);

// Markdown in both directions, and on to PDF — which pivots through DOCX exactly as
// HTML -> PDF does.
string md        = DocxToMarkdownConverter.Convert(docx);
byte[] fromMd    = MarkdownToDocxConverter.Convert("# Invoice\n\nTotal: **18,100.00**\n");
byte[] mdPdf     = MarkdownToPdfConverter.Convert("# Invoice\n\nTotal: **18,100.00**\n");

// Word 97-2003 binary .doc, the format on the old share drive. Reading never refuses;
// converting refuses by default when the file holds pictures, drawings or form fields,
// which a .docx cannot carry — in practice, any .doc containing a table.
string legacyText = DocToDocxConverter.ExtractText(legacyDoc);
byte[] migrated   = DocToDocxConverter.Convert(legacyDoc,
                        new LegacyDocOptions { AllowContentLoss = true });

// A sheet as CSV or as an HTML table fragment. Cell text is culture-invariant in both —
// a decimal comma would collide with the CSV delimiter itself.
string csv       = XlsxToCsvConverter.Convert(xlsx, "Sales");
string tableHtml = XlsxToHtmlConverter.Convert(xlsx, "Sales");

// Or build a DOCX from data rather than markup — no HTML to escape, so a value
// containing '<' cannot corrupt the document's structure
byte[] report = DocxEditor.Create(new[]
{
    DocxBlock.Heading("Quarterly Report", 1),
    DocxBlock.Paragraph("Revenue rose 12% against a flat cost base."),
    DocxBlock.Table(
        new[] { "Region", "Revenue" },
        new[] { new object[] { "EMEA", 1200 }, new object[] { "APAC", 980 } }),
});

// Fill a template — handles placeholders split across runs, and headers/footers
byte[] filled = DocxEditor.ReplaceText(docx, new Dictionary<string, string> { ["{{customer}}"] = "Contoso Ltd" });

// Repeat a table row per record — a row holding {{item.Desc}} becomes one row per line item,
// each keeping the template row's formatting
byte[] invoice = DocxEditor.FillRows(filled, "item", lineItems);

// Drop an image into a placeholder — sized from its own header, PNG or JPEG
invoice = DocxEditor.ReplaceImage(invoice, "{{logo}}", logoBytes);

// Read a table back — what makes a filled template verifiable. The index is 0-based, and a row
// comes back with the shape it has: a merged cell means fewer cells, not padding invented to fill it.
int tables = DocxEditor.TableCount(report);
IReadOnlyList<IReadOnlyList<string>> rows = DocxEditor.ReadTable(report, 0);
// rows[0] is the header row: ["Region", "Revenue"]

// Or work directly with files — no ReadAllBytes/WriteAllBytes dance, and input/output may be the same file
await DocxEditor.ReplaceTextAsync("invoice.docx", "invoice.docx", new Dictionary<string, string> { ["{{customer}}"] = "Contoso Ltd" });

byte[] xlsx = WorkbookEditor.Create("Sales", new[] { new object?[] { "Region", "Total" } });

// Or build a multi-sheet workbook from data, with a formula cell computing across sheets
byte[] workbook = WorkbookEditor.Create(new[]
{
    XlsxSheet.Named("Sales", new[] { new object?[] { "Region", "Total" }, new object?[] { "EMEA", 1200 } }),
    XlsxSheet.Named("Summary", new[] { new object?[] { "Grand total", XlsxFormula.From("SUM(Sales!B2:B2)") } }),
});

// Read one back without knowing its shape in advance
IReadOnlyList<string> sheets = WorkbookEditor.SheetNames(xlsx);
IReadOnlyList<IReadOnlyList<string>> grid = WorkbookEditor.ReadSheet(xlsx, sheets[0]);
IReadOnlyList<string> slideText = PresentationEditor.ExtractText(pptx);

// Or build a deck from data - titles and bullets, no template needed
byte[] deck = PresentationEditor.Create(new[]
{
    PptxSlide.Titled("Q3 Results", "Revenue up 12%", "Costs flat"),
    PptxSlide.Titled("Outlook", "Hiring 3 engineers"),
});

// Swap a placeholder box for an image. Position and size come from the template - a designer
// draws the box in PowerPoint where the chart belongs and the image lands there, scaled to fit
// and centred. The box's text must be only the placeholder, or this is refused rather than done
// silently: "Chart: {{chart}} (Q3)" would lose the words around it.
byte[] withChart = PresentationEditor.ReplaceImage(pptx, "{{chart}}", pngBytes);
```

**Formulas carry no cached value.** A cell written with `XlsxFormula` holds the formula and nothing
else. Excel recalculates when it opens the file, and `ReadCell`/`ReadSheet` compute the value on
read — but a third-party reader that only reads cached values, such as openpyxl with
`data_only=True`, sees an empty cell until Excel has opened and saved the file. A formula that
cannot be evaluated reads back as its Excel error string (`#DIV/0!`, `#NAME?`, `#REF!`) rather than
throwing.

### Page setup

Generated documents are **A4 with one-inch margins**. Pass a `PageSetup` for anything else:

<!-- BEGIN SNIPPET: readme-page-setup-options -->

```csharp
byte[] pdf = await HtmlToPdfConverter.ConvertAsync(
    html,
    PageSetup.A4.Landscape().WithMargins(36));

byte[] docx = DocxEditor.Create(blocks, PageSetup.Letter);
```

<!-- END SNIPPET -->

`PageSetup` is immutable, measured in points, and offers `A4`, `Letter` and
`Custom(widthPoints, heightPoints)`, plus `Landscape()` and `WithMargins(…)`. Every producer —
`HtmlToDocxConverter`, `HtmlToPdfConverter` and `DocxEditor.Create` — takes one.

> **Changed in this release.** Documents produced before this stated **no page size at all**, so
> Word applied its Normal template — US Letter on a US install, A4 on most others — and the
> generated PDF was always US Letter regardless. The same content therefore printed on different
> paper depending on who opened it. Every producer now states its page setup explicitly, defaulting
> to A4. If you want the old PDF behaviour, pass `PageSetup.Letter`.

`DocxToPdfConverter` takes no `PageSetup`: it renders a document that already carries its own page
setup, and honours it.

Six static classes, each stateless and safe to call concurrently, with a `byte[]` overload, a
`Stream` overload and a file-path overload for every capability, wrapping failures in
`DocumentConversionException`. Full surface:
**[package README](src/DocToolkit/README.md)** · [API docs](https://ank-khoaho.github.io/DocToolkit/).

Page setup and remote images combine: `ConvertAsync(html, page, options)`. Before 0.18.0
they were mutually exclusive - `(html, page)` converted offline and `(html, options)` laid out
on A4 - so asking for both silently dropped one.

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

**A scanned PDF has no text layer**, so `ExtractText` returns an empty string per page for one —
that is what the file contains, not a failure, and OCR is out of scope.

A PDF that needs a **password to open** raises `DocumentConversionException`. One whose owner set
permission flags such as "no copying" is still read — measured, and standard across the
ecosystem, but stated here rather than left to surprise anyone.

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

## Telemetry

One `ActivitySource` and one `Meter`, both named `Ank.DocToolkit`:

```csharp
builder.Services.AddOpenTelemetry()
    .WithTracing(t => t.AddSource(DocToolkitTelemetry.ActivitySourceName))
    .WithMetrics(m => m.AddMeter(DocToolkitTelemetry.MeterName));
```

**Only the opt-in remote-image fetch is instrumented**, deliberately. Every other call is one
synchronous, in-process, stateless operation that throws a typed exception on failure — you can
time and log around it and learn everything a span would tell you.

The fetch path is different: it is the only place this library touches the network, the
allow/refuse decision happens deep inside HtmlToOpenXml's pipeline, and **a refused fetch is
silent** — the image is skipped and your document still succeeds. On an air-gapped host every
remote image lands there. Without telemetry there was nothing to tell you an image never arrived,
or why.

`doctoolkit.remote_image.fetches` counts attempts by outcome — `ok`, `scheme_refused`,
`host_not_allowed`, `blocked_address`, `http_error`, `too_large`, `failed` —
and `doctoolkit.remote_image.bytes` records the size of images that arrived.

**Only the host is ever recorded, never the URL.** A query string routinely carries a signed
token, and telemetry leaves the machine and is retained.

It costs nothing when nobody subscribes, and adds **no packages** — `ActivitySource` and `Meter`
are in the shared framework on both target frameworks.

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

Things this package deliberately does not do, or does only partly. Listed because the
alternative is that you find out by reading the source.

| Limitation | Detail |
|---|---|
| **DOCX → HTML returns a full document, not a fragment** | `DocxToHtmlConverter.Convert` emits `<html><head>…<body>`. There is no fragment mode: producing one would mean re-serialising the renderer's output, so if you are embedding the result, extract the body with an HTML parser. Both text converters embed images as `data:` URIs, so the output is self-contained. |
| **PDF fidelity is bounded, and unsupported features are dropped silently** | All four converters (`DocxToPdfConverter`, `HtmlToPdfConverter`, `XlsxToPdfConverter`, `PptxToPdfConverter`) render what the underlying engine can represent. Features it cannot — charts, conditional formatting, some shape effects — are omitted rather than reported: **the PDF converters have no warning channel**, because the renderer beneath them produces no report to surface. The output is a valid PDF either way. The DOCX → HTML and DOCX → Markdown exporters are the exception and do report — see `ConvertWithReport`. |
| **PDF fonts depend on the machine doing the conversion** | Where a system font is available it is **embedded**: on a Windows dev box the same invoice produces a ~167 KB PDF carrying Arial-Regular and Arial-Bold. In a slim container with no fonts installed, nothing is embedded and the PDF falls back to the **base-14 standard fonts** (Helvetica), giving ~1.5 KB. **Both are valid and both render**, and Arial and Helvetica are metric-compatible so line breaks do not move — but the glyphs are not identical, so a PDF built in your container will not be byte-identical to one built on your laptop. Install fonts in the image if you need a specific face. |
| **HTML → PDF goes through DOCX** | So PDF fidelity is bounded by what HtmlToOpenXml maps into WordprocessingML, not by what a browser would render. Complex CSS layout — flexbox, grid, floats, absolute positioning — does not survive. Text, headings, tables, lists, inline styling and images do. |
| **No external stylesheets** | `<link rel="stylesheet">` is not fetched, by design: nothing here opens a socket by default. Inline `<style>` and `style=` attributes are honoured. |
| **Headers and footers are one line each** | A header or footer is a single aligned line of text and page-number fields, set on `PageSetup`. One running header and footer per document, plus an optional distinct first page — per-section headers and odd/even (mirrored) variants are not supported. |
| **One page setup per document** | `PageSetup` applies to the whole document. Multiple sections with different paper is a real Word feature and is not supported. |
| **Formulas carry no cached value** | A cell written with `XlsxFormula` holds the formula only. Excel recalculates on open, and `ReadCell`/`ReadSheet` evaluate on read — but a reader that only reads cached values (openpyxl with `data_only=True`, say) sees an empty cell until Excel has opened and saved the file. |
| **Memory scales with the document, not the file** | Peak memory is dominated by the OOXML object model. Measured on a 1.9 MB `.xlsx` of 40,000 rows: ~120 MB for `ReadSheet`, ~233 MB for `SetCell`. The `Stream` overloads are **not** cheaper — they exist for forward-only sources, not to save memory. There is no input size limit; sizing the host is the caller's decision. |
| **Remote images are bounded per image, not per document** | With the opt-in enabled, `RemoteImageOptions` caps each fetch by time and bytes. A document naming many images has no aggregate ceiling; your own `CancellationToken` is the backstop. |
| **Below 1.0.0, permanently** | Anything may change in a minor version. See [CONTRIBUTING.md](CONTRIBUTING.md). |

## Dependency injection

`Ank.DocToolkit` needs no container. For ASP.NET Core or worker services, a thin companion package
adds fifteen injectable interfaces that delegate one-for-one to the static API — same conversion
logic, no duplication.

```bash
dotnet add package Ank.DocToolkit.Extensions.DependencyInjection
```

<!-- BEGIN SNIPPET: readme-di-registration -->

```csharp
services.AddDocToolkit();

// Or opt in to remote image download for HTML->DOCX/PDF. This still succeeds in an
// air-gapped environment - an unreachable host leaves that image out rather than failing
// the conversion.
services.AddDocToolkit(o => o.AllowRemoteImageDownload = true);
```

<!-- END SNIPPET -->

Both packages ship at the same version from the same tag. See the
[extension package's README](src/DocToolkit.Extensions.DependencyInjection/README.md).

## Offline by default

No default code path performs network I/O — enforced, not merely intended. 40 guard tests assert
**zero** socket connections across the whole public API, against markup naming a loopback listener
sixteen ways (`<img src>`, `srcset`, `<link rel=stylesheet>`, `@import`, `background-image`,
`<iframe>`, `<object>`, `<script>`). The guard is proved by mutation: enabling downloads turns nine
of those tests red with real request lines, so it discriminates rather than passing vacuously.

`MarkdownToDocxConverter` takes the same stance and is measured the same way: an image URL in
Markdown becomes a **hyperlink rather than a fetch**, a local file reference is **refused** (a path
in untrusted content would let the document choose which file lands in the output), and `data:`
images are inlined because they carry their own bytes. Its zero-connection test ships with a
positive control, so it cannot pass because the probe stopped working.
**That proof is re-earned weekly rather than asserted** — Stryker.NET mutates the guard-critical
files in CI and fails if the mutation score drops (74.3% when the gate was added).

The one exception is explicit and opt-in, and **is silently image-less in an air-gapped
environment** — an unreachable host is skipped, not fatal:

<!-- BEGIN SNIPPET: readme-remote-opt-in -->

```csharp
byte[] docx = await HtmlToDocxConverter.ConvertAsync(html, allowRemoteImageDownload: true);

// Bounded instead of wide open: timeout, byte cap, host allow-list and a block on
// loopback/private/link-local addresses, all on by default. Not a complete SSRF defence — see
// the package README.
byte[] bounded = await HtmlToDocxConverter.ConvertAsync(html, new RemoteImageOptions());
```

<!-- END SNIPPET -->

## Reporting a vulnerability

Privately, through GitHub — [open an advisory](https://github.com/Ank-KhoaHo/DocToolkit/security/advisories/new),
not a public issue. [`SECURITY.md`](SECURITY.md) says what is in scope, and lists the limits that
are already documented and therefore aren't findings on their own — chief among them that the
remote-image guard above is **not a complete SSRF defence**.

## Design notes

**HTML → PDF goes through DOCX.** No permissively-licensed, NuGet-only, Linux-safe library renders
HTML to PDF directly — the only free renderers *are* browsers, and a browser is a native binary.
Pivoting through DOCX keeps the whole chain pure managed.

**ShapeCrawler was removed.** PPTX originally used it, until it turned out to pull SkiaSharp and
Magick.NET: 38 native `.so`/`.dylib` files, 664 MB of `runtimes/`, and 26 CVE advisories. PPTX now
sits directly on `DocumentFormat.OpenXml`. The lesson — checking a library's *API* tells you
nothing about what it drags in — is why the CI guard exists.

**`SixLabors.Fonts` is pinned to an exact 1.0.x version** (currently `[1.0.1]`). Version 2.x switches to the Six Labors Split License,
Apache-2.0 only under $1M annual revenue. CI asserts the pin holds, so a feed carrying only 2.x
fails restore loudly rather than silently relicensing you.

## Dependencies

Direct: `DocumentFormat.OpenXml` · `HtmlToOpenXml.dll` · `OfficeIMO.Word.Pdf` · `ClosedXML` ·
`SixLabors.Fonts [1.0.1]`. Full closure is 30 packages — 28 MIT, 2 Apache-2.0; see
[`THIRD-PARTY-NOTICES.txt`](src/DocToolkit/THIRD-PARTY-NOTICES.txt).

**Mirroring to a private feed?** Four things catch people out: `System.IO.Packaging` resolves to
**two** versions (8.0.1 for net8.0, 10.0.2 for net10.0); `Microsoft.Extensions.Logging.Abstractions`
is pinned at **6.0.0** by OfficeIMO; the package is **`RBush.Signed`**, not `RBush`; and
`SixLabors.Fonts` must be mirrored at exactly 1.0.0.

## Build and test

```bash
dotnet build DocToolkit.sln -c Release
dotnet test  DocToolkit.sln -c Release      # 426 tests x 2 target frameworks = 852 results

docker build -f Dockerfile.linux-test -t doctoolkit-linux-test .   # verify Linux locally
docker run --rm doctoolkit-linux-test

# samples reference the *published* packages, not this source - the restore a consumer gets
dotnet run --project samples/HtmlConversion      # one folder per capability - see samples/README.md
dotnet run --project samples/MinimalApi
```

Both packages ship at one version, from a single tag.
[release-please](https://github.com/googleapis/release-please) keeps a Release PR up to date as
commits land on `main`; merging it is a deliberate, manual decision, and
[`release.yml`](.github/workflows/release.yml) then publishes, authenticated with
[Trusted Publishing](https://learn.microsoft.com/en-us/nuget/nuget-org/trusted-publishing) — no
stored API key. Releases are cut by a maintainer merging a release pull request by hand;
nothing else publishes.

## Repository layout

```
src/DocToolkit/                                         the library
src/DocToolkit.Extensions.DependencyInjection/          DI extensions package
tests/                                                  the public-API approval guard, Stream-overload proofs, and the air-gap/dependency guards
samples/                                                twelve runnable samples, each answering one question, on the published packages
docfx/                                                  API docs source, published to GitHub Pages on release
```

## Licence

MIT — see [LICENSE](LICENSE).
