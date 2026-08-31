# DocToolkit

![DocToolkit - convert HTML to PDF and DOCX in C#, no browser, no native binaries](https://raw.githubusercontent.com/Ank-KhoaHo/DocToolkit/main/assets/banner.png)

[![NuGet](https://img.shields.io/nuget/v/Ank.DocToolkit.svg)](https://www.nuget.org/packages/Ank.DocToolkit/)
[![NuGet Downloads](https://img.shields.io/nuget/dt/Ank.DocToolkit.svg)](https://www.nuget.org/packages/Ank.DocToolkit/)
[![License: MIT](https://img.shields.io/badge/licence-MIT-blue.svg)](https://github.com/Ank-KhoaHo/DocToolkit/blob/main/LICENSE)

Generating Word, Excel, PowerPoint or PDF files from .NET usually means one of these:

- `System.Drawing.Common` throwing **`PlatformNotSupportedException`** the first time your container
  runs on Linux — it restored fine, it built fine, and it fails at runtime;
- a library whose licence is not permissive for commercial use — **EPPlus** is Polyform
  Noncommercial, **Syncfusion**'s community licence is revenue-gated, and **Spire** and
  **IronPDF** are commercial products;
- installing **LibreOffice** or a headless **Chromium** into your image to render a PDF, and owning
  a few hundred MB and a CVE feed for the privilege;
- or discovering the package downloads fonts or images at runtime, on a machine with no route out.

**DocToolkit exists because all four are avoidable.**

<!-- BEGIN GENERATED (scripts/gen-capability-matrix.py) - do not edit by hand -->

| From ↓ / To → | CSV | DOC | DOCX | HTML | Markdown | PDF | PPTX | XLSX |
|---|---|---|---|---|---|---|---|---|
| **CSV** | — | · | · | · | · | · | · | · |
| **DOC** | · | — | **✅** | · | · | · | · | · |
| **DOCX** | · | · | — | **✅** | **✅** | **✅** | · | · |
| **HTML** | · | · | **✅** | — | · | **✅** | · | · |
| **Markdown** | · | · | **✅** | · | — | **✅** | · | · |
| **PDF** | · | · | · | · | · | — | · | · |
| **PPTX** | · | · | · | · | · | **✅** | — | · |
| **XLSX** | **✅** | · | · | **✅** | · | **✅** | · | — |

A **✅** is a converter that ships; **·** is a pair with no converter, not a promise about one. Read a row as "from this format, into these".

<!-- END GENERATED -->

Plus PDF text extraction, open/edit for DOCX, XLSX and PPTX, and password protection for all
four — from .NET, with:

| | |
|---|---|
| **Permissive licences only** | MIT / Apache-2.0. No revenue thresholds, no per-seat fees, nothing to read twice. |
| **No native binaries** | `dotnet restore` is the whole install. No browser, no LibreOffice, no Office interop. |
| **Runs everywhere .NET does** | The full suite runs in CI on Linux, Windows, macOS and **arm64** — measured on each, not inferred. |
| **No runtime network I/O** | Nothing opens a socket by default. Proven by an air-gap suite that points markup at a loopback listener sixteen ways and asserts **zero** connections — including tests that assert a fetch *does* happen, so the zero can never pass vacuously. |

All four are properties of the *resolved dependency graph*, so CI re-checks every one on every push.

Install and usage come first below, deliberately. The measurements behind those four claims are at
the end of this page rather than the top — corpus pass rates on files the library has never seen,
the comparison against the alternatives, how the offline guarantee is built, and what conversions
actually cost in memory.

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

// Read and edit an existing Markdown document directly — front matter, headings and tables,
// or update one section's content in place — without converting to another format first.
IReadOnlyDictionary<string, object> meta = MarkdownEditor.ReadFrontMatter(markdown);

MarkdownHeading? changed = MarkdownEditor.FindHeading(markdown, "Changed");

string updatedDoc = MarkdownEditor.ReplaceSection(
    markdown, "Changed", "\n- fixed a bug\n- fixed another\n\n");

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
// and columns auto-fitted. Everything else is a With... call on top of it.
//
// The boundary is a CLOSED vocabulary rather than a small one: six rule conditions, five
// validation kinds, four highlights, a freeze position, a column width - each enumerable,
// measured and guaranteed. If what you need cannot be expressed as a closed set (arbitrary
// fonts, borders, fills, colour scales), use ClosedXML directly rather than a thinner API.
byte[] report = WorkbookEditor.Format(xlsx, "Sales", XlsxFormat.Report
    .WithNumberFormat("B", "#,##0.00")
    .WithColumnWidth("A", 42)                    // explicit; beats auto-fit for this column
    .WithFreezeAt(row: 2, column: 1)             // an XlsxFreeze position, not just the header
    .WithAutoFilter()
    // XlsxRuleKind: GreaterThan, LessThan, Between, EqualTo, Contains, Blank.
    // XlsxHighlight names an INTENT - Red, Amber, Green, Grey - never a colour, because a
    // colour picker cannot be enumerated and the moment one exists the boundary is gone.
    .WithRule(XlsxRule.GreaterThan("B2:B999", 10000, XlsxHighlight.Red))
    // XlsxValidationKind: WholeNumber, Decimal, TextLength, Date, List. This is the half of a
    // generated workbook that survives a human editing it.
    .WithValidation(XlsxValidation.OneOf("C2:C999", "Free", "Pro", "Team")));

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

// A chart, sharing one ChartType/ChartData model with PresentationEditor.AddChart below
var chart = WorkbookEditor.AddChart(
    xlsx, "Sales", "D2", ChartType.ColumnClustered,
    new ChartData(new[] { "North", "South" }, new[] { new ChartSeries("Total", new double[] { 1200, 980 }) }),
    title: "Regional Totals");

// A pivot table, aggregating existing sheet data. Its result grid is populated only when
// Excel opens the file - see Known limitations below for what that means for ReadCell and
// XlsxToPdfConverter.
var withPivot = WorkbookEditor.AddPivotTable(
    xlsx, "Sales", "A1:C10", "E1", "RegionSummary",
    rowFields: new[] { "Region" },
    dataFields: new[] { new PivotDataField("Amount", PivotFunction.Sum) });

// Presentations
byte[] pptx = File.ReadAllBytes("deck.pptx");
int slides = PresentationEditor.SlideCount(pptx);
IReadOnlyList<string> slideText = PresentationEditor.ExtractText(pptx);   // in deck order

// A SmartArt diagram's text lives in a different OOXML construct entirely, so ExtractText
// reports it separately from ordinary shape text - one entry per diagram, in slide order
IReadOnlyList<IReadOnlyList<string>> diagrams = PresentationEditor.ReadSmartArt(pptx, index: 1);

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

// A chart on slide 1, the same ChartType/ChartData model WorkbookEditor.AddChart uses above
byte[] deckWithChart = PresentationEditor.AddChart(
    deck, 1, ChartType.ColumnClustered,
    new ChartData(new[] { "North", "South" }, new[] { new ChartSeries("Total", new double[] { 1200, 980 }) }),
    title: "Regional Totals");

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

Two operations exist for anything that is not Excel and not this package's own readers:

```csharp
// Write the computed value into the file itself, not just into memory for this call.
byte[] withCachedValues = WorkbookEditor.EvaluateFormulas(xlsx);

// Ask first, rather than trust a value the engine may not actually understand.
XlsxFormulaInspection inspection = WorkbookEditor.InspectFormulas(xlsx);
if (!inspection.AllSupported)
{
    IEnumerable<XlsxFormulaCell> unsupported = inspection.Formulas.Where(f => !f.IsSupported);
    // Each XlsxFormulaCell's UnsupportedReason names the specific function or construct.
}
```

`EvaluateFormulas` is what `XlsxToPdfConverter` calls internally before rendering, so a formula
cell shows its value rather than its own source text in the PDF. A formula `InspectFormulas` marks
unsupported is left exactly as it was by `EvaluateFormulas` — no plausible-looking value is
invented for one the engine does not understand.

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

## Footnotes, endnotes and a table of contents

`DocxEditor.AddFootnote` and `DocxEditor.AddEndnote` insert a footnote or endnote reference at
every occurrence of a placeholder — the same placeholder shape `ReplaceImage` uses.
`DocxEditor.AddTableOfContents` replaces a placeholder **paragraph** instead: unlike a footnote or
an image, a table of contents is whole paragraphs rather than something that splices into a run,
so the placeholder must be that paragraph's entire content.

```csharp
byte[] withFootnote = DocxEditor.AddFootnote(docx, "{{note}}", "See the appendix for detail.");
byte[] withEndnote  = DocxEditor.AddEndnote(withFootnote, "{{cite}}", "Source: internal audit, 2026.");

// The placeholder paragraph must contain nothing else -- inserting a table of contents replaces
// the whole paragraph, which cannot preserve neighbouring text the way an inline splice can.
byte[] withToc = DocxEditor.AddTableOfContents(withEndnote, "{{toc}}", minLevel: 1, maxLevel: 3);
```

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

DOCX, XLSX and PPTX carry the identical document-properties concept — one shared `DocumentMetadata`,
read and written by `DocxEditor.ReadMetadata`/`WithMetadata`, `WorkbookEditor.ReadMetadata`/`WithMetadata`
and `PresentationEditor.ReadMetadata`/`WithMetadata`:

<!-- BEGIN SNIPPET: readme-document-metadata -->

```csharp
byte[] stamped = DocxEditor.WithMetadata(docx, new DocumentMetadata
{
    Title = "Q3 board report",
    Creator = "Contoso Ltd",
});

DocumentMetadata info = DocxEditor.ReadMetadata(stamped);
```

<!-- END SNIPPET -->

Same `null`-means-absent rule as `PdfMetadata` above. **`Creator` is not `PdfMetadata.Author` under
another name** — OOXML's `Creator` names the person who wrote the document, PDF's own `Creator`
names the application that produced the file, and the two ecosystems just happen to share the word.
One PPTX-specific caveat: `PresentationEditor.WithMetadata` can never leave `Creator` as `null` -
OfficeIMO's own save path stamps a default there whenever it is empty, even on a call that never
touches `Creator` at all. `ReadMetadata` alone, with no save involved, is unaffected.

## Word mail merge

Fill a template authored in Word — the kind carrying real `MERGEFIELD`s, showing as `«FirstName»` —
from named values.

```csharp
// What does this template ask for?
DocxMailMergeTemplate template = DocxMailMerge.InspectTemplate(docx);
Console.WriteLine(string.Join(", ", template.FieldNames));   // FirstName, Balance

byte[] letter = DocxMailMerge.Merge(docx, new Dictionary<string, string>
{
    ["FirstName"] = "Khoa",
    ["Balance"] = "1,204.55",
});
```

**This is not `DocxEditor.FillRows` under another name, and the difference is who authored the
template.** `DocxEditor` reads `{{placeholder}}` — plain text, typed by anyone in any editor, a
convention this library invented. Mail merge reads what Word itself writes. Neither substitutes for
the other: an existing Word mail-merge template has not one `{{` in it, and a caller without Word
cannot author merge fields.

**`Merge` refuses to produce a document that still has an unfilled field**, naming every one of
them. That is deliberate, and measured: an unfilled field survives as a live field and the document
reads `Your balance is «Balance»` — valid, opening cleanly, and looking finished. Nothing about it
says otherwise.

When you want the document anyway, ask for it by name:

```csharp
DocxMailMergeResult result = DocxMailMerge.MergeWithReport(docx, values);

foreach (DocxMailMergeField field in result.Report.Fields)
    Console.WriteLine($"{field.Name}: {field.Status}");     // a DocxMailMergeFieldStatus:
                                                            // Merged | MissingValue |
                                                            // UnsupportedFormatting | Malformed

if (!result.Report.IsComplete)
    Console.WriteLine($"unfilled: {string.Join(", ", result.Report.MissingFieldNames)}");
```

Worth knowing before you rely on it, all measured rather than assumed:

- **Both on-disk field encodings work** — the complex form Word writes, and the `w:fldSimple` form
  most generators emit.
- **Field names match case-insensitively**, so `firstname` fills `FirstName`.
- **A `null` value is refused.** The engine beneath merges it as an empty string and reports the
  document complete, so a database NULL would produce "Your balance is " with nothing flagging it.
  Pass `string.Empty` to mean "leave it blank" — that is accepted, because it is a decision.
- **Produced documents are flattened**: the merged fields become ordinary text, so re-opening the
  result in Word cannot re-merge it.
- **A document with no merge fields is not an error.** It comes back unchanged and reports itself
  complete — which is what `InspectTemplate` is for.
- `DocxMailMergeTemplate.IsValid` reports whether the *template* is sound, via
  `DocxMailMergeIssue` and `DocxMailMergeIssueKind`. It says nothing about whether a merge will be
  complete; that depends on the values you supply.

`Stream` overloads: `InspectTemplateAsync`, `MergeAsync` and `MergeWithReportAsync`. The last
returns a `DocxMailMergeReport` rather than a `DocxMailMergeResult`, because the document went to
your `destination`.

### Batch: one document per record

`MergeBatch` fills the same template once per record, yielding one document per entry:

```csharp
var records = new[]
{
    new Dictionary<string, string> { ["FirstName"] = "Alice", ["Balance"] = "100" },
    new Dictionary<string, string> { ["FirstName"] = "Bob",   ["Balance"] = "250" },
};

foreach (byte[] letter in DocxMailMerge.MergeBatch(template, records))
{
    // write each one out as it's produced -- memory stays proportional to one document
    // in flight, not the whole batch
}
```

It is lazy and strict, matching `Merge`'s own "refuse rather than produce an unfinished document"
rule: a record missing a required value throws, naming which record and which field, and nothing
after that record runs. `MergeBatchWithReport` is the lenient counterpart — it never throws, and
yields every record's document together with a `DocxMailMergeBatchItem` for each record, which holds
its `Document`, `RecordIndex` and `Report` (the same shape `MergeWithReport` returns for one document).

`MergeBatchToFiles`/`MergeBatchToFilesWithReport` write straight to disk instead of yielding bytes,
given a path for each record:

```csharp
IReadOnlyList<DocxMailMergeFileBatchItem> items = DocxMailMerge.MergeBatchToFilesWithReport(
    "template.docx", records, (index, record) => $"letter-{index}.docx");

foreach (var item in items)
    Console.WriteLine($"Record {item.RecordIndex} → {item.OutputPath}");
```

A `DocxMailMergeFileBatchItem` holds the `OutputPath`, `RecordIndex` and `Report` for each merge. Two
records producing the same output path are refused before anything is written — a caller's
`outputPathFactory` returning a duplicate is not silently allowed to overwrite one record's document
with another's.

### Conditional blocks, repeating blocks, and table rows

Beyond filling `MERGEFIELD`s, a Word mail-merge template can carry a conditional block
(`{{#Name}}` … `{{/Name}}`), a repeating block (`{{#each Name}}` … `{{/each Name}}`, flat or
nested), or a table row repeated by index:

```csharp
byte[] resolved = DocxMailMerge.MergeConditional(template, new Dictionary<string, bool>
{
    ["ShowDiscount"] = customer.HasDiscount,
});

byte[] expanded = DocxMailMerge.MergeRepeating(template, new Dictionary<string, IEnumerable<IReadOnlyDictionary<string, string>>>
{
    ["Items"] = order.Lines.Select(line => (IReadOnlyDictionary<string, string>)new Dictionary<string, string>
    {
        ["Sku"] = line.Sku,
        ["Qty"] = line.Quantity.ToString(),
    }),
});

byte[] rows = DocxMailMerge.MergeTableRows(template, tableIndex: 0, templateRowIndex: 1, records);

// A group header row and its detail rows are two row indexes on the same table, and each
// DocxMailMergeTableRowGroup carries one header's values plus the rows beneath it.
byte[] grouped = DocxMailMerge.MergeTableRowGroups(
    template, tableIndex: 0, groupTemplateRowIndex: 0, detailTemplateRowIndex: 1, groups);
```

A region nested inside another region is `MergeRepeatingRegions`, which takes a
`DocxMailMergeBlockData` per record so a record can carry its own inner regions. Each of the three
marker-based methods also has a `*WithReport` form, which always produces a document — a
`DocxMailMergeBlockResult` carrying the bytes and a `DocxMailMergeBlockReport` naming what went
unsupplied — instead of refusing. The one exception is a genuinely unbalanced marker structure,
which no supplied value can work around, so a `*WithReport` form throws for that too:

```csharp
var regions = new Dictionary<string, IEnumerable<DocxMailMergeBlockData>>
{
    ["Orders"] = customer.Orders.Select(o => new DocxMailMergeBlockData(
        new Dictionary<string, string> { ["Ref"] = o.Reference },
        new Dictionary<string, IEnumerable<DocxMailMergeBlockData>>
        {
            ["Lines"] = o.Lines.Select(line => new DocxMailMergeBlockData(
                new Dictionary<string, string> { ["Sku"] = line.Sku })),
        })),
};

DocxMailMergeBlockResult result = DocxMailMerge.MergeRepeatingRegionsWithReport(template, regions);
DocxMailMergeBlockReport report = result.Report;   // IsComplete, MissingNames, Issues
```

Run structural expansion (`MergeRepeating`/`MergeRepeatingRegions`/`MergeTableRows`/
`MergeTableRowGroups`) before `MergeConditional`, and both before `Merge`/`MergeWithReport`. The
order is a requirement, not a preference: a table index is positional, so resolving a conditional
block first can delete the table the index counted on, and a conditional block or merge field
inside a repeating region only exists once the repeating pass has expanded it.

## Fill-in forms: Word content controls

A content control is a named region Word itself protects - the format's own answer to a fill-in
field, and sturdier than a `{{placeholder}}` an author can break by editing inside it.

```csharp
DocxFormReport form = DocxForm.Inspect(docx);
foreach (DocxFormField field in form.Fields)
{
    // Read by Kind: a check box arrives as Checked, a date picker as Date, a picture as Bytes.
    // Text alone is blank for all three - a DocxFormValueKind says which to read.
    Console.WriteLine($"{field.Key}: {field.Value.Kind}");   // DocxFormValue, DocxFormValueKind
}

var values = new Dictionary<string, DocxFormValue>
{
    ["FullName"] = DocxFormValue.FromText("Khoa Ho"),
    ["Plan"] = DocxFormValue.FromChoice("Team"),
    ["Signed"] = DocxFormValue.FromChecked(true),
};

// Validate reports EVERY control with no value, so IsValid is false for any partial fill.
// Gate on the issues you actually care about, not on IsValid, unless you mean to supply all of them.
DocxFormValidation check = DocxForm.Validate(docx, values);  // DocxFormIssue, DocxFormIssueKind
bool blocked = check.Issues.Any(i => i.Kind != DocxFormIssueKind.MissingValue);

if (!blocked) docx = DocxForm.Fill(docx, values);
```

**This is a third template model, not a replacement for the other two.** `DocxEditor` fills
`{{placeholder}}` text and `DocxMailMerge` fills `MERGEFIELD` instructions; which one you need is
decided by whoever authored the document, not by preference.

**`Validate` checks keys and, for a typed control, values.** It reports which controls got no
value, which values matched nothing, which names are ambiguous - and whether a value fits: a
drop-down value outside its list, a non-date for a date picker and a non-boolean for a check box
each come back under their own `DocxFormIssueKind`. A **plain text** control validates anything,
because there is no constraint to check against, so a clean result means "nothing detectably wrong"
rather than "every value is right". `IsValid` means no issues of any kind; filter `Issues` by `Kind`
if you do not care about one of them.

**`Fill` is lenient about a MISSING value.** A control you supply no value for keeps its own
existing text - there is no injected marker - so filling half a form is a supported workflow.

**It is not lenient about a value that does not fit a typed control, and the three typed kinds do not
agree.** Measured: a drop-down value outside its list **throws**, while a non-date for a date picker
and a non-boolean for a check box are silently skipped and the control keeps its old content. That
asymmetry comes from the library underneath rather than a choice made here, and it is the strongest
reason to run `Validate` first - it reports all three the same way, before anything is written.

**A clean `Validate` is not a promise that `Fill` will succeed.** Nothing here decodes image bytes,
so content that is not a readable image validates clean and then throws from `Fill`.

**Images are supplied as bytes**, through `DocxFormValue.FromPicture(bytes, fileName)`. There is
deliberately no overload taking a path: this package does not read files you did not hand it.

**Keys** come from a control's tag or its alias, and `DocxFormKey` chooses. The default falls back
between them, so a template keyed either way works without you knowing which.

**`Fields` is what the document exposes under that key mode, not every control it contains.** Only
the first of several controls sharing a name appears, and a control with no name under the mode in
use does not appear at all - so a tag-only template read with `DocxFormKey.Alias` comes back **empty**,
which looks exactly like a document with no form. `Validate` reports both, as `DuplicateKey` and
`UnmappedControl`. **The order is not document order**, so sort by `Key` if the sequence matters.

**Only the document BODY is read or written.** A content control in a header or footer is invisible
to all three methods: it is not in the report, a value aimed at it comes back as `UnusedValue` - which
reads as though you invented the name - and `Fill` leaves it alone. `DocxMailMerge` **does** reach
headers, so if your form lives in one, that is the API to use.

`Stream` overloads: `InspectAsync`, `ValidateAsync`, `FillAsync`.

## Comments and tracked changes

`DocxReview` reads what a document carries from having been through review, and resolves it.

```csharp
DocxReviewReport review = DocxReview.Inspect(docx);

foreach (DocxComment comment in review.Comments)
    Console.WriteLine($"{comment.Author}: {comment.Text} ({comment.Replies.Count} replies)");

foreach (DocxRevision change in review.Revisions)
    Console.WriteLine($"{change.Kind} by {change.Author}: {change.AffectedText}");

if (review.UnresolvedThreadCount == 0)
    docx = DocxReview.AcceptRevisions(DocxReview.RemoveComments(docx));
```

`Inspect` returns comments and tracked changes together, because reading them separately can return
counts that disagree. A reply appears on its parent's `Replies` rather than as a comment of its own,
so `Comments.Count` is a thread count. `DocxRevision.Kind` is `Insertion`, `Deletion`, or `Other` —
Word records eleven kinds and the nine that describe formatting rather than content all arrive as
`Other` (`DocxRevisionKind`).

`RemoveComments`, `AcceptRevisions` and `RejectRevisions` each return a new document. Accepting keeps
insertions and drops deletions; rejecting does the reverse. **Neither can be undone from the result**
— keep the original if that matters. `Stream` overloads exist for all four operations.

This class reads and resolves review state; it cannot create any. There is no method here to add a
comment or to record an edit as a tracked change.

## Knowing before you convert: `DocxToPdfPreflight`

The DOCX → PDF renderer drops some things silently (see *Known limitations* below). If you convert
documents you did not author, this tells you which ones need a human to look at them.

```csharp
DocxToPdfPreflightReport preflight = DocxToPdfPreflight.Inspect(docx);

if (preflight.HasFindings)
    foreach (DocxToPdfPreflightFinding finding in preflight.Findings)
        Console.WriteLine($"{finding.Count}x {finding.Construct}: {finding.Message}");
```

**It reports what your document CONTAINS, not what was lost.** Nothing here converts anything, so
nothing here can claim to know what the renderer did — it answers *"is there anything in this file
worth a second look?"* That is the claim that stays true whatever a future renderer improves.

`DocxToPdfRisk.Known` means this project has converted a document carrying that construct and watched
the content fail to arrive, and each one has a test that **fails if the renderer ever starts carrying
it**. `Possible` would mean listed on reasoning alone; nothing is `Possible` today, deliberately.

**Three findings, and every one was measured rather than assumed:** unstyled footnote references
(one authored by Word, or by `DocxEditor.AddFootnote`, carries the character style that survives
render and is not reported), tables nested inside a table cell, and **content controls in a table**
— inside a cell, or wrapping a cell or a row. Text
boxes were measured to *survive* the render and are deliberately **not** reported, and neither is a
content control at **body level**, which also renders; a control in a table inside a text box is not
reported either, because that renders too. An inventory that fires on documents the renderer handles
perfectly
is one you learn to ignore. Charts, SmartArt and embedded objects are unmeasured, so they are absent
too rather than guessed at.

`Stream` overload: `InspectAsync`.

## Password-protected DOCX, XLSX and PPTX

Open one someone sent you, and produce one. `DocxEditor`, `WorkbookEditor` and `PresentationEditor`
each carry the same three members, with `Stream` overloads for both directions:

```csharp
byte[] locked = WorkbookEditor.Protect(xlsx, "s3cret");   // encrypt the whole file
byte[] opened = WorkbookEditor.Unprotect(locked, "s3cret");
bool needsPassword = WorkbookEditor.IsProtected(bytes);   // signature check, no password needed
```

**This is file encryption, not the "protect workbook / restrict editing" flag.** Office puts both
under the same menu and they are very different: this scrambles the whole file so nothing can be
read without the password, while the other kind is a request a reader may ignore. Only the first is
offered here.

**An encrypted Office file is not a package any more.** A plain `.docx`/`.xlsx`/`.pptx` is a ZIP; the
encrypted form is a compound file with the package sealed inside. So every other method on these
classes refuses one — that refusal is honest rather than awkward, because they genuinely cannot read
the content. Call `Unprotect` first. `IsProtected` answers "would they refuse this?" from the file's
first bytes, without a password.

A wrong password and a file that was never encrypted are reported as **different** failures, because
a caller can only act on one of them.

## Digital signatures

Inspect a document for OPC package signatures, and validate what they actually prove.
`DocxEditor`, `WorkbookEditor` and `PresentationEditor` each carry the same four members, with
`Stream` overloads:

```csharp
DocumentSignatureInfo info = WorkbookEditor.InspectSignatures(xlsx);
// info.HasSignatures, info.SignatureCount, info.Signers (CLAIMED identity - see below)

DocumentSignatureValidationReport report = WorkbookEditor.ValidateSignatures(xlsx);
// report.IsCryptographicallyValid   - was the signed content tampered with since signing?
// report.Signatures[0].CertificateChainStatus - does the signer's certificate chain to a
//                                                certificate this machine trusts?
```

`report.Signatures` is a list of `DocumentSignatureValidationResult`, one per signature, each
carrying its own `CryptographicStatus`, `CertificateChainStatus` and `RevocationStatus` — every
one a `DocumentSignatureStatus`.

**`report.IsCryptographicallyValid` is the tamper-detection verdict — the per-signature
`Signatures[0].CryptographicStatus`, despite the similar name, is not.** Measured directly: a
document altered after signing, without re-signing, still reports `CryptographicStatus = Passed`
on the affected signature — that field only confirms the signature block itself is well-formed,
not that the content it covers is unchanged. `IsCryptographicallyValid` is the field that
correctly goes `false` on a tampered document. For a document carrying more than one signature it
is an aggregate across all of them ("was anything altered"), not a per-signature answer.

**`InspectSignatures` reports a claimed identity, not a proven one.** `Signers` is read from each
signing certificate's own subject name, without validating the signature at all — anyone can put
any name on a self-signed certificate. Use `ValidateSignatures` before treating a signer's name as
real.

**Cryptographic integrity and certificate trust are independent findings, deliberately.**
Measured directly against a real self-signed certificate: an untampered signature reports
`IsCryptographicallyValid = true` even when the certificate never chains to a trusted root
(`CertificateChainStatus = Failed`, the ordinary and expected outcome for an internal/enterprise
signer). Set `ValidateCertificateTrust = false` on `DocumentSignatureValidationOptions` to skip
chain checking entirely and only ask "was this tampered with." **There is no option to trust an
internal certificate authority without installing it in this machine's own trust store** — an
earlier draft of this feature had one and it was removed before release, measured not to actually
confer trust (see `DocumentSignatureValidationOptions`'s own remarks for what was measured).

**No revocation checking, ever, and no network access of any kind.** Not configurable. Chain
validation checks only this machine's local trust store.

**An unsigned document and a tampered one can both report `IsCryptographicallyValid = false`.**
Check `HasSignatures` first — the two are different findings.

## Legacy binary Office files (.ppt, .xls, .doc)

**`.doc` and `.ppt` are read; `.xls` is not.** That split is measured rather than arbitrary, and it
is worth knowing before you plan around it.

| input | supported | how |
|---|---|---|
| **`.doc`** (Word 97-2003) | yes | `DocToDocxConverter` — read its text, or convert it to `.docx` |
| **`.ppt`** (PowerPoint 97-2003) | yes, **to PDF only** | `PptxToPdfConverter` accepts one directly |
| **`.xls`** (Excel 97-2003) | **no** | refused immediately; save it as `.xlsx` |

### Converting a `.doc` refuses more often than it succeeds

**Expect to pass `AllowContentLoss`.** Measured across 111 real `.doc` files from a public `.gov`
crawl:

| call | succeeded |
|---|---|
| `DocToDocxConverter.ExtractText` | **99 / 111 — 89%** |
| `DocToDocxConverter.Convert(doc, new LegacyDocOptions { AllowContentLoss = true })` | **99 / 111 — 89%** |
| `DocToDocxConverter.Convert(doc)` — the default | **12 / 111 — 11%** |

The default refuses whenever the source holds a payload a `.docx` cannot carry — pictures, drawings
or form fields, kept in a binary stream. On real documents that is **the common case, not the
exception**, which is worth knowing before you conclude the feature is broken.

**The refusal is deliberate**: it would rather fail than hand back a document quietly missing its
pictures. But if you are converting a share drive rather than one known file, the useful call is
`ConvertWithReport`, which returns the same bytes as the opt-in **and** tells you exactly what was
dropped:

```csharp
var result = DocToDocxConverter.ConvertWithReport(
    doc, new LegacyDocOptions { AllowContentLoss = true });

byte[] docx = result.Value;
foreach (var warning in result.Warnings)
    logger.LogInformation("{Code}: {Message}", warning.Code, warning.Message);
```

**Text, tables with every cell, and character formatting survive either way** — what is lost is the
unprojected binary payload, and nothing else. The ~11% that cannot be read at all are mostly files
older than the format this reads: two of the twelve were pre-97 Word binaries, not compound files.

**`.ppt` → PDF works but the editors do not accept `.ppt`.** `PresentationEditor` is OOXML-only, so
`SlideCount`, `ExtractText` and the rest still refuse a `.ppt`. Rendering it to PDF and reading the
PDF is the way round that.

**Not every `.ppt` converts.** Measured across real files from a public `.gov` crawl, roughly half to
three-quarters succeed; the rest fail with a stated reason rather than producing a damaged document.

**`.xls` is refused for cost, not capability.** The renderer underneath can read it, but measured on
real files a 101 KB workbook took 10.9 seconds, a 2.3 MB one did not finish in ten minutes, and a
7.7 MB one spent 161 seconds before failing anyway — while the supported `.xlsx` path renders 20,000
rows in under four. Accepting that on a path a caller can feed arbitrary uploads to is a cost nobody
chose, and no cheap bound exists because the work tracks content rendered rather than input size.

## Bulleted lists, and the one substitution this library makes

**A Word document containing a bulleted list renders to PDF, and to make that work one character is
substituted.** Word's default bullet is not `U+2022` — it is `U+F0B7`, a Symbol-font glyph in the
Unicode private-use area, stored in the document's numbering definitions. The PDF renderer cannot
encode it, and before this substitution such a document could not be rendered **at all**.

The marker becomes an ordinary `U+2022 BULLET` (or `U+00B7` for a square sub-bullet). Visually
near-identical; the alternative was no conversion.

**It applies only to list markers.** Document text is never altered, and a document with no lists is
returned untouched rather than repackaged.

### Non-Latin text depends on the fonts the machine has

**Rendering a document containing Cyrillic, Greek or CJK to PDF may fail, and whether it does is a
property of the host rather than of this library.** Measured 2026-08-17 on the same document: it
renders on Linux and macOS with its text intact, and is refused on Windows, because the fonts
available to the renderer differ. That is the same host-dependence that makes PDF output size vary
about a hundredfold — see *Fonts* under [running in
production](https://ank-khoaho.github.io/DocToolkit/guides/production.html).

When it is refused, it is refused **loudly**: a `DocumentConversionException` naming the character it
could not encode. It is never silently dropped, and that guarantee is the one this library actually
holds you to — a test asserts that either the conversion fails or the text is present in the PDF,
with no third outcome.

**Latin text is unaffected everywhere**, including smart quotes, em dashes and the rest of the
WinAnsi range. And **reading is never affected**: `DocxEditor.ExtractText` returns non-Latin text
correctly on every platform. It is only PDF rendering that depends on fonts.

**You can now take the machine out of the answer by supplying the font yourself.** `PdfFontOptions`
carries font bytes you already license, and the converters that render PDF accept it:

```csharp
var fonts = new PdfFontOptions("Noto Sans", File.ReadAllBytes("NotoSans-Regular.ttf"));

byte[] pdf     = DocxToPdfConverter.Convert(docx, fonts);
byte[] fromWeb = await HtmlToPdfConverter.ConvertAsync(html, fonts);
```

**For HTML → PDF, `HtmlToPdfOptions` carries fonts, page setup and the remote-image policy
together** — the overload above fixes the page at A4, so it is the one to reach for when a
conversion needs more than one of the three:

```csharp
byte[] letter = await HtmlToPdfConverter.ConvertAsync(html, new HtmlToPdfOptions
{
    Page  = PageSetup.Letter.WithMargins(36),
    Fonts = fonts,
});
```

Every property defaults to what the simpler overloads already do — A4, no fonts, and **no remote
fetching** — so setting only what you care about cannot quietly opt you into anything. There are
`Stream` and file-path forms too, matching every other capability.

Nothing is fetched and nothing is read from disk by this library — the bytes come from you, which is
why this works on an air-gapped host. **No font ships inside this package**, deliberately: one
covering Cyrillic, Greek and CJK is measured in megabytes against a package measured in tens of
kilobytes, and every consumer would pay for it to serve the few converting non-Latin text.

**Supply fonts covering everything your documents use, not just the script that failed.** The fonts
you pass **replace** the host's own fallbacks rather than adding to them, so supplying too few is
worse than supplying none. Measured over 99 real documents:

| fonts supplied | rendered |
|---|---|
| none | 71 / 99 |
| one (Arial) | **63 / 99** — fixed the 4 needing Cyrillic, broke 12 the host had been covering |
| four | **77 / 99** |

The refusal names the character it could not encode, which tells you what is still missing.

**It also changes how fonts are embedded generally**: measured on an ordinary Latin document, output
went from 128,755 bytes to 1,306. Both render correctly — the smaller one leans on the standard
fonts every PDF reader already has.

Otherwise, install fonts covering the script on the machine that does the rendering, and convert on
a host you control rather than assuming the developer machine's behaviour carries over.

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

### 0.45.0 - Pivot tables in XLSX

`WorkbookEditor.AddPivotTable` creates a pivot table from existing sheet data. **Its result grid
is empty until Excel opens and recalculates it** — nothing that writes the file, this method
included, computes a pivot aggregation. That is a harder version of the caveat `XlsxFormula`
already carries (see *Known limitations* below): a formula's value **is**
computed by `ReadCell`/`ReadSheet` on read, because this library's own engine evaluates it, while a
pivot table's is not. Reading the pivot's own cells back with `ReadCell`/`ReadSheet` immediately
after this call returns empty strings, and `XlsxToPdfConverter` renders nothing where the pivot's
results would be, for the identical reason it renders a formula's literal text rather than its
computed value. See the [spreadsheets and presentations
guide](https://ank-khoaho.github.io/DocToolkit/guides/editing/spreadsheets-and-presentations.html#pivot-tables)
for a worked example.

### 0.45.0 - Chart creation in XLSX and PPTX

`WorkbookEditor.AddChart` and `PresentationEditor.AddChart` create charts, sharing one
`ChartType`/`ChartData` model. DOCX chart creation is not included — see the guide. `ChartType`
has 15 values, not the full 17 OfficeIMO's own chart-kind enum offers — Scatter and Bubble are
excluded because both reject the shared categories-and-series `ChartData` shape on both Excel and
PowerPoint, measured directly rather than assumed.

### 0.45.0 - `PresentationEditor.ExtractText` now includes SmartArt

A SmartArt diagram's text lives in a diagram data part, not a text-bearing shape body, so it was
**invisible to `ExtractText`** before this release — a deck containing SmartArt reported fewer
entries than it actually held. `ReadSmartArt` reads the same diagrams directly, one entry per
diagram on a given slide.

**If you assert an exact `ExtractText(...).Count` against a SmartArt-bearing deck, expect a higher
number now** — one more entry per SmartArt diagram, appended after that slide's ordinary
text-bearing bodies. A deck with no SmartArt is completely unaffected. Nothing that creates a
document changed; this is a correctness fix to what a caller already shipping SmartArt-bearing
decks (authored in PowerPoint itself, or via `OfficeIMO.PowerPoint` directly) gets back.

### 0.43.0 - `DocxToPdfPreflight` no longer reports a normally-authored footnote

`Inspect`/`InspectAsync` **stop reporting a `Footnote` finding** for a document whose footnotes were
authored by Word, or by `DocxEditor.AddFootnote` — the finding used to fire on any footnote at all.
Measured: that content was never actually lost in this case, only the check was wrong.

**If you gate on `report.HasFindings` or filter `report.Findings` for `Footnote`**, a
footnote-bearing document that used to trigger it may no longer. Nothing about `AddFootnote`'s own
output changed, and no PDF this library produces is different — only the report is more accurate.
The finding still fires for a footnote reference built by hand, or produced by another tool, without
the character style Word always applies.

### 0.41.0 - table indexes count content-controlled tables

`DocxEditor.TableCount`, `ReadTable` and `FillRows` **now see a table, row or cell wrapped in a Word
content control** (`w:sdt`), matching `ExtractText`, which has read them since 0.38.0.

**`ReadTable(index)` can return a different table than it did.** The clearest case: where a wrapped
table came *before* an ordinary one, `ReadTable(0)` used to return the ordinary one - the table that
is physically second - because the wrapped one was invisible and the index slid past it. It returns
the first table now.

Three other results change, and each was previously wrong rather than merely different:

| document | before | now |
|---|---|---|
| only table wrapped in a control | `TableCount` **0**, `ReadTable(0)` threw | `1`, reads it |
| a wrapped **row** in an ordinary table | that row silently missing | present |
| a template row wrapped, or in a wrapped table | `FillRows` threw | expands |

**If you store table indexes**, re-derive them against 0.41.0 for any document that uses content
controls - which is most template-driven documents, since a control is usually the part that varies.
A document containing no content controls is completely unaffected.

The `FillRows` refusal is worth calling out separately, because its message was actively misleading:
it said the marker *"must appear inside a table cell"* when the marker was inside a table cell. That
message now only appears when it is true.

### 0.38.0 - `ExtractText` reads Word content controls

`DocxEditor.ExtractText` **returns text it previously omitted**. A document whose content sat inside
a Word content control (`w:sdt`) used to extract to nothing at all, while the same document rendered
to PDF carried that text perfectly.

If you index, search, diff or hash extracted text, **expect the output to change** for any document
containing a content control - which is most template-driven documents, since a content control is
usually the part that varies. Nothing about the API changed; the same call now answers correctly.

All six positions a control can occupy are covered: at body level, nested inside another control,
holding a table, inside a table cell, and controls wrapping a whole row or a single cell. The last
two also used to shift the `\t` positions of the cells beside them.

### 0.36.0 is broken — upgrade to 0.36.1, or stay on 0.35.0

**Do not use 0.36.0.** It was published carrying `DocToolkit.dll` alone; the six assemblies the
library was split into were missing from the package. Referencing it produces compile errors on
almost everything this library exposes:

```
error CS0012: The type 'PageSetup' is defined in an assembly that is not referenced.
              You must add a reference to assembly 'DocToolkit.Primitives'.
error CS0103: The name 'DocxEditor' does not exist in the current context
```

`0.36.1` is the same library with the packaging fixed. **Nothing in your code needs to change** —
this was never an API change, and 0.36.0 could not be compiled against at all, so no working code
depends on it.

A nuget.org version cannot be replaced or deleted, only unlisted, which is why this entry exists
rather than a corrected 0.36.0.

### 0.33.x to 0.33.4 - a file-path overload names the parameter you passed

Before 0.33.4, handing a **zero-byte file** to a file-path overload raised an `ArgumentException`
naming the parameter of the `byte[]` method underneath, not the one you called:

```csharp
// 0.33.3  ->  ParamName "docx",  message "DOCX content was empty."
// 0.33.4  ->  ParamName "path",  message "The file at '/data/empty.docx' is empty."
await DocxEditor.ExtractTextAsync(path);
```

A blank or null path was always reported correctly; only an existing but **empty** file was
affected. Nineteen overloads across `DocxEditor`, `WorkbookEditor`, `PresentationEditor`,
`PdfEditor`, `DocxToPdfConverter`, `PptxToPdfConverter` and `XlsxToPdfConverter` behaved this way.

The exception **type** is unchanged, so `catch (ArgumentException)` needs no edit. **Code that
switches on `ParamName`, or matches the message text, will need updating** - that is the whole of
the change. Each overload's documented `<exception>` tag already described the new behaviour; the
code disagreed with it.

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

### Choosing a page size

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
`Custom(widthPoints, heightPoints)`, plus `Landscape()` and `WithMargins(...)`. Every producer -
`HtmlToDocxConverter`, `HtmlToPdfConverter` and `DocxEditor.Create` - takes one.

`DocxToPdfConverter` takes no `PageSetup`: it renders a document that already carries its own page
setup, and honours it.

Page setup and remote images combine: `ConvertAsync(html, page, options)`. Before 0.18.0 they were
mutually exclusive - `(html, page)` converted offline and `(html, options)` laid out on A4 - so
asking for both silently dropped one.

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
| **A form's typed controls are only as checkable as the document makes them** | `DocxForm.Validate` reports a value that does not fit a **typed** control - a drop-down value outside its list, a non-date for a date picker, a non-boolean for a check box. A plain text control has no constraint, so any value validates. A clean result therefore means "nothing detectably wrong", not "every value is right". Related: `DocxForm.Inspect` returns what the document exposes under the `DocxFormKey` in use, which is not always every control - see *Fill-in forms*. |
| **PDF fidelity is bounded, and unsupported features drop silently** | Conditional formatting and some shape effects are omitted rather than reported; the PDF converters have no warning channel, because the renderer beneath them produces no report to surface. **A chart is not one of those omissions when it was added by this library** — `WorkbookEditor.AddChart` and `PresentationEditor.AddChart` produce a chart that renders correctly in XLSX → PDF and PPTX → PDF, title and category labels included, measured directly rather than assumed. DOCX has no chart-authoring path at all (see *Charts* in the [spreadsheets and presentations guide](https://ank-khoaho.github.io/DocToolkit/guides/editing/spreadsheets-and-presentations.html)), so a Word chart's PDF fidelity remains unmeasured. **Three losses are MEASURED rather than listed, and all are content, not styling.** Converted 2026-08-25 and read back out of the PDF, each fixture carrying a sibling paragraph so a missing token could not be confused with an empty render: **an unstyled footnote reference loses its text** — the render keys on the character style Word, and `DocxEditor.AddFootnote`, always applies to a footnote reference, so a footnote authored either way survives; only a reference built by hand, or by another tool, without that style is lost — and **a table nested inside a table cell loses its content entirely**. Both produce a valid PDF and raise nothing. And **a content control inside a table loses its text** — in a cell, or wrapping a cell or a row — while the same control at body level renders. **Measured to survive, so you do not have to wonder:** text boxes render, a content control at body level renders, and so does one inside a table that is itself inside a text box. **`DocxToPdfPreflight` reports all three measured losses on a document you are about to convert** — see *Knowing before you convert* above; it lists what the file contains, and does not claim to know what the renderer dropped. The output is a valid PDF either way. The DOCX → HTML and DOCX → Markdown exporters are the exception — see `ConvertWithReport`. **Paragraph styles ARE resolved on the DOCX → PDF path** — a `Heading1` whose size lives only in `styles.xml`, with no direct run formatting, renders at that size; measured 2026-08-25, 24pt heading against 11pt body. So documents authored from a corporate template keep their heading hierarchy. (The style-resolution caveat in `ROADMAP.md` is about the unshipped page-image renderer, not this path.) |
| **Word content controls survive some exports and not others** | A `w:sdt` is the wrapper Word puts around content an author marked up, and it is usually the part of a template that varies. **Measured 2026-08-27**, every fixture carrying a sibling paragraph so a missing token could not be confused with an empty result: `DocxEditor.ExtractText` and `ReadTable` read all of them, and an ordinary paragraph or table cell survives every exporter. What is lost: **DOCX → HTML drops a block-level control** at body level, nested, or holding a table — but *keeps* one inside a table cell, so this is not "controls are unsupported"; **DOCX → Markdown drops every block-level control**, and also drops a paragraph whose ONLY content is an inline control, while keeping an inline control that sits beside ordinary text; and **DOCX → PDF drops a control inside a table cell**, which `DocxToPdfPreflight` now reports as a `Known` finding. A body-level control still renders to PDF. **None of these is fixable here**: those three paths are pass-throughs to OfficeIMO and nothing in this package walks the document on them. **And the report channel does not cover it evenly** — measured on a body-level control, `DocxToMarkdownConverter.ConvertWithReport` raises a warning naming the unsupported tag, while `DocxToHtmlConverter.ConvertWithReport` raises **none** and drops the text silently. |
| **PDF fonts depend on the machine doing the conversion** | Where a system font is available it is **embedded**: on a Windows dev box the same invoice produces a ~167 KB PDF carrying Arial-Regular and Arial-Bold. In a slim container with no fonts installed, nothing is embedded and the PDF falls back to the **base-14 standard fonts** (Helvetica), giving ~1.5 KB. **Both are valid and both render**, and Arial and Helvetica are metric-compatible so line breaks do not move — but the glyphs are not identical, so a PDF built in your container will not be byte-identical to one built on your laptop. Install fonts in the image if you need a specific face. |
| **HTML → PDF goes through DOCX** | So fidelity is bounded by what HtmlToOpenXml maps into WordprocessingML, not by what a browser would render. Complex CSS layout — flexbox, grid, floats, absolute positioning — does not survive. Text, headings, tables, lists, inline styling and images do. |
| **No external stylesheets** | `<link rel="stylesheet">` is not fetched, by design. Inline `<style>` and `style=` are honoured. |
| **Headers and footers are one line each** | A header or footer is a single aligned line of text and page-number fields, set on `PageSetup`. One running header and footer per document, plus an optional distinct first page — per-section headers and odd/even (mirrored) variants are not supported. |
| **One page setup per document** | `PageSetup` applies to the whole document; multiple sections with different paper is not supported. |
| **DOCX → HTML returns a full document, not a fragment** | Extract the body with a parser if you are embedding it. |
| **Formulas carry no cached value** | Excel recalculates on open and `ReadCell`/`ReadSheet` evaluate on read, but a reader that only reads cached values sees an empty cell until Excel has opened and saved the file. |
| **Pivot table results carry no cached value either, and not even `ReadCell` computes them** | `WorkbookEditor.AddPivotTable`'s result grid is populated only when Excel opens and recalculates the file — a harder version of the row above: a formula's value *is* computed by `ReadCell`/`ReadSheet` on read, while a pivot table's is not, because nothing in this library evaluates a pivot aggregation. `XlsxToPdfConverter` renders nothing where the results would be, for the same reason it renders a formula's literal text rather than its computed value. See the [spreadsheets and presentations guide](https://ank-khoaho.github.io/DocToolkit/guides/editing/spreadsheets-and-presentations.html#pivot-tables). |
| **An OLE-embedded object survives some `WorkbookEditor` operations and not others** | `WorkbookEditor.AddChart`, `AddPivotTable`, `Protect` and `Unprotect` go through `OfficeIMO.Excel.ExcelDocument`, editing the package in place, and preserve a worksheet's embedded object exactly. `SetCell`, `AppendRows`, and the rest of that class's ClosedXML-backed surface for editing an *existing* workbook go through `ClosedXML.Excel.XLWorkbook`, which reconstructs the package from its own object model on save — it silently drops the `<drawing>` element and its part that anchor the object to the sheet, while the embedded content's own bytes survive as an orphaned, unreachable part. Measured directly, not assumed: a picture ClosedXML inserted itself survives the identical `SetCell` round-trip, so this is specific to drawing content ClosedXML did not create. **`PresentationEditor` is unaffected** — every operation, including PPTX's `AddChart`, `RemoveSlides`, `ReorderSlides` and `InsertSlides`, was measured to preserve an embedded/linked OLE object correctly. **`DocxEditor.ReplaceText` and `ReplaceImage` were measured the same way and also preserve one** — its other editing operations (`FillRows`, footnotes/endnotes, table of contents, `Protect`/`Unprotect`) have not been measured against an embedded object. |
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

## How often it works on files it has never seen

A test suite proves only that a library agrees with itself — every fixture in this repository was
produced by the code under test. So the conversions are also run against
[govdocs1](https://digitalcorpora.org/corpora/file-corpora/files/), a public crawl of real `.gov`
documents. Measured on chunk `000`, 2026-08-20:

| conversion | succeeded | of |
|---|---|---|
| HTML → DOCX | **97.8%** | 181 real pages |
| legacy `.doc` → DOCX | **89.2%** | 111 real documents |
| HTML → PDF | **88.4%** | 181 real pages |

Reading PDFs is stronger: across **200 real PDFs, 4,588 pages, a dozen producers**, every operation
succeeded on every file it did not refuse — and the refusals were 11 permission-restricted
documents, reported as exactly that rather than as a failure.

**Legacy PowerPoint 97-2003 `.ppt` also converts to PDF**, through the same
`PptxToPdfConverter.Convert` call — **at 60.2%** over 88 real binary decks, which is a lower bar
than the OOXML path and is stated rather than rounded up. The refusals are dominated by one
upstream limitation and none produced a corrupt PDF.

**What is deliberately not in that table.** The corpus predates both `.pptx` and `.docx`, so there
is no measured **PPTX → PDF** or **DOCX → PDF** rate — chunk `000` contains neither format. Both
conversions ship and are covered by the test suite; what is missing is a rate on files this project
did not author, and manufacturing one by chaining `.doc` → DOCX → PDF would measure the chain rather
than the converter.

**Published because they are unflattering and still useful.** A rate below 100% is what real input
looks like, and the alternative is telling you what is *supported* and letting you discover the
rest.

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
