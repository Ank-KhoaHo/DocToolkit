# DocToolkit

Convert HTML to DOCX and PDF, and open/edit DOCX, XLSX and PPTX from .NET.

**Pure managed.** No native binaries, no browser, no LibreOffice, no Office interop.
Works after `dotnet restore` alone, and runs on Linux.

## Offline by default — safe in air-gapped environments

**No method on DocToolkit's public API opens a network connection.** Not for images, not for
stylesheets, not for fonts, not for linked pictures or external workbook references. Once the
package is restored, DocToolkit never needs the network again.

There is exactly one way to change that, and you have to ask for it by name:

```csharp
// The ONLY API that makes an outbound request. It downloads and embeds the images the markup
// names, so it FAILS in an air-gapped environment - a host that will not answer fails the whole
// conversion, after a connect timeout. Leave it alone unless your machines have internet access.
byte[] docx = await HtmlToDocxConverter.ConvertAsync(html, allowRemoteImageDownload: true, ct);
byte[] pdf  = await HtmlToPdfConverter.ConvertAsync(html, allowRemoteImageDownload: true, ct);
```

Everything else — `ConvertAsync(html)`, `ConvertToFileAsync`, `DocxToPdfConverter`, `DocxEditor`,
`WorkbookEditor`, `PresentationEditor` — is offline, unconditionally.

This is enforced, not merely intended. The test suite starts a real TCP listener on loopback,
feeds every public API markup that names it as an `<img src>`, a `<link rel="stylesheet">`, a CSS
`@import`, a `background-image`, an `<a href>`, an externally linked DOCX picture, an external
XLSX workbook link and more, and requires the accepted-connection count to be **exactly zero**. A
companion test points the same APIs at an unroutable address (TEST-NET-3) and requires them to
return promptly rather than stall on a connect timeout.

`dotnet restore` is the one step that still needs a package feed. `THIRD-PARTY-NOTICES.txt` lists
the full dependency closure with resolved versions, so it can be mirrored onto an internal feed;
every entry is a plain managed assembly with no native payload and no post-restore download.

## Install

```bash
dotnet add package Ank.DocToolkit
```

Targets `net8.0` and `net10.0`.

## Usage

```csharp
using DocToolkit;

// HTML -> DOCX
byte[] docx = await HtmlToDocxConverter.ConvertAsync("<h1>Invoice</h1><p>Total: 18,100.00</p>");

// HTML -> PDF (pivots through DOCX internally)
byte[] pdf = await HtmlToPdfConverter.ConvertAsync("<h1>Invoice</h1><p>Total: 18,100.00</p>");

// DOCX -> PDF
byte[] rendered = DocxToPdfConverter.Convert(docx);

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

// Presentations
byte[] pptx = File.ReadAllBytes("deck.pptx");
int slides = PresentationEditor.SlideCount(pptx);
IReadOnlyList<string> slideText = PresentationEditor.ExtractText(pptx);   // in deck order
byte[] editedPptx = PresentationEditor.ReplaceText(pptx, new Dictionary<string, string>
{
    ["{{title}}"] = "Q3 Results",
});
```

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
