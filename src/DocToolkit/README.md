# DocToolkit

Convert HTML to DOCX and PDF, and open/edit DOCX, XLSX and PPTX from .NET.

**Pure managed.** No native binaries, no browser, no LibreOffice, no Office interop.
Works after `dotnet restore` alone, and runs on Linux.

## Install

```bash
dotnet add package DocToolkit
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

// Presentations
byte[] pptx = File.ReadAllBytes("deck.pptx");
int slides = PresentationEditor.SlideCount(pptx);
IReadOnlyList<string> slideText = PresentationEditor.ExtractText(pptx);   // in deck order
byte[] editedPptx = PresentationEditor.ReplaceText(pptx, new Dictionary<string, string>
{
    ["{{title}}"] = "Q3 Results",
});
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

## No network access

`HtmlToDocxConverter` and `HtmlToPdfConverter` make **no outbound requests** by default.
Images referenced by an absolute `http`/`https` URL are skipped; only `data:` URI images are
embedded. The underlying HTML converter defaults to downloading them, which would give every
caller of a `byte[]`-in/`byte[]`-out API an SSRF reach and an unbounded hang, so DocToolkit
turns that off and makes it opt-in:

```csharp
// Only for markup you trust. Issues HTTP requests to whatever hosts the markup names, and a
// host that fails to serve the image fails the whole conversion.
byte[] withImages = await HtmlToDocxConverter.ConvertAsync(
    html, allowRemoteImageDownload: true, ct);
```

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
