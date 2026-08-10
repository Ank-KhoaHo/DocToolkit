# Getting started

DocToolkit converts HTML into Word documents and PDFs, and reads and edits DOCX, XLSX and PPTX
files. It is pure managed code: no native binaries, no headless browser, no LibreOffice, no Office
interop. `dotnet restore` is the whole install, and nothing it does at runtime touches the network
unless you explicitly ask it to.

## Install

```bash
dotnet add package Ank.DocToolkit
```

That is the library. Everything in it is a static class, so there is nothing to register and no
container to configure. If you are in ASP.NET Core or a worker service and would rather inject
interfaces, add the companion package as well — see [Dependency injection](dependency-injection.md).

```bash
dotnet add package Ank.DocToolkit.Extensions.DependencyInjection
```

Both target `net8.0` and `net10.0`, and both are MIT licensed.

## Your first conversion

[!code-csharp[](../../samples/HtmlConversion/Program.cs#convert)]

Three things worth noticing in those three lines.

**HTML → PDF pivots through DOCX.** There is no direct HTML renderer here, because every free one
is a browser and a browser is a native binary. `HtmlToPdfConverter` builds a Word document and
renders that. This is why the PDF from `HtmlToPdfConverter` and the PDF from `DocxToPdfConverter`
above are the same size — they are the same document.

**Nothing was written to disk.** The default overloads take and return `byte[]`, which is usually
what a web handler wants. File and `Stream` overloads exist for when it isn't.

**Nothing reached the network.** An `<img src="https://…">` in that HTML would have been dropped,
not fetched. See [Remote images](html-to-word-and-pdf.md#images-the-html-points-at) for how to opt
in for specific hosts.

## The shape of the API

Every type follows the same three conventions, so learning one teaches you the rest.

| You have | You want | Use |
|---|---|---|
| A `byte[]` | A `byte[]` | `Convert(bytes)` — synchronous, no allocation surprises |
| A `Stream` | A `Stream` | `ConvertAsync(source, destination, ct)` |
| A path | A path | `ConvertFile(in, out)` or `…ToFileAsync(in, out, ct)` |

The `Stream` overloads are not merely wrappers that buffer into an array — they exist so a large
document never has to be resident in memory twice. What they do and do not guarantee is covered in
[Running in production](production.md#streaming).

Producers — anything that creates a document rather than reading one — take an optional
@DocToolkit.PageSetup. Left out, they lay out on A4. See
[Page size and margins](html-to-word-and-pdf.md#page-size-and-margins).

## When something goes wrong

Everything the library raises on your behalf arrives as a single exception type,
@DocToolkit.DocumentConversionException, so a caller needs one `catch` rather than one per
underlying library. The original failure is preserved as `InnerException`, which is what you want
in a log.

[!code-csharp[](../../samples/HtmlConversion/Program.cs#errors)]

```text
Rejected     : Failed to render DOCX to PDF.
Inner cause  : FileFormatException
```

A caller that hands user-supplied bytes to any reader should expect this: "is this really a DOCX"
is not a question you can answer from a filename or a content type.

## Where to go next

- [HTML to Word and PDF](html-to-word-and-pdf.md) — page setup, remote images, the network guard
- [Word documents](word-documents.md) — fill a template, build one from scratch, export it again
- [Spreadsheets and presentations](spreadsheets-and-presentations.md) — XLSX and PPTX
- [Dependency injection](dependency-injection.md) — `AddDocToolkit()` and the injectable interfaces
- [Running in production](production.md) — streaming, containers, trimming, telemetry, limits

Most code blocks in these guides are pulled from a
[runnable sample](https://github.com/Ank-KhoaHo/DocToolkit/tree/main/samples) that CI compiles
against the published package on Linux, Windows and macOS — if a snippet here is wrong, the build
is red. A few examples show API that has not reached the published package yet; those are marked
with a note at the point they appear.
