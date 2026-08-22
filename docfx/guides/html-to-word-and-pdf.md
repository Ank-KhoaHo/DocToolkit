---
description: Convert HTML to PDF and DOCX in C# without a browser - page setup, images, and what a conversion drops.
---

# HTML to Word and PDF

Two converters, one input. @DocToolkit.HtmlToDocxConverter produces a `.docx`, and
@DocToolkit.HtmlToPdfConverter produces a `.pdf`.

[!code-csharp[](../../samples/HtmlConversion/Program.cs#convert)]

## Why the PDF goes through Word first

There is no HTML renderer in this package. Every free one is a browser engine, and a browser engine
is a native binary — which would cost the guarantee the whole library is built on: that
`dotnet restore` is the entire install and the output runs on any platform .NET runs on.

So `HtmlToPdfConverter` builds a Word document and renders that. The consequence is worth
internalising, because it explains most of what does and does not survive a conversion:

**PDF fidelity is bounded by what WordprocessingML can express, not by what a browser would
render.** Text, headings, tables, lists, inline styling and images come through. Flexbox, grid,
floats and absolute positioning do not — Word has no equivalent to lay them out on.

Two related limits follow from the same design:

- `<link rel="stylesheet">` is never fetched. Inline `<style>` blocks and `style=` attributes are
  honoured. Nothing here opens a socket you did not ask it to.
- Unsupported features are dropped **silently**. There is no warning channel on the public API, so
  a chart or a conditional format simply will not be in the output. The PDF is valid either way.

If you need pixel-accurate HTML rendering, you need a browser, and you should reach for one
directly rather than expect this library to be one.

## Page size and margins

Every producer lays out on **A4** with one-inch margins unless you say otherwise. Pass a
@DocToolkit.PageSetup to choose something else.

[!code-csharp[](../../samples/HtmlConversion/Program.cs#page-setup)]

```text
Default      : 595.3 x 841.9 pt, margins 72/72/72/72 pt
This one     : 792 x 612 pt, margins 36/36/36/36 pt
Letter intact: True  (Landscape() did not mutate it)
```

Everything is in **points**, the unit Word and PDF both use natively — 72 to the inch. `PageSetup`
converts to OOXML's twentieths-of-a-point internally, so A4 lands on exactly the 11906 × 16838 that
Word itself writes.

`PageSetup` is **immutable**. `Landscape()`, `WithMargins()` and friends each return a new
instance, which is why `PageSetup.A4` and `PageSetup.Letter` are safe to read from any thread and
safe to hand around a long-running application. That last output line is the sample checking it: it
derived a landscape page from `PageSetup.Letter` and `PageSetup.Letter` is still portrait.

`PageSetup.Custom(width, height)` covers sizes that have no named property. It rejects zero,
negative, and `NaN` dimensions rather than producing a document Word will refuse to open.

Page setup and remote images can be given together — `ConvertAsync(html, page, options)` on
either converter. Until 0.18.0 they could not: `(html, page)` always converted offline and
`(html, options)` always laid out on A4, so asking for both silently lost one.

> [!NOTE]
> One page setup applies to the whole document. A document with a landscape section in the middle
> of a portrait one is not something this API expresses.

### If you were on 0.12.x

Before 0.13.0 a generated DOCX stated **no page size at all**, so Word fell back to its Normal
template — US Letter on a US install, A4 on most others — and a generated PDF was always US
Letter. The same content printed on different paper depending on who opened it. Passing
`PageSetup.Letter` restores the old PDF behaviour explicitly.

## Images the HTML points at

An `<img src="https://…">` in HTML you did not write is a request that **your server** fetch a URL
of someone else's choosing. That is a server-side request forgery primitive, so remote downloads
are **off by default**, and a document converts successfully with the image absent rather than
failing.

[!code-csharp[](../../samples/DocxImages/Program.cs#remote-default)]

When you do want them, opt in per call with @DocToolkit.RemoteImageOptions. The allow-list is the
part that matters — the timeout and size cap are damage control, not access control.

[!code-csharp[](../../samples/DocxImages/Program.cs#remote-allowlist)]

```text
Remote off   : 1,746 bytes (default - image dropped)
Not on list  : 1,746 bytes (cdn.example.com refused, no request made)
```

Those two documents are byte-for-byte the same size, which is the guarantee made visible: a host
that is not on the list is refused **before any connection is attempted**. No DNS lookup, no
socket. The check happens in the library, not at the network layer.

Requests that clear the allow-list are then checked against the address they resolve to.
Loopback, link-local, and every private range are refused unless `AllowPrivateAddresses` is set —
so an allowed host whose DNS answer points at `169.254.169.254` still does not reach your cloud
metadata endpoint.

**A refused fetch is silent.** Your document succeeds with the image missing. That is the right
default for a conversion pipeline, and it is also exactly the failure you will not notice in
production — which is why the fetch path is the one thing in this library that emits telemetry.
See [Remote-image telemetry](production.md#telemetry).

## Once you have a PDF

@DocToolkit.PdfEditor works on a PDF that already exists — the only part of this library that reads
one rather than writing it. Nothing here re-renders, so the fidelity limits above do not apply: a
page that came out of the renderer looking right still looks right after being merged and
extracted.

[!code-csharp[](../../samples/PdfUtilities/Program.cs#merge)]

```text
Three documents : 1 + 1 + 1 pages
Merged          : 3 pages, 387,361 bytes
Page 2 alone    : 1 page, 131,408 bytes
```

`firstPage` is **1-based**, the way a reader numbers pages rather than the way an array indexes
them. A range that is not entirely inside the document is refused rather than silently clamped, so
a slice running off the end is a bug you hear about instead of a short document you do not.

Merging nothing is refused too: a zero-page PDF is not a useful artefact and several readers will
not open one.

### Metadata

What a file manager shows in its properties panel, and what a search indexer reads.

[!code-csharp[](../../samples/PdfUtilities/Program.cs#metadata)]

Every @DocToolkit.PdfMetadata property is nullable, and `null` means **absent** rather than blank —
in both directions, which is the part worth remembering:

- **Reading**, that separates "no subject" from "a subject deliberately set to empty". Anything
  combining metadata from several sources needs the difference: the first should take a fallback,
  the second should not.
- **Writing**, a `null` property leaves what the document already had in place, so stamping a title
  cannot quietly erase the author. Pass an empty string to clear a field on purpose.

```text
After retitling : title "Superseded", author still "Contoso Ltd"
```

Injectable as @DocToolkit.Extensions.DependencyInjection.IPdfEditor if you are using the extensions
package — see [Dependency injection](dependency-injection.md).

## Writing somewhere other than memory

The `byte[]` overloads above are the common case. Both converters also take a destination
`Stream`, and both offer `ConvertToFileAsync` when a path is what you have. All of them accept a
`CancellationToken`, and honour it while reading.

```csharp
await HtmlToPdfConverter.ConvertAsync(html, PageSetup.Letter, response.Body, ct);
await HtmlToDocxConverter.ConvertToFileAsync(html, "invoice.docx", ct);
```

What the `Stream` overloads do and do not buy you is covered in
[Running in production](production.md#streaming).
