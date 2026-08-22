---
description: Convert Markdown to DOCX and PDF in .NET, and export DOCX back to Markdown - including what a conversion drops.
---

# Markdown, and what a conversion drops

Markdown is the format that documentation, changelogs, release notes and model output already
arrive in. @DocToolkit.MarkdownToDocxConverter and @DocToolkit.MarkdownToPdfConverter take it
straight to a document; @DocToolkit.DocxToMarkdownConverter goes the other way.

This guide is also where **conversion loss** lives, because Markdown is where you feel it first: it
is the least expressive format here, so it is the one that most obviously cannot carry everything.

## Markdown in

[!code-csharp[](../../samples/MarkdownConversion/Program.cs#convert)]

Headings, emphasis, lists, links, code blocks, blockquotes and **tables** all come through. The
table is the one worth checking rather than assuming: the sample's document reports
`DocxEditor.TableCount == 1`, so it is a real Word table with real cells — not a block of
monospaced text pretending to be one.

**The PDF pivots through DOCX**, exactly as `HtmlToPdfConverter` does, and for the same reason: no
permissively-licensed, NuGet-only, Linux-safe library renders either format to PDF directly. So
everything in [Why the PDF goes through Word first](html-to-word-and-pdf.md#why-the-pdf-goes-through-word-first)
applies unchanged.

> [!NOTE]
> These two converters take a `string`, not `byte[]`, because Markdown *is* text. There is no
> `ConvertFile` overload for the same reason — read the file yourself and hand over the string.
> `ConvertAsync(markdown, destination, ct)` writes to a `Stream`.

## Nothing here reaches the network or the disk

The same guarantee the HTML converters make, arrived at the same way — by there being no mechanism
rather than by a flag being off:

- **A remote image is never fetched.** It becomes a **hyperlink** carrying the alt text, so the
  content stays reachable rather than vanishing.
- **A local file reference is refused**, not read. `![](../../etc/passwd)` in a document somebody
  else wrote is a file-disclosure primitive, and the answer is not to have a code path for it.
- **`data:` URIs are honoured**, because they carry their own bytes and cost no I/O. They are
  capped at 32 MB so a document cannot make your process materialise arbitrary memory.

## What the conversion could not carry across

`Convert` returns the document and discards everything the conversion had to say about it.
`ConvertWithReport` hands both back.

[!code-csharp[](../../samples/MarkdownConversion/Program.cs#report)]

```text
  Approximation MarkdownToWordWarning    Remote image 'https://img.example.com/badge.svg' was not resolved because MarkdownToWordOptions.RemoteImageResolver is not configured.

HasLoss      : True
What it says : Release notes / build status / Everything shipped.
```

@DocToolkit.ConversionResult`1 is `Value` plus `Warnings`, with `HasLoss` derived from whether that
list is empty. There is no `Succeeded` property — a conversion that failed threw, so a result you
are holding always succeeded.

Each @DocToolkit.ConversionWarning carries a @DocToolkit.ConversionLossKind, and the distinction is
the useful part:

| Kind | Means | In the example |
|---|---|---|
| `Approximation` | It came across, in a different form | The image became a hyperlink — "build status" is still in the document |
| `Omission` | It did not come across | A feature with no representation in the target format |
| `Failure` | The converter could not process it | Malformed input the parser gave up on |

That last output line is the sample checking the label rather than trusting it: `Approximation`
claims the content survived in another form, and reading the text back proves it did.

> [!WARNING]
> **That message names `MarkdownToWordOptions.RemoteImageResolver`, which you cannot set.** The
> text comes from the underlying library and is passed through verbatim. In this package the
> resolver is deliberately and permanently `null` — that *is* the offline guarantee, not a default
> waiting to be configured. Read the message as "there is no remote image here", not as a
> suggestion.

### The other direction, and the empty report

[!code-csharp[](../../samples/MarkdownConversion/Program.cs#round-trip)]

DOCX → Markdown is the lossier direction by a wide margin: Word expresses far more than Markdown
can. The same `ConvertWithReport` pair exists on @DocToolkit.DocxToHtmlConverter and
@DocToolkit.DocxToMarkdownConverter, taking `byte[]` and returning a
`ConversionResult<string>`.

The sample's round trip reports **zero** warnings, which is worth reading carefully. It means
nothing was lost *that the converter knows about* — not that nothing was lost. An empty
`Warnings` list is a statement about the converter's coverage as much as about your document.

> [!IMPORTANT]
> **The PDF renderers have no report to give.** `DocxToPdfConverter`, `HtmlToPdfConverter`,
> `XlsxToPdfConverter` and `PptxToPdfConverter` offer no `ConvertWithReport`, and that absence is
> deliberate. They drop charts, conditional formatting and some shape effects **silently**, and
> nothing downstream computes what was dropped — so a `ConvertWithReport` there would return an
> empty list on every call, which is a documented lie rather than a feature.

## Where to go next

- [What it can convert](capabilities.md) — the full grid, generated from the shipped API
- [Word documents](word-documents.md) — exporting a DOCX to HTML or Markdown
- [HTML to Word and PDF](html-to-word-and-pdf.md) — the other markup input, and the network guard
