---
description: Create and edit DOCX in .NET - replace text, fill repeating table rows, swap images and read tables back.
---

# Create and edit Word documents in C# without Office interop

@DocToolkit.DocxEditor is the whole DOCX surface: fill a template somebody made in Word, build a
document from scratch when there is no template, read the text back out, and export it somewhere
that is not Word.

## Filling a template

The common case. Somebody in the business produces `invoice.docx` in Word with `{{customer}}` typed
where the name goes, and your job is to put the name there.

[!code-csharp[](../../samples/DocxTemplating/Program.cs#scalars)]

That looks too simple to need a library, and it would be, except for one thing:

**Word splits placeholders across runs.** `{{customer}}` is frequently three or four separate
`<w:t>` elements in the XML — because someone corrected a typo in the middle of it, or a spell
checker touched it, or it was pasted. A find-and-replace over the document XML finds nothing and
reports success, which is the single most common way hand-rolled Word templating fails. `ReplaceText`
splices the runs back together before matching, which is most of what it is for.

`ReplaceText` also reaches into headers and footers, so a customer name in a letterhead is
replaced too.

| you pass in | you get back |
|---|---|
| `invoice.docx` with `{{customer}}`, plus `{"customer": "Contoso Ltd"}` | the same document, `{{customer}}` replaced, formatting kept |
| a key with no matching placeholder | the document unchanged — no error, nothing to replace |
| a placeholder with no matching key | left as literal text, e.g. `{{unknown}}` stays in the output |

## One row per record

A table where the row count depends on your data — invoice lines, a roster, a statement. Put a
single template row in the document and let `FillRows` clone it.

[!code-csharp[](../../samples/DocxTemplating/Program.cs#rows)]

Each generated row keeps the template row's formatting: borders, shading, fonts, column widths.
That is the reason to do this rather than build the table yourself.

> [!IMPORTANT]
> **Expand rows first, then fill scalars.** `FillRows` clones the template row, so any scalar
> placeholder already inside it gets duplicated into every generated line. The sample follows this
> order even where it does not strictly matter, because the safe order is the one worth having in
> your fingers.

The collection name in the placeholder (`{{item.Desc}}`) matches the `collection` argument
(`"item"`). Anything the row does not consume is left alone.

## Images

`ReplaceImage` swaps a text placeholder for actual image bytes.

[!code-csharp[](../../samples/DocxImages/Program.cs#replace-image)]

Sizes are in points. Give one dimension and the other scales to keep the aspect ratio; give
neither and the image's own header decides, read at 96 DPI.

**The format is decided by the bytes, never by a filename.** PNG and JPEG are read by completely
different code paths — a PNG states its dimensions at a fixed offset in the IHDR chunk, while a
JPEG hides them in a Start-Of-Frame segment that has to be found by walking the marker chain. A
file called `logo.png` that actually holds JPEG bytes is read as the JPEG it is, because the
alternative is a blank frame in Word and no error anywhere.

For images referenced by URL rather than handed over as bytes, see
[Images the HTML points at](html-to-word-and-pdf.md#images-the-html-points-at).

## When there is no template

Sometimes the document's shape comes from your data, not from a file somebody made. Describe it as
a sequence of @DocToolkit.DocxBlock values and skip the round trip through HTML.

[!code-csharp[](../../samples/DocxTemplating/Program.cs#blocks)]

`DocxBlock` has four factories — `Heading`, `Paragraph`, `Table` and `Image` — which is
deliberately not a document model. It covers reports and statements. Anything that needs real
layout control wants a template, where a person with Word can do the layout.

`Create` takes an optional @DocToolkit.PageSetup, same as the HTML converters, and defaults to A4.

## Headers and footers

A header belongs to the page, so it goes on the @DocToolkit.PageSetup — which means every producer
honours it without a new overload.

[!code-csharp[](../../samples/DocxTemplating/Program.cs#headers)]

**The page number is a real field.** Written as text it would be fixed when the document was
generated — correct on one page, wrong on the rest, and looking right the whole time.

### A different first page

[!code-csharp[](../../samples/DocxTemplating/Program.cs#headers-first-page)]

Calling `WithFirstPage` is the switch, and **null means blank on page one** rather than "use the
ordinary one". That is the format's own model — there is no inheritance to fall back on — and it is
what makes a title page with nothing running across it expressible.

> [!NOTE]
> `ExtractText` shows the field's cached placeholder rather than a real page number: nothing
> computes pagination until a reader opens the document.

Not supported: per-section headers, and odd/even (mirrored) pages. A distinct first page — shown
above — is.

## Reading text back out

`ExtractText` returns the document's text as a string, which is what you want for search
indexing, diffing, or asserting in a test that the fill actually worked.

```csharp
string text = DocxEditor.ExtractText(docx);
string withChrome = DocxEditor.ExtractText(docx, includeHeadersAndFooters: true);
```

The default excludes headers and footers, because a letterhead repeated on every page is noise in
an index. Pass `true` when you want the whole thing.

## Exporting it somewhere else

The same document, as a PDF, as HTML for a web page, or as Markdown for a record you can diff.

[!code-csharp[](../../samples/DocxTemplating/Program.cs#export)]

```text
As HTML      : 795 chars, has a <table>: True
As Markdown  : 147 chars, first line "# Invoice for Contoso Ltd"
```

And @DocToolkit.DocxToPdfConverter for the PDF:

```csharp
byte[] pdf = DocxToPdfConverter.Convert(invoice);           // bytes in, bytes out
File.WriteAllBytes("invoice.pdf", pdf);                     // ...and onto disk, if that is where it goes

DocxToPdfConverter.ConvertFile("invoice.docx", "invoice.pdf");   // or path to path, no array
```

> The second line is plain .NET rather than anything from this library, and it is spelled out
> because leaving it implicit was confusing enough to be
> [reported](https://github.com/Ank-KhoaHo/DocToolkit/issues/321). Every converter here stops at a
> `byte[]` on purpose — see
> [Getting the bytes out](getting-started.md#getting-the-bytes-out) for the three forms and when
> each is the right one.

Two things to know before you wire these into something:

**The HTML is a full document, not a fragment.** `DocxToHtmlConverter.Convert` emits
`<html><head>…<body>`. There is no fragment mode — producing one would mean re-serialising the
renderer's output. If you are embedding the result in a page, extract the body with an HTML parser
rather than a regular expression. Both text converters embed images as `data:` URIs, so what you
get is self-contained with no asset files to host.

**Both text converters can tell you what they had to drop.** `ConvertWithReport` returns the same
string alongside a list of warnings, each saying whether something was approximated, omitted
outright, or failed. `Convert` is that call with the report discarded. See
[Markdown, and what a conversion drops](markdown.md#what-the-conversion-could-not-carry-across) —
and note that `DocxToPdfConverter` has no equivalent, deliberately.

**PDF fonts depend on the machine doing the conversion.** Where a system font is available it is
embedded; in a slim container with no fonts installed, nothing is embedded and the PDF falls back
to the base-14 standard fonts. Both are valid and both render, but they are not byte-identical.
See [Fonts](production.md#fonts-in-containers) before you compare PDF hashes across environments.
