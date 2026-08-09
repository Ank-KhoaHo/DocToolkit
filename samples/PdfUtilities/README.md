# PDF utilities

Working on a PDF that **already exists**: page count, merge, page extraction, and document
metadata.

```bash
dotnet run --project samples/PdfUtilities
```

Builds three one-page PDFs, joins them, takes the middle page back out, stamps a title and author,
and shows what happens when you stamp again.

## The non-obvious part

**Nothing here re-renders.** Every other PDF operation in this library renders *into* PDF, and
carries the fidelity caveats that come with that. These do not: pages move between documents as
they are, so a page that came out of the renderer looking right still looks right after being
merged, extracted and stamped.

**`firstPage` is 1-based**, the way a reader numbers pages rather than the way an array indexes
them. A range that is not entirely inside the document is **refused rather than clamped** — a slice
running off the end is a bug you hear about, instead of a short document you do not. Merging nothing
is refused for the same reason: a zero-page PDF is not a useful artefact and several readers will
not open one.

**`null` means absent, not blank — in both directions.** This is the detail most worth taking away:

- *Reading*, it separates "no subject" from "a subject deliberately set to empty". Anything
  combining metadata from several sources needs that difference, because the first should take a
  fallback and the second should not. PDFsharp's own typed properties return `""` for a missing
  key, so the library reads through the underlying dictionary to preserve it.
- *Writing*, a `null` property leaves what the document already had alone. The sample stamps a
  title over an already-stamped document and prints that the author survived. Pass an empty string
  to clear a field on purpose.

## Why this sample arrived after the others

`PdfEditor` shipped in 0.15.0, and every sample here references the **published** package. So this
project could not exist until that release went out — which is the lag described in
[samples/README.md](../README.md) working as intended, not a gap somebody forgot to fill.
