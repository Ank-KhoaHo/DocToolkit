---
description: Every conversion DocToolkit supports, generated from the shipped public API rather than written by hand.
---

# Supported conversions: HTML, Markdown, DOCX, XLSX, PPTX and PDF

The complete list, in two tables. **Both are generated** from
`tests/DocToolkit.Tests/PublicApi/DocToolkit.approved.txt` — the reviewed record of what the package
actually ships — by `scripts/gen-capability-matrix.py`, and CI fails when they drift from it.

That is not ceremony. This library's capability list has been written by hand three times and gone
stale all three: the landing page described a version without Markdown in either direction, without
PDF text extraction and without spreadsheet export, five days after all three had shipped. Nobody
re-derives a table while adding a feature. So this one derives itself, and the guides linked below
carry the prose that a generator has no business writing.

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

## Editing an existing document

| Format | Operations |
|---|---|
| **DOCX** (`DocxEditor`) | `AddEndnote`, `AddFootnote`, `AddTableOfContents`, `Create`, `ExtractText`, `FillRows`, `InspectSignatures`, `IsProtected`, `Protect`, `ReadMetadata`, `ReadTable`, `ReplaceImage`, `ReplaceText`, `TableCount`, `Unprotect`, `ValidateSignatures`, `WithMetadata` |
| **DOCX** (`DocxMailMerge`) | `InspectTemplate`, `Merge`, `MergeBatch`, `MergeBatchToFiles`, `MergeBatchToFilesWithReport`, `MergeBatchWithReport`, `MergeConditional`, `MergeConditionalWithReport`, `MergeRepeating`, `MergeRepeatingRegions`, `MergeRepeatingRegionsWithReport`, `MergeRepeatingWithReport`, `MergeTableRowGroups`, `MergeTableRows`, `MergeWithReport` |
| **Markdown** (`MarkdownEditor`) | `FindHeading`, `ReadFrontMatter`, `ReadTable`, `ReplaceSection`, `TableCount` |
| **PDF** (`PdfEditor`) | `ExtractPages`, `ExtractText`, `InsertPages`, `Merge`, `PageCount`, `Protect`, `ReadMetadata`, `RemovePages`, `ReorderPages`, `RotatePages`, `Unprotect`, `WithMetadata` |
| **PPTX** (`PresentationEditor`) | `AddChart`, `Create`, `ExtractText`, `InsertSlides`, `InspectSignatures`, `IsProtected`, `Protect`, `ReadMetadata`, `ReadSlide`, `ReadSmartArt`, `RemoveSlides`, `ReorderSlides`, `ReplaceImage`, `ReplaceText`, `SlideCount`, `Unprotect`, `ValidateSignatures`, `WithMetadata` |
| **XLSX** (`WorkbookEditor`) | `AddChart`, `AddDefinedName`, `AddImage`, `AddPivotTable`, `AppendRows`, `Create`, `EvaluateFormulas`, `Format`, `InspectFormulas`, `InspectSignatures`, `IsProtected`, `Protect`, `ReadCell`, `ReadMetadata`, `ReadSheet`, `SetCell`, `SheetNames`, `Unprotect`, `ValidateSignatures`, `WithMetadata` |

Method names only. What each one does, and the traps in it, are in the guides — this table exists to be complete and current, which prose has repeatedly failed to be.

<!-- END GENERATED -->

## Legacy binary formats, which the grid above does not show

The grid is generated from the shipped converter names, so it can only describe the modern formats.
Two pre-2007 binary formats are also accepted, and one deliberately is not:

| input | supported | how |
|---|---|---|
| **`.doc`** | yes | `DocToDocxConverter` — text, or conversion to `.docx` |
| **`.ppt`** | yes, **to PDF only** | `PptxToPdfConverter` accepts one directly; `PresentationEditor` does not |
| **`.xls`** | **no** | refused immediately, with a message saying to save it as `.xlsx` |

`.xls` is refused for **cost**, not capability: measured on real files, a 101 KB workbook took 10.9
seconds and a 2.3 MB one did not finish in ten minutes, while the supported `.xlsx` path renders
20,000 rows in under four. See the package README for the full numbers.

## Measured on documents nobody here wrote

Every fixture in this repository's own test suite was produced by the code under test, so it can
only prove the library agrees with itself. These numbers come from
**[govdocs1](https://digitalcorpora.org/corpora/file-corpora/files/)**, a public crawl of real
`.gov` files, run monthly by
[`corpus.yml`](https://github.com/Ank-KhoaHo/DocToolkit/blob/main/.github/workflows/corpus.yml).
Measured on chunk
`000`, 2026-08-20:

| conversion | succeeded | of |
|---|---|---|
| HTML → DOCX | **97.8%** | 181 real pages |
| legacy `.doc` → DOCX | **89.2%** | 111 real documents |
| HTML → PDF | **88.4%** | 181 real pages |
| legacy `.ppt` → PDF | **60.2%** | 88 real decks |

Reading PDFs is separate and stronger: across **200 real PDFs, 4,588 pages, a dozen producers**,
every operation succeeded on every file it did not refuse — the only refusals were 11
permission-restricted documents, reported as exactly that.

**Published because the numbers are unflattering and still useful** — a rate below 100% is what
real input looks like. **Not measured**: PPTX → PDF and DOCX → PDF have no corpus rate, because
chunk `000` predates both formats; chaining `.doc` → DOCX → PDF to manufacture one would measure
the chain, not the converter.

## Where the detail is

- [HTML to Word and PDF](conversions/html-to-word-and-pdf.md) — page setup, remote images, the network guard
- [Markdown, and conversion loss](conversions/markdown.md) — Markdown in and out, and what `ConvertWithReport` tells you
- [Word documents](editing/word-documents.md) — templates, repeating rows, images, text export
- [Spreadsheets and presentations](editing/spreadsheets-and-presentations.md) — sheets, formulas, formatting, export
- [Running in production](production.md) — streaming, containers, fonts, trimming, telemetry
- [Dependency injection](dependency-injection.md) — the same surface as injectable interfaces

## Two things the tables do not say

**Every capability comes in three forms.** A `byte[]` overload, a `Stream` overload
(`…Async(source, destination, ct)`), and for most, a file-path overload. The table lists the
capability once; see [the shape of the API](getting-started.md#the-shape-of-the-api).

**A ✅ is not a fidelity claim.** Every conversion through the PDF renderers drops what it cannot
represent — conditional formatting, some shape effects — and does so *silently*, because those
renderers produce no report to surface. A chart is the one exception when it was added through this
library's own `AddChart` methods: it renders correctly in XLSX → PDF and PPTX → PDF, measured
directly. The DOCX text exporters and the Markdown importers do report their losses, through
`ConvertWithReport`. Which is which, and what each one drops, is in the guides and in the
[known limitations](https://github.com/Ank-KhoaHo/DocToolkit#known-limitations).
