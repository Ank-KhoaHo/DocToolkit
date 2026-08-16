# What it can convert

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
| **DOCX** (`DocxEditor`) | `Create`, `ExtractText`, `FillRows`, `ReadTable`, `ReplaceImage`, `ReplaceText`, `TableCount` |
| **PDF** (`PdfEditor`) | `ExtractPages`, `ExtractText`, `InsertPages`, `Merge`, `PageCount`, `ReadMetadata`, `RemovePages`, `ReorderPages`, `RotatePages`, `WithMetadata` |
| **PPTX** (`PresentationEditor`) | `Create`, `ExtractText`, `ReplaceImage`, `ReplaceText`, `SlideCount` |
| **XLSX** (`WorkbookEditor`) | `AppendRows`, `Create`, `Format`, `ReadCell`, `ReadSheet`, `SetCell`, `SheetNames` |

Method names only. What each one does, and the traps in it, are in the guides — this table exists to be complete and current, which prose has repeatedly failed to be.

<!-- END GENERATED -->

## Where the detail is

- [HTML to Word and PDF](html-to-word-and-pdf.md) — page setup, remote images, the network guard
- [Markdown, and conversion loss](markdown.md) — Markdown in and out, and what `ConvertWithReport` tells you
- [Word documents](word-documents.md) — templates, repeating rows, images, text export
- [Spreadsheets and presentations](spreadsheets-and-presentations.md) — sheets, formulas, formatting, export
- [Running in production](production.md) — streaming, containers, fonts, trimming, telemetry
- [Dependency injection](dependency-injection.md) — the same surface as injectable interfaces

## Two things the tables do not say

**Every capability comes in three forms.** A `byte[]` overload, a `Stream` overload
(`…Async(source, destination, ct)`), and for most, a file-path overload. The table lists the
capability once; see [the shape of the API](getting-started.md#the-shape-of-the-api).

**A ✅ is not a fidelity claim.** Every conversion through the PDF renderers drops what it cannot
represent — charts, conditional formatting, some shape effects — and does so *silently*, because
those renderers produce no report to surface. The DOCX text exporters and the Markdown importers do
report their losses, through `ConvertWithReport`. Which is which, and what each one drops, is in the
guides and in the [known limitations](https://github.com/Ank-KhoaHo/DocToolkit#known-limitations).
