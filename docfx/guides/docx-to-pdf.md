---
description: Convert DOCX to PDF in .NET with no Word installation, no LibreOffice and no Office interop - one call, pure managed, and it runs in a Linux container.
---

# Convert DOCX to PDF in .NET without Word or LibreOffice

`DocxToPdfConverter.Convert` renders a Word document to PDF in one call. **No Word installation, no
LibreOffice, no Office interop and no native binaries** — the whole conversion is managed code, so it
runs inside a Linux container and on arm64 exactly as it does on a Windows desktop.

```csharp
byte[] docx = File.ReadAllBytes("contract.docx");
byte[] pdf  = DocxToPdfConverter.Convert(docx);
```

There are `Stream` and file-path forms too, and an overload taking fonts — see
[Word documents](word-documents.md) for creating and editing the DOCX in the first place.

## Why this is usually hard

The three common approaches each cost something this one does not:

| approach | what it costs |
|---|---|
| **Office interop** | a licensed Word installation on the server, and a COM automation model Microsoft explicitly does not support for server-side use |
| **LibreOffice / headless** | a few hundred megabytes in your image, a process to supervise, and a CVE feed to watch |
| **A commercial SDK** | a licence fee, and for several of them a revenue threshold you have to re-check as you grow |

DocToolkit's conversion is a NuGet package and nothing else. `dotnet restore` is the whole install,
and CI re-checks on every push that the resolved dependency graph contains **no native binaries** and
**opens no socket at runtime**.

## How well it actually works

**A test suite only proves a library agrees with itself** — every fixture in this repository was
produced by the code under test. So the conversion is also measured against real documents.

Rendering 99 real Word documents to PDF, with no fonts supplied to the converter:

| fonts supplied | rendered |
|---|---|
| none | **71 / 99** |
| four | **77 / 99** |

**That number is published because it is unflattering and still useful.** The shortfall is almost
entirely non-Latin text: the renderer falls back to whatever fonts the host machine happens to have,
so the same document can convert on one machine and be refused on another. A refusal names the
character it could not encode, rather than silently producing a page of boxes.

**Supplying fonts takes the machine out of the answer** — and supplying too few is worse than
supplying none, because the fonts you pass *replace* the host's fallbacks rather than adding to them.
One font scored **63 / 99**, below supplying nothing at all. See
[Running in production](production.md#containers) for the whole of that.

```csharp
var fonts = new PdfFontOptions("Noto Sans", File.ReadAllBytes("NotoSans-Regular.ttf"));
byte[] pdf = DocxToPdfConverter.Convert(docx, fonts);
```

## What it does not do

- **It is not a layout-identical clone of Word.** Complex floating layouts, some field codes and
  certain drawing constructs are approximated. Text, headings, tables, images, page setup and
  headers/footers survive.
- **Nothing is fetched.** A document referencing a remote image does not get one; the conversion
  succeeds without it. That is the offline guarantee, not a limitation to work around.
- **PDF size varies with the host's fonts, by around 100x**, and both extremes are correct. A machine
  with Arial embeds it and produces a larger file; a bare container falls back to the standard fonts
  every PDF reader already has. Do not assert on the byte count.

## Related

- [Convert HTML to PDF and DOCX in C# without a browser](html-to-word-and-pdf.md)
- [Create and edit Word documents in C# without Office interop](word-documents.md)
- [Run document conversion in production: Linux, Docker and air-gapped hosts](production.md)
