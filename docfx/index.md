---
_layout: landing
title: Convert HTML to PDF and DOCX in .NET without a browser
description: A .NET library that converts HTML and Markdown to DOCX and PDF and edits DOCX, XLSX and PPTX. Pure managed, no native binaries, no browser, runs on Linux and offline.
---

# DocToolkit

<p class="lead">Convert and edit Office documents and PDFs from .NET. No browser, no native binaries, runs offline.</p>

## What do you want to do?

<!-- landing-cards:start -->
<div class="row row-cols-1 row-cols-sm-2 row-cols-lg-3 g-3 mb-4">
<div class="col"><a class="card h-100 text-decoration-none" href="guides/conversions/html-to-word-and-pdf.md"><div class="card-body"><div class="d-flex align-items-center gap-2 mb-2"><i class="bi bi-filetype-pdf fs-4 text-primary"></i><h3 class="h6 mb-0 text-body">HTML to PDF</h3></div><code class="small d-block text-break">HtmlToPdfConverter.ConvertAsync(html)</code></div></a></div>
<div class="col"><a class="card h-100 text-decoration-none" href="guides/conversions/html-to-word-and-pdf.md"><div class="card-body"><div class="d-flex align-items-center gap-2 mb-2"><i class="bi bi-filetype-docx fs-4 text-primary"></i><h3 class="h6 mb-0 text-body">HTML to Word</h3></div><code class="small d-block text-break">HtmlToDocxConverter.ConvertAsync(html)</code></div></a></div>
<div class="col"><a class="card h-100 text-decoration-none" href="guides/conversions/docx-to-pdf.md"><div class="card-body"><div class="d-flex align-items-center gap-2 mb-2"><i class="bi bi-file-earmark-richtext fs-4 text-primary"></i><h3 class="h6 mb-0 text-body">DOCX to PDF</h3></div><code class="small d-block text-break">DocxToPdfConverter.Convert(docx)</code></div></a></div>
<div class="col"><a class="card h-100 text-decoration-none" href="guides/conversions/markdown.md"><div class="card-body"><div class="d-flex align-items-center gap-2 mb-2"><i class="bi bi-markdown fs-4 text-primary"></i><h3 class="h6 mb-0 text-body">Markdown to PDF</h3></div><code class="small d-block text-break">MarkdownToPdfConverter.Convert(md)</code></div></a></div>
<div class="col"><a class="card h-100 text-decoration-none" href="guides/editing/word-documents.md"><div class="card-body"><div class="d-flex align-items-center gap-2 mb-2"><i class="bi bi-file-earmark-word fs-4 text-primary"></i><h3 class="h6 mb-0 text-body">Fill a Word template</h3></div><code class="small d-block text-break">DocxEditor.ReplaceText(docx, values)</code></div></a></div>
<div class="col"><a class="card h-100 text-decoration-none" href="guides/editing/spreadsheets-and-presentations.md"><div class="card-body"><div class="d-flex align-items-center gap-2 mb-2"><i class="bi bi-table fs-4 text-primary"></i><h3 class="h6 mb-0 text-body">Edit a spreadsheet</h3></div><code class="small d-block text-break">WorkbookEditor.SetCell(xlsx, "Sheet1", "B2", 42)</code></div></a></div>
</div>
<!-- landing-cards:end -->

## Quick start

```bash
dotnet add package Ank.DocToolkit
```

[!code-csharp[](../samples/HtmlConversion/Program.cs#convert)]

Targets `net8.0` and `net10.0`. MIT licensed.

## More guides

- **[Getting started](guides/getting-started.md)** — install, your first conversion, and the three
  conventions the whole API follows
- **[What it can convert](guides/capabilities.md)** — the complete grid, generated from the shipped
  API rather than written by hand
- **[Dependency injection](guides/dependency-injection.md)** — `AddDocToolkit()` and the injectable
  interfaces
- **[Running in production](guides/production.md)** — streaming, containers, fonts, trimming,
  telemetry

## Two packages

- **[Ank.DocToolkit](https://www.nuget.org/packages/Ank.DocToolkit/)** — the library. Static
  classes, no DI container required. [API reference](xref:DocToolkit).
- **[Ank.DocToolkit.Extensions.DependencyInjection](https://www.nuget.org/packages/Ank.DocToolkit.Extensions.DependencyInjection/)**
  — `services.AddDocToolkit()`, for ASP.NET Core / worker-service consumers.
  [API reference](xref:DocToolkit.Extensions.DependencyInjection).

See the [GitHub repository](https://github.com/Ank-KhoaHo/DocToolkit) for source, runnable
samples (`samples/`), and the full README.
