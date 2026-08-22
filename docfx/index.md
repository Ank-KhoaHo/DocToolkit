---
_layout: landing
title: Convert HTML to PDF and DOCX in .NET without a browser
description: A .NET library that converts HTML and Markdown to DOCX and PDF and edits DOCX, XLSX and PPTX. Pure managed, no native binaries, no browser, runs on Linux and offline.
---

# DocToolkit

Convert **HTML and Markdown → DOCX and PDF**, export **DOCX → HTML/Markdown** and
**XLSX → CSV/HTML**, read text back out of a **PDF**, and open/edit **DOCX, XLSX and PPTX**,
from .NET.

**Pure managed. No native binaries, no browser, no LibreOffice, no Office interop.**
Works after `dotnet restore` alone, runs on Linux, and makes **no network calls at runtime**.
Verified **trim-safe and native-AOT compatible** by publishing and running the real thing in CI.

```bash
dotnet add package Ank.DocToolkit
```

Targets `net8.0` and `net10.0`. MIT licensed.

## Start here

- **[Getting started](guides/getting-started.md)** — install, your first conversion, and the three
  conventions the whole API follows
- **[What it can convert](guides/capabilities.md)** — the complete grid, generated from the shipped
  API rather than written by hand
- **[HTML to Word and PDF](guides/html-to-word-and-pdf.md)** — page setup, and what happens to an
  `<img>` that points at a URL
- **[Markdown, and conversion loss](guides/markdown.md)** — Markdown in and out, and what
  `ConvertWithReport` tells you
- **[Word documents](guides/word-documents.md)** — fill a template, build one from scratch, export
  it again
- **[Spreadsheets and presentations](guides/spreadsheets-and-presentations.md)** — XLSX and PPTX,
  report formatting, and CSV/HTML export
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
