---
_layout: landing
---

# DocToolkit

Convert **HTML → DOCX and PDF**, and open/edit **DOCX, XLSX and PPTX**, from .NET.

**Pure managed. No native binaries, no browser, no LibreOffice, no Office interop.**
Works after `dotnet restore` alone, runs on Linux, and makes **no network calls at runtime**.

```bash
dotnet add package Ank.DocToolkit
```

Targets `net8.0` and `net10.0`. MIT licensed.

## Two packages

- **[Ank.DocToolkit](https://www.nuget.org/packages/Ank.DocToolkit/)** — the library. Static
  classes, no DI container required. [API reference](xref:DocToolkit).
- **[Ank.DocToolkit.Extensions.DependencyInjection](https://www.nuget.org/packages/Ank.DocToolkit.Extensions.DependencyInjection/)**
  — `services.AddDocToolkit()`, for ASP.NET Core / worker-service consumers.
  [API reference](xref:DocToolkit.Extensions.DependencyInjection).

See the [GitHub repository](https://github.com/Ank-KhoaHo/DocToolkit) for source, runnable
samples (`samples/`), and the full README.
