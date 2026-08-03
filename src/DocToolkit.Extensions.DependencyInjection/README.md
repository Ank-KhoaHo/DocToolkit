# Ank.DocToolkit.Extensions.DependencyInjection

[![NuGet](https://img.shields.io/nuget/v/Ank.DocToolkit.Extensions.DependencyInjection.svg)](https://www.nuget.org/packages/Ank.DocToolkit.Extensions.DependencyInjection/)
[![NuGet Downloads](https://img.shields.io/nuget/dt/Ank.DocToolkit.Extensions.DependencyInjection.svg)](https://www.nuget.org/packages/Ank.DocToolkit.Extensions.DependencyInjection/)

Dependency-injection registration for [Ank.DocToolkit](https://www.nuget.org/packages/Ank.DocToolkit) —
`services.AddDocToolkit()` registers six injectable interfaces over the same pure-managed
HTML/DOCX/PDF/XLSX/PPTX conversion and editing logic.

```bash
dotnet add package Ank.DocToolkit.Extensions.DependencyInjection
```

Targets `net8.0` and `net10.0`. MIT licensed.

## Usage

```csharp
using DocToolkit.Extensions.DependencyInjection;

services.AddDocToolkit();
// or, to allow remote image download for HTML->DOCX/PDF (fails in an air-gapped environment):
services.AddDocToolkit(o => o.AllowRemoteImageDownload = true);
```

```csharp
public class InvoiceService
{
    private readonly IHtmlToDocxConverter _toDocx;
    private readonly IHtmlToPdfConverter _toPdf;

    public InvoiceService(IHtmlToDocxConverter toDocx, IHtmlToPdfConverter toPdf)
    {
        _toDocx = toDocx;
        _toPdf = toPdf;
    }

    public Task<byte[]> RenderAsync(string html) => _toPdf.ConvertAsync(html);
}
```

```csharp
// Every interface also has a Stream overload, so a large document never has to be buffered
// into a byte[] — write straight to an HTTP response body instead:
app.MapPost("/invoices/pdf", async (string html, IHtmlToPdfConverter toPdf, HttpResponse response) =>
{
    response.ContentType = "application/pdf";
    await toPdf.ConvertAsync(html, response.Body);
});
```

All six interfaces — `IHtmlToDocxConverter`, `IDocxToPdfConverter`, `IHtmlToPdfConverter`,
`IDocxEditor`, `IWorkbookEditor`, `IPresentationEditor` — mirror
[`Ank.DocToolkit`](https://www.nuget.org/packages/Ank.DocToolkit)'s static API one-for-one,
including its `Stream`-based async overloads, are registered as singletons (each wraps stateless
logic), and are safe to inject and call concurrently. See the core package's README for what each
one does and the offline/licensing guarantees behind them.

## Why a separate package

A console app, Lambda or simple script that only wants the static `byte[]`-based API installs
just `Ank.DocToolkit`, with zero DI dependencies. ASP.NET Core and worker-service consumers add
this package too.

## Licence

MIT — see the parent repository's [LICENSE](https://github.com/Ank-KhoaHo/DocToolkit/blob/main/LICENSE).
