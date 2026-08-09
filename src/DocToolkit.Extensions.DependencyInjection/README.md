# Ank.DocToolkit.Extensions.DependencyInjection

[![NuGet](https://img.shields.io/nuget/v/Ank.DocToolkit.Extensions.DependencyInjection.svg)](https://www.nuget.org/packages/Ank.DocToolkit.Extensions.DependencyInjection/)
[![NuGet Downloads](https://img.shields.io/nuget/dt/Ank.DocToolkit.Extensions.DependencyInjection.svg)](https://www.nuget.org/packages/Ank.DocToolkit.Extensions.DependencyInjection/)

Dependency-injection registration for [Ank.DocToolkit](https://www.nuget.org/packages/Ank.DocToolkit) —
`services.AddDocToolkit()` registers eleven injectable interfaces over the same pure-managed
HTML/DOCX/PDF/XLSX/PPTX conversion and editing logic.

```bash
dotnet add package Ank.DocToolkit.Extensions.DependencyInjection
```

Targets `net8.0` and `net10.0`. MIT licensed.

📖 **[Dependency injection guide](https://ank-khoaho.github.io/DocToolkit/guides/dependency-injection.html)**
— registration, options, and why they are read per call rather than captured at startup ·
🔎 **[API reference](https://ank-khoaho.github.io/DocToolkit/)**

## Usage

```csharp
using DocToolkit.Extensions.DependencyInjection;

services.AddDocToolkit();

// Or opt in to remote image download for HTML->DOCX/PDF. This still succeeds in an air-gapped
// environment - an unreachable host leaves that image out rather than failing the conversion.
services.AddDocToolkit(o => o.AllowRemoteImageDownload = true);
```

### Configuration reload takes effect immediately

`DocToolkitOptions` is consumed through `IOptionsMonitor` and read on **every call**, so changing
configuration at runtime applies without restarting the process.

That matters most for `AllowRemoteImageDownload`, which is the only switch that lets this library
open a socket. Setting it to `false` in configuration - as an incident response, say - stops the
next conversion fetching, rather than the next deployment.

The services remain singletons; only the option read is live.

### Bounding the remote-image opt-in

`AllowRemoteImageDownload` is the only switch that decides *whether* anything is fetched. When it
is `true`, every fetch is bounded by `RemoteImage`, whose defaults are already the restrictive
ones — **loopback, private and link-local addresses are refused** (including `169.254.169.254`, the
cloud metadata endpoint), only `http` and `https` are spoken, redirects are not followed, and each
fetch is capped at 10 seconds and 5 MB counted on bytes actually read.

```csharp
services.AddDocToolkit(o =>
{
    o.AllowRemoteImageDownload = true;
    o.RemoteImage.Timeout = TimeSpan.FromSeconds(3);
    o.RemoteImage.AllowedHosts.Add("cdn.example.com");   // empty means "any public host"
});
```

`RemoteImage` is configured **in place**, not assigned: the property is get-only so that a
restrictive default cannot be lost by dropping in an object that missed one.

> **Fetching from an intranet image host?** The address block refuses private ranges by default, so
> that image is skipped silently. Set `o.RemoteImage.AllowPrivateAddresses = true` to allow it —
> and be aware that doing so is what re-opens the SSRF reach if any caller converts untrusted HTML.

**This is not a complete SSRF defence.** A host's address is resolved and checked, then resolved
again by the HTTP stack when it connects; a DNS answer that changes in between defeats the check.
See the [core package README](https://www.nuget.org/packages/Ank.DocToolkit) and
[`SECURITY.md`](https://github.com/Ank-KhoaHo/DocToolkit/blob/main/SECURITY.md).

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
// Every interface also has Stream-based async members, so a large document never has to be
// duplicated into a caller-visible byte[] — write straight to an HTTP response body instead:
record InvoiceRequest(string Html);

app.MapPost("/invoices/pdf", async (InvoiceRequest request, IHtmlToPdfConverter toPdf, HttpResponse response) =>
{
    response.ContentType = "application/pdf";
    // The PDF is written to response.Body as it is rendered, not assembled first, so the
    // status code and headers are committed on the first write. A failure part-way through
    // cannot be turned into a clean 500 — the response is already underway.
    await toPdf.ConvertAsync(request.Html, response.Body);
});
```

All eleven interfaces — `IHtmlToDocxConverter`, `IDocxToPdfConverter`, `IHtmlToPdfConverter`,
`IXlsxToPdfConverter`, `IPptxToPdfConverter`, `IDocxToHtmlConverter`,
`IDocxToMarkdownConverter`, `IDocxEditor`, `IWorkbookEditor`, `IPresentationEditor`,
`IPdfEditor` — mirror
[`Ank.DocToolkit`](https://www.nuget.org/packages/Ank.DocToolkit)'s static API, including both its
`byte[]` and its `Stream`-based async overloads. They are registered as singletons (each wraps
stateless logic) and are safe to inject and call concurrently. See the core package's README for
what each one does and the offline/licensing guarantees behind them.

Two things on the static API deliberately do **not** appear on these interfaces:

- **The file-path helpers** (`ConvertToFileAsync`, `ConvertFile`). Inject the converter, take the
  `byte[]` or write to a `Stream`, and put the bytes wherever they belong — that keeps the
  injected surface free of filesystem coupling.
- **The per-call `allowRemoteImageDownload` argument and `RemoteImageOptions` overloads.** Remote
  image download is configured once, at registration, via `DocToolkitOptions` — so whether an
  application may reach the network, and how far, is a property of how it is composed rather than a
  decision at each call site. It is `false` unless you opt in.

## Why a separate package

A console app, Lambda or simple script that only wants the static `byte[]`-based API installs
just `Ank.DocToolkit`, with zero DI dependencies. ASP.NET Core and worker-service consumers add
this package too.

## Licence

MIT — see the parent repository's [LICENSE](https://github.com/Ank-KhoaHo/DocToolkit/blob/main/LICENSE).
