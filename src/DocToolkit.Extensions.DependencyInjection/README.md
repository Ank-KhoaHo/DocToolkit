# Ank.DocToolkit.Extensions.DependencyInjection

![DocToolkit - convert HTML to PDF and DOCX in C#, no browser, no native binaries](https://raw.githubusercontent.com/Ank-KhoaHo/DocToolkit/main/assets/banner.png)

[![NuGet](https://img.shields.io/nuget/v/Ank.DocToolkit.Extensions.DependencyInjection.svg)](https://www.nuget.org/packages/Ank.DocToolkit.Extensions.DependencyInjection/)
[![NuGet Downloads](https://img.shields.io/nuget/dt/Ank.DocToolkit.Extensions.DependencyInjection.svg)](https://www.nuget.org/packages/Ank.DocToolkit.Extensions.DependencyInjection/)
[![License: MIT](https://img.shields.io/badge/licence-MIT-blue.svg)](https://github.com/Ank-KhoaHo/DocToolkit/blob/main/LICENSE)

Dependency-injection registration for [Ank.DocToolkit](https://www.nuget.org/packages/Ank.DocToolkit) —
`services.AddDocToolkit()` registers an injectable interface per capability, over the same
pure-managed HTML/DOCX/PDF/XLSX/PPTX/Markdown conversion and editing logic. They are named in full
below.

```bash
dotnet add package Ank.DocToolkit.Extensions.DependencyInjection
```

Targets `net8.0` and `net10.0`. MIT licensed.

📖 **[Dependency injection guide](https://ank-khoaho.github.io/DocToolkit/guides/dependency-injection.html)**
— registration, options, and why they are read per call rather than captured at startup ·
🔎 **[API reference](https://ank-khoaho.github.io/DocToolkit/)**

## Usage

`AddDocToolkit` is an extension method in the `DocToolkit.Extensions.DependencyInjection`
namespace — import it to bring `services.AddDocToolkit()` into scope:

<!-- BEGIN SNIPPET: readme-di-registration -->

```csharp
services.AddDocToolkit();

// Or opt in to remote image download for HTML->DOCX/PDF. This still succeeds in an
// air-gapped environment - an unreachable host leaves that image out rather than failing
// the conversion.
services.AddDocToolkit(o => o.AllowRemoteImageDownload = true);
```

<!-- END SNIPPET -->

### Configuration reload takes effect immediately

`DocToolkitOptions` is consumed through `IOptionsMonitor` and read on **every call**, so changing
configuration at runtime applies without restarting the process.

That matters most for `AllowRemoteImageDownload`, which is the only switch that lets this library
open a socket. Setting it to `false` in configuration - as an incident response, say - stops the
next conversion fetching, rather than the next deployment.

The services remain singletons; only the option read is live.

### Fonts for non-Latin text

Whether a document containing Cyrillic, Greek or CJK renders to PDF is otherwise a property of the
**machine** — the renderer falls back to whatever fonts the host has, so the same document converts
on one and is refused on another. `Fonts` takes that out of the answer:

```csharp
services.AddDocToolkit(o =>
    o.Fonts = new PdfFontOptions("Noto Sans", File.ReadAllBytes("NotoSans-Regular.ttf"))
                  .Add("Noto Sans CJK", File.ReadAllBytes("NotoSansCJK-Regular.ttf")));
```

Configured once rather than passed per call, because needing a font is a property of the deployment
rather than of the document: somebody converting Cyrillic needs it for every document, not for some.
Nothing is fetched and nothing is read from disk by this library — the bytes come from you.

**Supply fonts covering everything your documents use, not only the script that failed.** They
*replace* the host's own fallbacks rather than adding to them, so too few is worse than none —
measured over 99 real documents, one font rendered 63 where none rendered 71, and four rendered 77.

**`Fonts` applies to every converter that renders a PDF** — `IDocxToPdfConverter` and
`IHtmlToPdfConverter` alike, and on the same conversion as your page setup and remote-image
settings.

It reached only the first before **0.35.0**, because the core package had no overload carrying fonts
alongside the other two; wiring it anyway would have made the setting apply only when neither of the
others was in play, and a setting that silently stops taking effect depending on unrelated
configuration is worse than one documented as absent.

> **If you already set `Fonts` and convert HTML to PDF, this changes your output.** Those
> conversions previously ignored the setting and used the host's fonts. They now use yours — and
> supplied fonts **replace** the host's fallbacks rather than adding to them, so a list that covers
> less than your documents do can render *fewer* of them than before. Supply fonts covering
> everything you convert, or clear `Fonts` if you were relying on the host.

### Bounding the remote-image opt-in

`AllowRemoteImageDownload` is the only switch that decides *whether* anything is fetched. When it
is `true`, every fetch is bounded by `RemoteImage`, whose defaults are already the restrictive
ones — **loopback, private and link-local addresses are refused** (including `169.254.169.254`, the
cloud metadata endpoint), only `http` and `https` are spoken, redirects are not followed, and each
fetch is capped at 10 seconds and 5 MB counted on bytes actually read.

<!-- BEGIN SNIPPET: readme-di-options -->

```csharp
services.AddDocToolkit(o =>
{
    o.AllowRemoteImageDownload = true;
    o.RemoteImage.Timeout = TimeSpan.FromSeconds(3);
    o.RemoteImage.AllowedHosts.Add("cdn.example.com");   // empty means "any public host"
});
```

<!-- END SNIPPET -->

`RemoteImage` is configured **in place**, not assigned: the property is get-only so that a
restrictive default cannot be lost by dropping in an object that missed one.

> **Fetching from an intranet image host?** The address block refuses private ranges by default, so
> that image is skipped silently. Set `o.RemoteImage.AllowPrivateAddresses = true` to allow it —
> and be aware that doing so is what re-opens the SSRF reach if any caller converts untrusted HTML.

**This is not a complete SSRF defence.** A host's address is resolved and checked, then resolved
again by the HTTP stack when it connects; a DNS answer that changes in between defeats the check.
See the [core package README](https://www.nuget.org/packages/Ank.DocToolkit) and
[`SECURITY.md`](https://github.com/Ank-KhoaHo/DocToolkit/blob/main/SECURITY.md).

<!-- BEGIN SNIPPET: readme-di-consume -->

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

    public Task<byte[]> RenderDocxAsync(string html) => _toDocx.ConvertAsync(html);
    public Task<byte[]> RenderAsync(string html) => _toPdf.ConvertAsync(html);
}
```

<!-- END SNIPPET -->

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

Every interface — `IHtmlToDocxConverter`, `IDocxToPdfConverter`, `IHtmlToPdfConverter`,
`IXlsxToPdfConverter`, `IPptxToPdfConverter`, `IDocxToHtmlConverter`,
`IDocxToMarkdownConverter`, `IMarkdownToDocxConverter`, `IMarkdownToPdfConverter`,
`IXlsxToCsvConverter`, `IXlsxToHtmlConverter`, `IDocToDocxConverter`, `IDocxEditor`,
`IWorkbookEditor`, `IPresentationEditor`, `IPdfEditor`, `IDocxReview`, `IDocxMailMerge`,
`IDocxForm`, `IDocxToPdfPreflight`, `IMarkdownEditor` — mirrors
[`Ank.DocToolkit`](https://www.nuget.org/packages/Ank.DocToolkit)'s static API, including both its
`byte[]` and its `Stream`-based async overloads.

**That mirroring is now enforced rather than asserted.** It had gone stale nine times — most
recently with seven gaps at once, four of them whole interfaces that simply did not exist — because
the only check was a snippet someone had to remember to run against a hand-written list of pairs.
A test in this package now derives both sides by reflection and fails naming anything missing.

**The count is deliberately not written down here any more.** The interface NAMES above are checked
against the shipped API by `check-readme-coverage.py`, so the list cannot go stale silently — but a
number never was checked, and this file said "six" while ten shipped. The package `<Description>`
made the same mistake independently and now says nothing countable either. They are registered as singletons (each wraps
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

## Setting the paper once

`DocToolkitOptions.Page` is the page every producer lays out on when a call does not name
one. It defaults to `PageSetup.A4`, which is what the static API already uses, so leaving it
alone changes nothing.

<!-- BEGIN SNIPPET: readme-di-page-setup -->

```csharp
services.AddDocToolkit(o => o.Page = PageSetup.Letter);
```

<!-- END SNIPPET -->

It reaches all three producers - both HTML converters and `IDocxEditor.Create` - because an
option true of two out of three is one a consumer discovers a document at a time.

An explicit argument still wins: `ConvertAsync(html, PageSetup.A4)` produces A4 whatever the
option says, since a call naming a page is answering a narrower question than configuration.

> A null assigned to `Page` throws, but on **first use** rather than at registration - the
> configure delegate runs when the options are first materialised, not when
> `AddDocToolkit` is called.

## Migrating

### 0.20.x to 0.21.0 - `DocxEditor.ExtractText` now separates blocks

Before 0.21.0 `ExtractText` returned the document's raw concatenated text, with **no separator
between blocks at all**. A heading `Title` followed by a paragraph `Body text.` came back as the
single token `TitleBody text.`, and adjacent table cells `A` and `B` came back as `AB`. Word
boundaries were lost, so anything that tokenised, indexed or diffed the result got fused words.

From 0.21.0 blocks are separated by `\n` and the cells of a table row by `\t` - which is what
Word's own *save as plain text* writes, and what this method already did between the body and any
headers or footers.

```csharp
// 0.20.x  ->  "TitleBody text."
// 0.21.0  ->  "Title\nBody text."
string text = DocxEditor.ExtractText(docx);
```

Substring checks such as `text.Contains("Title")` are unaffected. **Exact-match comparisons against
the old fused output will need updating** - that is the whole of the breaking change. If you need
the previous shape, `text.Replace("\n", "").Replace("\t", "")` reproduces it.

### 0.15.0 to 0.16.0 - the extensions package needs a newer DI abstraction

No behaviour change, but a **floor you may have to satisfy**.
`Ank.DocToolkit.Extensions.DependencyInjection` now requires
`Microsoft.Extensions.DependencyInjection.Abstractions` **8.0.2**, up from 8.0.0.

It follows from 0.15.0: PDFsharp raised `Microsoft.Extensions.Logging.Abstractions` from 6.0.0 to
8.0.3 in the core package's graph, and 8.0.3 requires DI abstractions >= 8.0.2. If your application
pins 8.0.0 or 8.0.1 you will see NuGet report a package downgrade rather than resolve silently -
raise your reference, or remove the pin and let it float.

The core package `Ank.DocToolkit` is unaffected.

## Why a separate package

A console app, Lambda or simple script that only wants the static `byte[]`-based API installs
just `Ank.DocToolkit`, with zero DI dependencies. ASP.NET Core and worker-service consumers add
this package too.

## Licence

MIT — see the parent repository's [LICENSE](https://github.com/Ank-KhoaHo/DocToolkit/blob/main/LICENSE).
