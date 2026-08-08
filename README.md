# DocToolkit

[![CI](https://github.com/Ank-KhoaHo/DocToolkit/actions/workflows/ci.yml/badge.svg)](https://github.com/Ank-KhoaHo/DocToolkit/actions/workflows/ci.yml)
[![codecov](https://codecov.io/gh/Ank-KhoaHo/DocToolkit/branch/main/graph/badge.svg)](https://codecov.io/gh/Ank-KhoaHo/DocToolkit)
[![NuGet](https://img.shields.io/nuget/v/Ank.DocToolkit.svg?label=Ank.DocToolkit)](https://www.nuget.org/packages/Ank.DocToolkit/)
[![NuGet Downloads](https://img.shields.io/nuget/dt/Ank.DocToolkit.svg)](https://www.nuget.org/packages/Ank.DocToolkit/)
[![NuGet](https://img.shields.io/nuget/v/Ank.DocToolkit.Extensions.DependencyInjection.svg?label=Ank.DocToolkit.Extensions.DependencyInjection)](https://www.nuget.org/packages/Ank.DocToolkit.Extensions.DependencyInjection/)

Convert **HTML → DOCX and PDF**, and open/edit **DOCX, XLSX and PPTX**, from .NET.

**Pure managed. No native binaries, no browser, no LibreOffice, no Office interop.**
Works after `dotnet restore` alone, runs on Linux, and makes **no network calls at runtime**.

```bash
dotnet add package Ank.DocToolkit
```

Targets `net8.0` and `net10.0`. MIT licensed.

📖 [API documentation](https://ank-khoaho.github.io/DocToolkit/) · 📦
[package README](src/DocToolkit/README.md) · 📝 [CHANGELOG](CHANGELOG.md)

> **Contributing?** Branch from `main` and open a pull request back into it — `main` itself cannot
> be pushed directly. See [CONTRIBUTING.md](CONTRIBUTING.md) for the build, the commit-message
> rules, and the four constraints that will get a pull request rejected.

## Why this exists

Most .NET document stacks fail at least one of these. This one satisfies all four:

| Constraint | How |
|---|---|
| **Free for commercial use** | 16 dependencies: 15 MIT, 1 Apache-2.0. No revenue thresholds, no per-seat fees. |
| **NuGet only** | No Chromium download, no LibreOffice install, no native binaries. |
| **Runs on Linux** | Verified in CI on `ubuntu-24.04`, not inferred. |
| **Works offline** | No runtime network I/O. Proven by 37 air-gap tests. |

All four are properties of the *resolved dependency graph*, so a single upstream bump can break
them silently — which is why CI re-checks every one on every push. That has happened once already
(see [Design notes](#design-notes)).

### Trimming

Both packages are marked `IsTrimmable`, so `PublishTrimmed` apps keep only what they use.

That claim is checked the same way as the four above — CI trim-publishes an application over the
real dependency graph, then **runs it** and asserts every capability still works. A trim failure
does not appear at publish time: the trimmer removes a type something looks up by name, and the app
throws, or quietly produces an empty document, in production.

One caveat that is a dependency's rather than ours: **ClosedXML emits a trim warning**
(`IL2090`, in `DescribedEnumParser<T>`). Spreadsheet reading and writing work correctly in the
trimmed app CI runs, but the warning will appear in your publish output.

**Native AOT is not claimed.** `IsAotCompatible` is a strictly stronger promise than `IsTrimmable`,
and it has not been verified end to end here. An unverified compatibility claim is worse than an
absent one, so it is absent until a CI job compiles *and* runs an AOT build.

### Verifying what you downloaded

Every published `.nupkg` carries a signed **build provenance attestation** naming the workflow, the
commit and the runner that produced it. To check a package really came from this repository's CI
rather than someone's laptop:

```bash
gh attestation verify Ank.DocToolkit.<version>.nupkg --repo Ank-KhoaHo/DocToolkit
```

Attestation is produced immediately before the push, so the bytes that are verified are the bytes
that were published.

**The packages are not code-signed.** Authenticode signing needs a code-signing certificate this
project does not hold; provenance attestation answers "did this come from that source" without one,
which is the question a consumer of an open-source package usually has.

## Usage

```csharp
using DocToolkit;

byte[] docx = await HtmlToDocxConverter.ConvertAsync("<h1>Invoice</h1><p>Total: 18,100.00</p>");
byte[] pdf  = await HtmlToPdfConverter.ConvertAsync("<h1>Invoice</h1>");  // pivots through DOCX
byte[] rendered = DocxToPdfConverter.Convert(docx);

// Or build a DOCX from data rather than markup — no HTML to escape, so a value
// containing '<' cannot corrupt the document's structure
byte[] report = DocxEditor.Create(new[]
{
    DocxBlock.Heading("Quarterly Report", 1),
    DocxBlock.Paragraph("Revenue rose 12% against a flat cost base."),
    DocxBlock.Table(
        new[] { "Region", "Revenue" },
        new[] { new object[] { "EMEA", 1200 }, new object[] { "APAC", 980 } }),
});

// Fill a template — handles placeholders split across runs, and headers/footers
byte[] filled = DocxEditor.ReplaceText(docx, new() { ["{{customer}}"] = "Contoso Ltd" });

// Repeat a table row per record — a row holding {{item.Desc}} becomes one row per line item,
// each keeping the template row's formatting
byte[] invoice = DocxEditor.FillRows(filled, "item", lineItems);

// Drop an image into a placeholder — sized from its own header, PNG or JPEG
invoice = DocxEditor.ReplaceImage(invoice, "{{logo}}", logoBytes);

// Or work directly with files — no ReadAllBytes/WriteAllBytes dance, and input/output may be the same file
await DocxEditor.ReplaceTextAsync("invoice.docx", "invoice.docx", new() { ["{{customer}}"] = "Contoso Ltd" });

byte[] xlsx = WorkbookEditor.Create("Sales", new[] { new object?[] { "Region", "Total" } });

// Or build a multi-sheet workbook from data, with a formula cell computing across sheets
byte[] workbook = WorkbookEditor.Create(new[]
{
    XlsxSheet.Named("Sales", new[] { new object?[] { "Region", "Total" }, new object?[] { "EMEA", 1200 } }),
    XlsxSheet.Named("Summary", new[] { new object?[] { "Grand total", XlsxFormula.From("SUM(Sales!B2:B2)") } }),
});

// Read one back without knowing its shape in advance
IReadOnlyList<string> sheets = WorkbookEditor.SheetNames(xlsx);
IReadOnlyList<IReadOnlyList<string>> grid = WorkbookEditor.ReadSheet(xlsx, sheets[0]);
IReadOnlyList<string> slideText = PresentationEditor.ExtractText(pptx);

// Or build a deck from data - titles and bullets, no template needed
byte[] deck = PresentationEditor.Create(new[]
{
    PptxSlide.Titled("Q3 Results", "Revenue up 12%", "Costs flat"),
    PptxSlide.Titled("Outlook", "Hiring 3 engineers"),
});
```

**Formulas carry no cached value.** A cell written with `XlsxFormula` holds the formula and nothing
else. Excel recalculates when it opens the file, and `ReadCell`/`ReadSheet` compute the value on
read — but a third-party reader that only reads cached values, such as openpyxl with
`data_only=True`, sees an empty cell until Excel has opened and saved the file. A formula that
cannot be evaluated reads back as its Excel error string (`#DIV/0!`, `#NAME?`, `#REF!`) rather than
throwing.

### Page setup

Generated documents are **A4 with one-inch margins**. Pass a `PageSetup` for anything else:

```csharp
byte[] pdf = await HtmlToPdfConverter.ConvertAsync(
    html,
    PageSetup.A4.Landscape().WithMargins(36));

byte[] docx = DocxEditor.Create(blocks, PageSetup.Letter);
```

`PageSetup` is immutable, measured in points, and offers `A4`, `Letter` and
`Custom(widthPoints, heightPoints)`, plus `Landscape()` and `WithMargins(…)`. Every producer —
`HtmlToDocxConverter`, `HtmlToPdfConverter` and `DocxEditor.Create` — takes one.

> **Changed in this release.** Documents produced before this stated **no page size at all**, so
> Word applied its Normal template — US Letter on a US install, A4 on most others — and the
> generated PDF was always US Letter regardless. The same content therefore printed on different
> paper depending on who opened it. Every producer now states its page setup explicitly, defaulting
> to A4. If you want the old PDF behaviour, pass `PageSetup.Letter`.

`DocxToPdfConverter` takes no `PageSetup`: it renders a document that already carries its own page
setup, and honours it.

Six static classes, each stateless and safe to call concurrently, with a `byte[]` overload, a
`Stream` overload and a file-path overload for every capability, wrapping failures in
`DocumentConversionException`. Full surface:
**[package README](src/DocToolkit/README.md)** · [API docs](https://ank-khoaho.github.io/DocToolkit/).

## Dependency injection

`Ank.DocToolkit` needs no container. For ASP.NET Core or worker services, a thin companion package
adds injectable interfaces that delegate one-for-one to the static API — same conversion logic,
no duplication.

```csharp
// dotnet add package Ank.DocToolkit.Extensions.DependencyInjection
services.AddDocToolkit();                                       // six interfaces, as singletons

services.AddDocToolkit(o =>
{
    o.AllowRemoteImageDownload = true;                  // opt in to remote images...
    o.RemoteImage.AllowedHosts.Add("cdn.example.com");  // ...bounded; defaults are restrictive
});

public class InvoiceService(IHtmlToPdfConverter toPdf)
{
    public Task<byte[]> RenderAsync(string html) => toPdf.ConvertAsync(html);
}
```

Both packages ship at the same version from the same tag. See the
[extension package's README](src/DocToolkit.Extensions.DependencyInjection/README.md).

## Offline by default

No default code path performs network I/O — enforced, not merely intended. 37 guard tests assert
**zero** socket connections across the whole public API, against markup naming a loopback listener
sixteen ways (`<img src>`, `srcset`, `<link rel=stylesheet>`, `@import`, `background-image`,
`<iframe>`, `<object>`, `<script>`). The guard is proved by mutation: enabling downloads turns nine
of those tests red with real request lines, so it discriminates rather than passing vacuously.

The one exception is explicit and opt-in, and **is silently image-less in an air-gapped
environment** — an unreachable host is skipped, not fatal:

```csharp
await HtmlToDocxConverter.ConvertAsync(html, allowRemoteImageDownload: true);

// Bounded instead of wide open: timeout, byte cap, host allow-list and a block on
// loopback/private/link-local addresses, all on by default. Not a complete SSRF defence — see
// the package README.
await HtmlToDocxConverter.ConvertAsync(html, new RemoteImageOptions());
```

## Reporting a vulnerability

Privately, through GitHub — [open an advisory](https://github.com/Ank-KhoaHo/DocToolkit/security/advisories/new),
not a public issue. [`SECURITY.md`](SECURITY.md) says what is in scope, and lists the limits that
are already documented and therefore aren't findings on their own — chief among them that the
remote-image guard above is **not a complete SSRF defence**.

## Design notes

**HTML → PDF goes through DOCX.** No permissively-licensed, NuGet-only, Linux-safe library renders
HTML to PDF directly — the only free renderers *are* browsers, and a browser is a native binary.
Pivoting through DOCX keeps the whole chain pure managed.

**ShapeCrawler was removed.** PPTX originally used it, until it turned out to pull SkiaSharp and
Magick.NET: 38 native `.so`/`.dylib` files, 664 MB of `runtimes/`, and 26 CVE advisories. PPTX now
sits directly on `DocumentFormat.OpenXml`. The lesson — checking a library's *API* tells you
nothing about what it drags in — is why the CI guard exists.

**`SixLabors.Fonts` is pinned to `[1.0.0]`.** Version 2.x switches to the Six Labors Split License,
Apache-2.0 only under $1M annual revenue. CI asserts the pin holds, so a feed carrying only 2.x
fails restore loudly rather than silently relicensing you.

## Dependencies

Direct: `DocumentFormat.OpenXml` · `HtmlToOpenXml.dll` · `OfficeIMO.Word.Pdf` · `ClosedXML` ·
`SixLabors.Fonts [1.0.1]`. Full closure is 16 packages — 15 MIT, 1 Apache-2.0; see
[`THIRD-PARTY-NOTICES.txt`](src/DocToolkit/THIRD-PARTY-NOTICES.txt).

**Mirroring to a private feed?** Four things catch people out: `System.IO.Packaging` resolves to
**two** versions (8.0.1 for net8.0, 10.0.2 for net10.0); `Microsoft.Extensions.Logging.Abstractions`
is pinned at **6.0.0** by OfficeIMO; the package is **`RBush.Signed`**, not `RBush`; and
`SixLabors.Fonts` must be mirrored at exactly 1.0.0.

## Build and test

```bash
dotnet build DocToolkit.sln -c Release
dotnet test  DocToolkit.sln -c Release      # 426 tests x 2 target frameworks = 852 results

docker build -f Dockerfile.linux-test -t doctoolkit-linux-test .   # verify Linux locally
docker run --rm doctoolkit-linux-test

# samples reference the *published* packages, not this source - the restore a consumer gets
dotnet run --project samples/HtmlConversion      # one folder per capability - see samples/README.md
dotnet run --project samples/MinimalApi
```

Both packages ship at one version, from a single tag.
[release-please](https://github.com/googleapis/release-please) keeps a Release PR up to date as
commits land on `main`; merging it is a deliberate, manual decision, and
[`release.yml`](.github/workflows/release.yml) then publishes, authenticated with
[Trusted Publishing](https://learn.microsoft.com/en-us/nuget/nuget-org/trusted-publishing) — no
stored API key. Releases are cut by a maintainer merging a release pull request by hand;
nothing else publishes.

## Repository layout

```
src/DocToolkit/                                         the library
src/DocToolkit.Extensions.DependencyInjection/          DI extensions package
tests/                                                  426 tests, including the public-API approval guard, Stream-overload proofs and the air-gap/dependency guards
samples/                                                six runnable samples, each answering one question, on the published packages
docfx/                                                  API docs source, published to GitHub Pages on release
```

## Licence

MIT — see [LICENSE](LICENSE).
