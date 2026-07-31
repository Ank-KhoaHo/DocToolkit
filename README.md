# DocToolkit

[![CI](https://github.com/Ank-KhoaHo/DocToolkit/actions/workflows/ci.yml/badge.svg)](https://github.com/Ank-KhoaHo/DocToolkit/actions/workflows/ci.yml)
[![NuGet](https://img.shields.io/nuget/v/Ank.DocToolkit.svg?label=Ank.DocToolkit)](https://www.nuget.org/packages/Ank.DocToolkit/)
[![NuGet](https://img.shields.io/nuget/v/Ank.DocToolkit.Extensions.DependencyInjection.svg?label=Ank.DocToolkit.Extensions.DependencyInjection)](https://www.nuget.org/packages/Ank.DocToolkit.Extensions.DependencyInjection/)

Convert **HTML → DOCX and PDF**, and open/edit **DOCX, XLSX and PPTX**, from .NET.

**Pure managed. No native binaries, no browser, no LibreOffice, no Office interop.**
Works after `dotnet restore` alone, runs on Linux, and makes **no network calls at runtime**.

```bash
dotnet add package Ank.DocToolkit
```

Targets `net8.0` and `net10.0`. MIT licensed. See [CHANGELOG.md](CHANGELOG.md) for release notes.

## Why this exists

Most .NET document stacks fail at least one of these. This one is built to satisfy all four:

| Constraint | How |
|---|---|
| **Free for commercial use** | 19 dependencies: 18 MIT, 1 Apache-2.0. No revenue thresholds, no per-seat fees. |
| **NuGet only** | No Chromium download, no LibreOffice install, no native binaries. |
| **Runs on Linux** | Verified in CI on `ubuntu-24.04`, not inferred. |
| **Works offline** | No runtime network I/O. Proven by 35 air-gap tests. |

Each is re-checked by CI on every push, because all four are properties of the *resolved
dependency graph* — a single upstream bump can break them silently. That happened once already
(see [Design notes](#design-notes)).

## Usage

```csharp
using DocToolkit;

// HTML -> DOCX
byte[] docx = await HtmlToDocxConverter.ConvertAsync("<h1>Invoice</h1><p>Total: 18,100.00</p>");

// HTML -> PDF  (pivots through DOCX internally)
byte[] pdf = await HtmlToPdfConverter.ConvertAsync("<h1>Invoice</h1><p>Total: 18,100.00</p>");

// DOCX -> PDF
byte[] rendered = DocxToPdfConverter.Convert(docx);

// Fill a DOCX template (handles placeholders split across runs, and headers/footers)
byte[] filled = DocxEditor.ReplaceText(docx, new Dictionary<string, string>
{
    ["{{customer}}"] = "Contoso Ltd",
});
string text = DocxEditor.ExtractText(filled);

// Spreadsheets
byte[] xlsx = WorkbookEditor.Create("Sales", new[]
{
    new object?[] { "Region", "Total" },
    new object?[] { "North", 1200 },
});
string cell = WorkbookEditor.ReadCell(xlsx, "Sales", "B2");
byte[] updated = WorkbookEditor.SetCell(xlsx, "Sales", "B2", 1500);

// Presentations
int slides = PresentationEditor.SlideCount(pptx);
IReadOnlyList<string> slideText = PresentationEditor.ExtractText(pptx);
byte[] edited = PresentationEditor.ReplaceText(pptx, new Dictionary<string, string>
{
    ["{{title}}"] = "Q3 Results",
});
```

Every method is `byte[]` in / `byte[]` out and stateless, so the static classes are safe to call
concurrently. Failures are wrapped in `DocumentConversionException`.

## Dependency injection

This repo ships two packages. `Ank.DocToolkit` (`src/DocToolkit/`) is the library itself — static
classes, no DI container required. `Ank.DocToolkit.Extensions.DependencyInjection`
(`src/DocToolkit.Extensions.DependencyInjection/`) is a thin extension package that sits on top of
it: it adds injectable interfaces that each delegate one-for-one to the static API, without
changing or duplicating any conversion logic. A console app, Lambda or script installs just the
core package; an ASP.NET Core or worker-service app adds the extension alongside it.

```bash
dotnet add package Ank.DocToolkit.Extensions.DependencyInjection
```

```csharp
using DocToolkit.Extensions.DependencyInjection;

services.AddDocToolkit();
// or, to allow remote image download for HTML->DOCX/PDF (fails in an air-gapped environment):
services.AddDocToolkit(o => o.AllowRemoteImageDownload = true);
```

```csharp
public class InvoiceService(IHtmlToPdfConverter toPdf)
{
    public Task<byte[]> RenderAsync(string html) => toPdf.ConvertAsync(html);
}
```

`AddDocToolkit()` registers six interfaces as singletons — `IHtmlToDocxConverter`,
`IDocxToPdfConverter`, `IHtmlToPdfConverter`, `IDocxEditor`, `IWorkbookEditor`,
`IPresentationEditor` — one per static class above. See the
[extension package's own README](src/DocToolkit.Extensions.DependencyInjection/README.md) for
details, and Releasing below — both packages ship at the same version, from the same tag.

## Offline by default

The package performs **no network I/O** on any default code path. This is enforced, not intended:
35 guard tests assert **zero** socket connections across the whole public API, against markup that
names a loopback listener sixteen different ways (`<img src>`, `srcset`, `<link rel=stylesheet>`,
`@import`, `background-image`, table-cell images, `<iframe>`, `<object>`, `<script>`).

The guard is proved by mutation — enabling downloads turns nine of those tests red with real
request lines — so it genuinely discriminates rather than passing vacuously.

The one exception is explicit and opt-in:

```csharp
// Issues outbound HTTP. FAILS in an air-gapped environment.
await HtmlToDocxConverter.ConvertAsync(html, allowRemoteImageDownload: true);
```

## Design notes

**HTML → PDF goes through DOCX.** No permissively-licensed, NuGet-only, Linux-safe library renders
HTML to PDF directly — the only free renderers *are* browsers, and a browser is a native binary.
Pivoting through DOCX keeps the whole chain pure managed.

**ShapeCrawler was removed.** PPTX originally used it, until it turned out to pull SkiaSharp and
Magick.NET: 38 native `.so`/`.dylib` files, 664 MB of `runtimes/`, and 26 CVE advisories. PPTX is
now implemented directly on `DocumentFormat.OpenXml`. The lesson — checking a library's *API* tells
you nothing about what it drags in — is why the CI guard exists.

**`SixLabors.Fonts` is pinned to `[1.0.0]`.** Version 2.x switches to the Six Labors Split License,
which is Apache-2.0 only under $1M annual revenue. CI asserts the pin holds; a feed that only
carries 2.x will fail restore loudly rather than silently relicensing you.

## Dependencies

Direct: `DocumentFormat.OpenXml` · `HtmlToOpenXml.dll` · `OfficeIMO.Word.Pdf` · `ClosedXML` ·
`SixLabors.Fonts [1.0.0]`

Full closure is 19 packages — 18 MIT, 1 Apache-2.0. See
[`THIRD-PARTY-NOTICES.txt`](src/DocToolkit/THIRD-PARTY-NOTICES.txt) for attribution and the
resolved versions.

**Mirroring to a private feed?** Four things that catch people out: `System.IO.Packaging` resolves
to **two** versions (8.0.1 for net8.0, 10.0.2 for net10.0); `Microsoft.Extensions.Logging.Abstractions`
is pinned at **6.0.0** by OfficeIMO; the package is **`RBush.Signed`**, not `RBush`; and
`SixLabors.Fonts` must be mirrored at exactly 1.0.0.

## Build and test

```bash
dotnet build DocToolkit.sln -c Release
dotnet test  DocToolkit.sln -c Release      # 205 tests x 2 target frameworks = 410 results
dotnet pack  src/DocToolkit/DocToolkit.csproj -c Release
dotnet pack  src/DocToolkit.Extensions.DependencyInjection -c Release
```

Verify the Linux story locally with Docker:

```bash
docker build -f Dockerfile.linux-test -t doctoolkit-linux-test .
docker run --rm doctoolkit-linux-test
```

## Releasing

Publishing is driven by tags, and **one tag releases both packages at the same version.**
`Ank.DocToolkit` and `Ank.DocToolkit.Extensions.DependencyInjection` are meant to stay in
lockstep — the tag is the single source of truth for the version; the `<Version>` in each csproj
is only a local dev default.

```bash
git tag v1.0.0
git push origin v1.0.0
```

That runs [`.github/workflows/release.yml`](.github/workflows/release.yml), which **re-proves
everything before it pushes**, because a publish to nuget.org is irreversible — a version can be
unlisted but never deleted or replaced:

1. [`CHANGELOG.md`](CHANGELOG.md) has a `## [X.Y.Z]` entry for the tag's version
2. full build with `-warnaserror`, then the whole test suite on both target frameworks
3. the three guards: no native binaries, no banned packages, `SixLabors.Fonts` still on 1.x —
   checked against both projects
4. pack both projects at the tag's version, then verify each `.nupkg` (both TFMs, metadata,
   deps, MIT, no `runtimes/` payload, version matches the tag)
5. one OIDC exchange, then push both packages to nuget.org (`--skip-duplicate`, so a version
   already published for one of them is never an error) and create a GitHub Release with
   generated notes and both packages attached

A release that would break either package's own premise fails instead of shipping.

**Authentication is keyless.** Publishing uses [Trusted Publishing](https://learn.microsoft.com/en-us/nuget/nuget-org/trusted-publishing):
GitHub mints a short-lived, signed OIDC token describing this repo and workflow, nuget.org
validates it against a policy you configured, and hands back an API key valid for one hour. **No
long-lived key is stored anywhere** — nothing to leak, expire, or rotate.

**Setup, once:**
1. On nuget.org → *your name → Trusted Publishing* → add a policy:
   Repository Owner `Ank-KhoaHo`, Repository `DocToolkit`, Workflow File `release.yml`
   (filename only, no path).
2. Add `NUGET_USER` as a repository variable holding your nuget.org **profile name** — not your
   email address. The workflow fails early with a clear message if it is missing.

**Prereleases** work as expected — `v1.2.3-beta.1` is detected and marked pre-release on GitHub.

**Dry run:** trigger the workflow manually from the Actions tab with *Run workflow*, give it a
version, and leave *publish* unticked. It packs and verifies both packages without pushing
anything.

## Repository layout

```
src/DocToolkit/                                         the library
tests/DocToolkit.Tests/                                 182 tests, including Stream-overload proofs and the air-gap/dependency guards
src/DocToolkit.Extensions.DependencyInjection/          DI extensions package (services.AddDocToolkit())
tests/DocToolkit.Extensions.DependencyInjection.Tests/  23 tests for the DI extensions package
spike/                                                  the original proof-of-concept, kept as reference
docs/                                                   the implementation plan this was built from
```

The research behind the library selection lives in a separate repository,
[AutoLnD](https://github.com/Ank-KhoaHo/AutoLnD), under `learning-docs/dotnet-doc-libs/`.

## Licence

MIT — see [LICENSE](LICENSE).
