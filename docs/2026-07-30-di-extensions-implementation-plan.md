# DI Extensions Package Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Ship `Ank.DocToolkit.Extensions.DependencyInjection`, a new NuGet package adding `services.AddDocToolkit()` — six injectable interfaces wrapping the existing static API — for ASP.NET Core and worker-service consumers, without changing anything already published in `Ank.DocToolkit` 0.1.0.

**Architecture:** A new project, `src/DocToolkit.Extensions.DependencyInjection/`, referencing the published `Ank.DocToolkit` package (not a project reference — the extensions package must pull the core one transitively, the same way any external consumer would). One interface per existing static class, each backed by a thin `internal sealed` service that delegates to the static method, differing only in that `HtmlToDocxConverter`/`HtmlToPdfConverter`'s `allowRemoteImageDownload` parameter becomes a registration-time `DocToolkitOptions.AllowRemoteImageDownload` setting instead of a per-call argument.

**Tech Stack:** .NET 8 / .NET 10 · `Microsoft.Extensions.DependencyInjection.Abstractions` · `Microsoft.Extensions.Options` · xUnit · the already-published `Ank.DocToolkit` 0.1.0

**Source design:** [`docs/2026-07-30-di-extensions-design.md`](2026-07-30-di-extensions-design.md) — read it for the *why*; this plan is the *how*.

## Global Constraints

- **Target frameworks `net8.0;net10.0`** for both the new library and its test project, matching the core package exactly. Every test runs once per framework, so *N* tests report *2N* results.
- **`Ank.DocToolkit` is a `PackageReference`, never a `ProjectReference`.** The new project must restore it from nuget.org like any other consumer. Verified available: `Ank.DocToolkit` 0.1.0 is live and public.
- **Interfaces are `byte[]`-only**, mirroring the *current published* 0.1.0 API — no `Stream` overloads on these interfaces in this plan. (Confirmed: every byte[] method these interfaces need, including `DocxEditor.ExtractText(byte[], bool)`, already exists in the published 0.1.0 package — verified against the `v0.1.0` git tag directly, not assumed.)
- **Service implementations are `internal sealed`** — never exposed as public types. Consumers depend on the interfaces only.
- **Never add a NuGet package** beyond `Ank.DocToolkit`, `Microsoft.Extensions.DependencyInjection.Abstractions`, and `Microsoft.Extensions.Options` (library), or `Microsoft.Extensions.DependencyInjection` (full, test project only — needed for `ServiceCollection`/`BuildServiceProvider`, which `.Abstractions` does not provide).
- **Never introduce** `System.Drawing.Common`, `SkiaSharp`, `Magick.NET*`, `ShapeCrawler`, `EPPlus`, `NPOI`, `Spire.*`, `Syncfusion.*`, `QuestPDF`, `IronPDF`.
- **Do not replace an existing `.csproj` wholesale once it carries real content** — edit in place with `dotnet add package`. (This only applies from Task 2 onward; Task 1 creates the new csproj from nothing, so there is nothing to preserve yet.)
- **Commit messages must NOT contain a `Co-Authored-By` trailer.**
- **Build stays at `-warnaserror`, 0 warnings.**
- Fully qualify references to the core `DocToolkit` namespace's types as `DocToolkit.HtmlToDocxConverter`, `DocToolkit.DocumentConversionException`, etc. **inside service implementation files** — those files live in `namespace DocToolkit.Extensions.DependencyInjection`, which also declares `IHtmlToDocxConverter` etc., so an unqualified `HtmlToDocxConverter` next to `IHtmlToDocxConverter` in the same file is needlessly easy to misread. Test files may use plain `using DocToolkit;` since they don't declare colliding names.

---

## File Structure

```
DocToolkit/
├── DocToolkit.sln                                          (modified: 2 new projects added)
├── .github/workflows/
│   ├── ci.yml                                              (modified: cover both projects)
│   └── release-extensions.yml                              (new: tag-driven release for this package)
├── src/
│   ├── DocToolkit/                                          UNCHANGED
│   └── DocToolkit.Extensions.DependencyInjection/           NEW
│       ├── DocToolkit.Extensions.DependencyInjection.csproj
│       ├── README.md
│       ├── THIRD-PARTY-NOTICES.txt
│       ├── DocToolkitOptions.cs
│       ├── ServiceCollectionExtensions.cs
│       ├── IHtmlToDocxConverter.cs
│       ├── HtmlToDocxConverterService.cs
│       ├── IDocxToPdfConverter.cs
│       ├── DocxToPdfConverterService.cs
│       ├── IHtmlToPdfConverter.cs
│       ├── HtmlToPdfConverterService.cs
│       ├── IDocxEditor.cs
│       ├── DocxEditorService.cs
│       ├── IWorkbookEditor.cs
│       ├── WorkbookEditorService.cs
│       ├── IPresentationEditor.cs
│       └── PresentationEditorService.cs
└── tests/
    ├── DocToolkit.Tests/                                    UNCHANGED
    └── DocToolkit.Extensions.DependencyInjection.Tests/      NEW
        ├── DocToolkit.Extensions.DependencyInjection.Tests.csproj
        ├── DocToolkitOptionsTests.cs
        ├── HtmlToDocxConverterServiceTests.cs
        ├── DocxToPdfConverterServiceTests.cs
        ├── HtmlToPdfConverterServiceTests.cs
        ├── DocxEditorServiceTests.cs
        ├── WorkbookEditorServiceTests.cs
        ├── PresentationEditorServiceTests.cs
        ├── ServiceCollectionExtensionsTests.cs
        └── LoopbackProbe.cs
```

One interface + one service per file, matching the core project's one-class-per-file convention. Tests are one file per interface, plus one for the registration extension.

---

### Task 1: Scaffold the project + DocToolkitOptions + IHtmlToDocxConverter

**Files:**
- Create: `src/DocToolkit.Extensions.DependencyInjection/DocToolkit.Extensions.DependencyInjection.csproj`
- Create: `src/DocToolkit.Extensions.DependencyInjection/README.md`
- Create: `src/DocToolkit.Extensions.DependencyInjection/THIRD-PARTY-NOTICES.txt`
- Create: `src/DocToolkit.Extensions.DependencyInjection/DocToolkitOptions.cs`
- Create: `src/DocToolkit.Extensions.DependencyInjection/IHtmlToDocxConverter.cs`
- Create: `src/DocToolkit.Extensions.DependencyInjection/HtmlToDocxConverterService.cs`
- Create: `tests/DocToolkit.Extensions.DependencyInjection.Tests/DocToolkit.Extensions.DependencyInjection.Tests.csproj`
- Create: `tests/DocToolkit.Extensions.DependencyInjection.Tests/DocToolkitOptionsTests.cs`
- Create: `tests/DocToolkit.Extensions.DependencyInjection.Tests/HtmlToDocxConverterServiceTests.cs`
- Modify: `DocToolkit.sln`

**Interfaces:**
- Consumes: `DocToolkit.HtmlToDocxConverter.ConvertAsync(string html, bool allowRemoteImageDownload, CancellationToken ct = default) -> Task<byte[]>` (published in `Ank.DocToolkit` 0.1.0), `DocToolkit.DocxEditor.ExtractText(byte[] docx) -> string` (for test content verification).
- Produces: `DocToolkit.Extensions.DependencyInjection.DocToolkitOptions` (public, settable `bool AllowRemoteImageDownload { get; set; }`, default `false`), `DocToolkit.Extensions.DependencyInjection.IHtmlToDocxConverter` with `Task<byte[]> ConvertAsync(string html, CancellationToken ct = default)`.

- [ ] **Step 1: Scaffold both projects and wire the solution**

```bash
cd /path/to/DocToolkit
dotnet new classlib -f net8.0 -o src/DocToolkit.Extensions.DependencyInjection -n DocToolkit.Extensions.DependencyInjection
rm src/DocToolkit.Extensions.DependencyInjection/Class1.cs
dotnet new xunit -f net8.0 -o tests/DocToolkit.Extensions.DependencyInjection.Tests -n DocToolkit.Extensions.DependencyInjection.Tests
rm tests/DocToolkit.Extensions.DependencyInjection.Tests/UnitTest1.cs
dotnet sln add src/DocToolkit.Extensions.DependencyInjection/DocToolkit.Extensions.DependencyInjection.csproj
dotnet sln add tests/DocToolkit.Extensions.DependencyInjection.Tests/DocToolkit.Extensions.DependencyInjection.Tests.csproj
dotnet add tests/DocToolkit.Extensions.DependencyInjection.Tests reference src/DocToolkit.Extensions.DependencyInjection
```

- [ ] **Step 1b: Multi-target the scaffolded test project**

`dotnet new xunit -f net8.0` only emits `net8.0` — the core project's own test project (`tests/DocToolkit.Tests`) is multi-targeted, and this one needs to match it exactly (same package versions), or `dotnet test DocToolkit.sln` won't exercise `net10.0` for these tests at all. Replace the generated `tests/DocToolkit.Extensions.DependencyInjection.Tests/DocToolkit.Extensions.DependencyInjection.Tests.csproj` in full:

```xml
<Project Sdk="Microsoft.NET.Sdk">

  <PropertyGroup>
    <TargetFrameworks>net8.0;net10.0</TargetFrameworks>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>

    <IsPackable>false</IsPackable>
    <IsTestProject>true</IsTestProject>
  </PropertyGroup>

  <ItemGroup>
    <PackageReference Include="coverlet.collector" Version="6.0.0" />
    <PackageReference Include="Microsoft.NET.Test.Sdk" Version="17.8.0" />
    <PackageReference Include="xunit" Version="2.5.3" />
    <PackageReference Include="xunit.runner.visualstudio" Version="2.5.3" />
  </ItemGroup>

  <ItemGroup>
    <Using Include="Xunit" />
  </ItemGroup>

  <ItemGroup>
    <ProjectReference Include="..\..\src\DocToolkit.Extensions.DependencyInjection\DocToolkit.Extensions.DependencyInjection.csproj" />
  </ItemGroup>

</Project>
```

- [ ] **Step 2: Write the final library csproj**

Replace the generated `src/DocToolkit.Extensions.DependencyInjection/DocToolkit.Extensions.DependencyInjection.csproj` in full — this is a brand-new file with nothing yet to preserve, unlike the core project's csproj:

```xml
<Project Sdk="Microsoft.NET.Sdk">

  <PropertyGroup>
    <TargetFrameworks>net8.0;net10.0</TargetFrameworks>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
    <LangVersion>latest</LangVersion>
    <GenerateDocumentationFile>true</GenerateDocumentationFile>
  </PropertyGroup>

  <PropertyGroup>
    <PackageId>Ank.DocToolkit.Extensions.DependencyInjection</PackageId>
    <Version>0.1.0</Version>
    <Authors>Khoa Ho</Authors>
    <Description>
      Dependency-injection registration for Ank.DocToolkit. services.AddDocToolkit() registers
      six injectable interfaces (IHtmlToDocxConverter, IDocxToPdfConverter, IHtmlToPdfConverter,
      IDocxEditor, IWorkbookEditor, IPresentationEditor) over the same pure-managed HTML/DOCX/
      PDF/XLSX/PPTX conversion and editing logic, for ASP.NET Core and worker-service consumers.
    </Description>
    <PackageTags>docx;pdf;xlsx;pptx;dependency-injection;aspnetcore;openxml</PackageTags>
    <PackageLicenseExpression>MIT</PackageLicenseExpression>
    <PackageProjectUrl>https://github.com/Ank-KhoaHo/DocToolkit</PackageProjectUrl>
    <PackageReadmeFile>README.md</PackageReadmeFile>
    <PackageOutputPath>$(MSBuildThisFileDirectory)..\..\artifacts</PackageOutputPath>
    <IncludeSymbols>true</IncludeSymbols>
    <SymbolPackageFormat>snupkg</SymbolPackageFormat>
    <Deterministic>true</Deterministic>
  </PropertyGroup>

  <ItemGroup>
    <None Include="README.md" Pack="true" PackagePath="\" />
    <None Include="THIRD-PARTY-NOTICES.txt" Pack="true" PackagePath="\" />
  </ItemGroup>

  <!--
    Service classes are `internal sealed` (consumers depend on the interfaces only), but this
    task's own tests construct them directly (`new HtmlToDocxConverterService(...)`) rather than
    only through DI - InternalsVisibleTo is what lets that compile. Set once, whole-assembly:
    later tasks' internal sealed services (Tasks 2-6) need no repeat of this.
  -->
  <ItemGroup>
    <InternalsVisibleTo Include="DocToolkit.Extensions.DependencyInjection.Tests" />
  </ItemGroup>

</Project>
```

- [ ] **Step 3: Add the package references**

```bash
dotnet add src/DocToolkit.Extensions.DependencyInjection package Ank.DocToolkit --version 0.1.0
dotnet add src/DocToolkit.Extensions.DependencyInjection package Microsoft.Extensions.DependencyInjection.Abstractions
dotnet add src/DocToolkit.Extensions.DependencyInjection package Microsoft.Extensions.Options
```

Then open the csproj and change the `Ank.DocToolkit` line from a pinned `Version="0.1.0"` to a floor, so a consumer isn't forced onto exactly 0.1.0 once newer core versions ship:

```xml
    <PackageReference Include="Ank.DocToolkit" Version="[0.1.0, )" />
```

- [ ] **Step 4: Verify restore and an empty build succeed**

Run: `dotnet build DocToolkit.sln -c Release -warnaserror`
Expected: **succeeds** (both new projects are currently empty aside from scaffolding).

- [ ] **Step 5: Write package README and third-party-notices stubs**

These are referenced by the csproj's `Pack` items, so the build fails without them. Full content comes in Task 9, once every interface exists to document.

`src/DocToolkit.Extensions.DependencyInjection/README.md`:
```markdown
# Ank.DocToolkit.Extensions.DependencyInjection

Dependency-injection registration for [Ank.DocToolkit](https://www.nuget.org/packages/Ank.DocToolkit).
Full usage docs land once every service is implemented — see the parent repository for now:
https://github.com/Ank-KhoaHo/DocToolkit
```

`src/DocToolkit.Extensions.DependencyInjection/THIRD-PARTY-NOTICES.txt`:
```text
Ank.DocToolkit.Extensions.DependencyInjection third-party notices
===================================================================
Full notices are added once every dependency is finalised (see Task 9 of the implementation plan).
```

- [ ] **Step 6: Write the failing test for DocToolkitOptions**

Create `tests/DocToolkit.Extensions.DependencyInjection.Tests/DocToolkitOptionsTests.cs`:

```csharp
using DocToolkit.Extensions.DependencyInjection;
using Xunit;

namespace DocToolkit.Extensions.DependencyInjection.Tests;

public class DocToolkitOptionsTests
{
    [Fact]
    public void AllowRemoteImageDownload_DefaultsToFalse()
    {
        var options = new DocToolkitOptions();

        Assert.False(options.AllowRemoteImageDownload);
    }

    [Fact]
    public void AllowRemoteImageDownload_IsSettable()
    {
        var options = new DocToolkitOptions { AllowRemoteImageDownload = true };

        Assert.True(options.AllowRemoteImageDownload);
    }
}
```

- [ ] **Step 7: Run to verify it fails**

Run: `dotnet test tests/DocToolkit.Extensions.DependencyInjection.Tests --filter FullyQualifiedName~DocToolkitOptionsTests`
Expected: **build failure** — `The type or namespace name 'DocToolkitOptions' could not be found`.

- [ ] **Step 8: Implement DocToolkitOptions**

Create `src/DocToolkit.Extensions.DependencyInjection/DocToolkitOptions.cs`:

```csharp
namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>Options controlling the services registered by <c>AddDocToolkit</c> (added in Task 7 of this plan).</summary>
public sealed class DocToolkitOptions
{
    /// <summary>
    /// When true, HTML-to-DOCX and HTML-to-PDF conversion download images referenced by absolute
    /// <c>http</c>/<c>https</c> URLs. This issues outbound network requests - do not enable it in
    /// an air-gapped environment. Default: <c>false</c>.
    /// </summary>
    public bool AllowRemoteImageDownload { get; set; }
}
```

- [ ] **Step 9: Run to verify it passes**

Run: `dotnet test tests/DocToolkit.Extensions.DependencyInjection.Tests --filter FullyQualifiedName~DocToolkitOptionsTests`
Expected: **PASS**, 4 results (2 tests × 2 TFMs).

- [ ] **Step 10: Write the failing test for IHtmlToDocxConverter/HtmlToDocxConverterService**

Create `tests/DocToolkit.Extensions.DependencyInjection.Tests/HtmlToDocxConverterServiceTests.cs`:

```csharp
using System.Linq;
using DocToolkit;
using DocToolkit.Extensions.DependencyInjection;
using Microsoft.Extensions.Options;
using Xunit;

namespace DocToolkit.Extensions.DependencyInjection.Tests;

public class HtmlToDocxConverterServiceTests
{
    [Fact]
    public async Task ConvertAsync_ProducesADocxContainingTheGivenContent()
    {
        var sut = new HtmlToDocxConverterService(Options.Create(new DocToolkitOptions()));

        var docx = await sut.ConvertAsync("<h1>Title</h1><p>Body copy.</p>");

        Assert.Equal(new byte[] { 0x50, 0x4B, 0x03, 0x04 }, docx.Take(4).ToArray());
        Assert.Contains("Body copy.", DocxEditor.ExtractText(docx));
    }

    [Fact]
    public async Task ConvertAsync_RejectsNullHtml()
    {
        var sut = new HtmlToDocxConverterService(Options.Create(new DocToolkitOptions()));

        await Assert.ThrowsAsync<ArgumentNullException>(() => sut.ConvertAsync(null!));
    }
}
```

- [ ] **Step 11: Run to verify it fails**

Run: `dotnet test tests/DocToolkit.Extensions.DependencyInjection.Tests --filter FullyQualifiedName~HtmlToDocxConverterServiceTests`
Expected: **build failure** — `HtmlToDocxConverterService` and `IHtmlToDocxConverter` do not exist.

- [ ] **Step 12: Implement IHtmlToDocxConverter and HtmlToDocxConverterService**

Create `src/DocToolkit.Extensions.DependencyInjection/IHtmlToDocxConverter.cs`:

```csharp
namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>
/// Converts HTML to a Word (.docx) package. Registered by
/// <c>AddDocToolkit</c> (added in Task 7 of this plan); remote image download is controlled
/// once, at registration, via <see cref="DocToolkitOptions.AllowRemoteImageDownload"/>.
/// </summary>
public interface IHtmlToDocxConverter
{
    /// <summary>Converts <paramref name="html"/> to the bytes of a .docx package.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="html"/> is null.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The HTML could not be converted.</exception>
    Task<byte[]> ConvertAsync(string html, CancellationToken ct = default);
}
```

Create `src/DocToolkit.Extensions.DependencyInjection/HtmlToDocxConverterService.cs`:

```csharp
using Microsoft.Extensions.Options;

namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>
/// Default <see cref="IHtmlToDocxConverter"/>, delegating to
/// <see cref="DocToolkit.HtmlToDocxConverter"/> with <see cref="DocToolkitOptions.AllowRemoteImageDownload"/>.
/// </summary>
internal sealed class HtmlToDocxConverterService : IHtmlToDocxConverter
{
    private readonly DocToolkitOptions _options;

    public HtmlToDocxConverterService(IOptions<DocToolkitOptions> options) => _options = options.Value;

    public Task<byte[]> ConvertAsync(string html, CancellationToken ct = default)
        => DocToolkit.HtmlToDocxConverter.ConvertAsync(html, _options.AllowRemoteImageDownload, ct);
}
```

- [ ] **Step 13: Run to verify it passes**

Run: `dotnet test tests/DocToolkit.Extensions.DependencyInjection.Tests --filter FullyQualifiedName~HtmlToDocxConverterServiceTests`
Expected: **PASS**, 4 results (2 tests × 2 TFMs).

- [ ] **Step 14: Run the whole new test project and confirm no warnings**

Run: `dotnet build DocToolkit.sln -c Release -warnaserror && dotnet test tests/DocToolkit.Extensions.DependencyInjection.Tests`
Expected: **0 warnings**, all tests pass (8 results: 4 `DocToolkitOptionsTests` + 4 `HtmlToDocxConverterServiceTests`).

- [ ] **Step 15: Commit**

```bash
git add DocToolkit.sln src/DocToolkit.Extensions.DependencyInjection tests/DocToolkit.Extensions.DependencyInjection.Tests
git commit -m "feat(di-extensions): scaffold the project and add IHtmlToDocxConverter"
```

---

### Task 2: IDocxToPdfConverter

**Files:**
- Create: `src/DocToolkit.Extensions.DependencyInjection/IDocxToPdfConverter.cs`
- Create: `src/DocToolkit.Extensions.DependencyInjection/DocxToPdfConverterService.cs`
- Create: `tests/DocToolkit.Extensions.DependencyInjection.Tests/DocxToPdfConverterServiceTests.cs`

**Interfaces:**
- Consumes: `DocToolkit.DocxToPdfConverter.Convert(byte[] docx) -> byte[]`, `DocToolkit.Extensions.DependencyInjection.IHtmlToDocxConverter` is not needed here — build the test's DOCX input via the plain static `DocToolkit.HtmlToDocxConverter.ConvertAsync`.
- Produces: `DocToolkit.Extensions.DependencyInjection.IDocxToPdfConverter` with `byte[] Convert(byte[] docx)`.

- [ ] **Step 1: Write the failing test**

Create `tests/DocToolkit.Extensions.DependencyInjection.Tests/DocxToPdfConverterServiceTests.cs`:

```csharp
using System.Linq;
using DocToolkit;
using DocToolkit.Extensions.DependencyInjection;
using Xunit;

namespace DocToolkit.Extensions.DependencyInjection.Tests;

public class DocxToPdfConverterServiceTests
{
    [Fact]
    public async Task Convert_ProducesAPdf()
    {
        var sut = new DocxToPdfConverterService();
        var docx = await HtmlToDocxConverter.ConvertAsync("<h1>Invoice</h1><p>Total: 18,100.00</p>");

        var pdf = sut.Convert(docx);

        Assert.Equal(new byte[] { 0x25, 0x50, 0x44, 0x46, 0x2D }, pdf.Take(5).ToArray()); // "%PDF-"
        Assert.True(pdf.Length > 200, $"expected a real PDF, got {pdf.Length} bytes");
    }

    [Fact]
    public void Convert_RejectsEmptyInput()
    {
        var sut = new DocxToPdfConverterService();

        Assert.Throws<ArgumentException>(() => sut.Convert(Array.Empty<byte>()));
    }
}
```

- [ ] **Step 2: Run to verify it fails**

Run: `dotnet test tests/DocToolkit.Extensions.DependencyInjection.Tests --filter FullyQualifiedName~DocxToPdfConverterServiceTests`
Expected: **build failure** — `DocxToPdfConverterService` does not exist.

- [ ] **Step 3: Implement**

Create `src/DocToolkit.Extensions.DependencyInjection/IDocxToPdfConverter.cs`:

```csharp
namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>
/// Renders a Word (.docx) package to PDF. Registered by
/// <c>AddDocToolkit</c> (added in Task 7 of this plan).
/// </summary>
public interface IDocxToPdfConverter
{
    /// <summary>Renders the .docx in <paramref name="docx"/> and returns PDF bytes.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="docx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The DOCX could not be rendered.</exception>
    byte[] Convert(byte[] docx);
}
```

Create `src/DocToolkit.Extensions.DependencyInjection/DocxToPdfConverterService.cs`:

```csharp
namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>Default <see cref="IDocxToPdfConverter"/>, delegating to <see cref="DocToolkit.DocxToPdfConverter"/>.</summary>
internal sealed class DocxToPdfConverterService : IDocxToPdfConverter
{
    public byte[] Convert(byte[] docx) => DocToolkit.DocxToPdfConverter.Convert(docx);
}
```

- [ ] **Step 4: Run to verify it passes**

Run: `dotnet test tests/DocToolkit.Extensions.DependencyInjection.Tests --filter FullyQualifiedName~DocxToPdfConverterServiceTests`
Expected: **PASS**, 4 results (2 tests × 2 TFMs).

- [ ] **Step 5: Commit**

```bash
git add src/DocToolkit.Extensions.DependencyInjection tests/DocToolkit.Extensions.DependencyInjection.Tests
git commit -m "feat(di-extensions): add IDocxToPdfConverter"
```

---

### Task 3: IHtmlToPdfConverter

**Files:**
- Create: `src/DocToolkit.Extensions.DependencyInjection/IHtmlToPdfConverter.cs`
- Create: `src/DocToolkit.Extensions.DependencyInjection/HtmlToPdfConverterService.cs`
- Create: `tests/DocToolkit.Extensions.DependencyInjection.Tests/HtmlToPdfConverterServiceTests.cs`

**Interfaces:**
- Consumes: `DocToolkit.HtmlToPdfConverter.ConvertAsync(string html, bool allowRemoteImageDownload, CancellationToken ct = default) -> Task<byte[]>`.
- Produces: `DocToolkit.Extensions.DependencyInjection.IHtmlToPdfConverter` with `Task<byte[]> ConvertAsync(string html, CancellationToken ct = default)`.

- [ ] **Step 1: Write the failing test**

Create `tests/DocToolkit.Extensions.DependencyInjection.Tests/HtmlToPdfConverterServiceTests.cs`:

```csharp
using System.Linq;
using DocToolkit.Extensions.DependencyInjection;
using Microsoft.Extensions.Options;
using Xunit;

namespace DocToolkit.Extensions.DependencyInjection.Tests;

public class HtmlToPdfConverterServiceTests
{
    [Fact]
    public async Task ConvertAsync_ProducesAPdf()
    {
        var sut = new HtmlToPdfConverterService(Options.Create(new DocToolkitOptions()));

        var pdf = await sut.ConvertAsync("<h1>Invoice</h1><p>Total: 18,100.00</p>");

        Assert.Equal(new byte[] { 0x25, 0x50, 0x44, 0x46, 0x2D }, pdf.Take(5).ToArray());
        Assert.True(pdf.Length > 200, $"expected a real PDF, got {pdf.Length} bytes");
    }

    [Fact]
    public async Task ConvertAsync_RejectsNullHtml()
    {
        var sut = new HtmlToPdfConverterService(Options.Create(new DocToolkitOptions()));

        await Assert.ThrowsAsync<ArgumentNullException>(() => sut.ConvertAsync(null!));
    }
}
```

- [ ] **Step 2: Run to verify it fails**

Run: `dotnet test tests/DocToolkit.Extensions.DependencyInjection.Tests --filter FullyQualifiedName~HtmlToPdfConverterServiceTests`
Expected: **build failure** — `HtmlToPdfConverterService` does not exist.

- [ ] **Step 3: Implement**

Create `src/DocToolkit.Extensions.DependencyInjection/IHtmlToPdfConverter.cs`:

```csharp
namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>
/// Converts HTML straight to PDF by pivoting through DOCX. Registered by
/// <c>AddDocToolkit</c> (added in Task 7 of this plan); remote image download is controlled
/// once, at registration, via <see cref="DocToolkitOptions.AllowRemoteImageDownload"/>.
/// </summary>
public interface IHtmlToPdfConverter
{
    /// <summary>Converts <paramref name="html"/> straight to PDF bytes.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="html"/> is null.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The HTML could not be converted.</exception>
    Task<byte[]> ConvertAsync(string html, CancellationToken ct = default);
}
```

Create `src/DocToolkit.Extensions.DependencyInjection/HtmlToPdfConverterService.cs`:

```csharp
using Microsoft.Extensions.Options;

namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>
/// Default <see cref="IHtmlToPdfConverter"/>, delegating to
/// <see cref="DocToolkit.HtmlToPdfConverter"/> with <see cref="DocToolkitOptions.AllowRemoteImageDownload"/>.
/// </summary>
internal sealed class HtmlToPdfConverterService : IHtmlToPdfConverter
{
    private readonly DocToolkitOptions _options;

    public HtmlToPdfConverterService(IOptions<DocToolkitOptions> options) => _options = options.Value;

    public Task<byte[]> ConvertAsync(string html, CancellationToken ct = default)
        => DocToolkit.HtmlToPdfConverter.ConvertAsync(html, _options.AllowRemoteImageDownload, ct);
}
```

- [ ] **Step 4: Run to verify it passes**

Run: `dotnet test tests/DocToolkit.Extensions.DependencyInjection.Tests --filter FullyQualifiedName~HtmlToPdfConverterServiceTests`
Expected: **PASS**, 4 results (2 tests × 2 TFMs).

- [ ] **Step 5: Commit**

```bash
git add src/DocToolkit.Extensions.DependencyInjection tests/DocToolkit.Extensions.DependencyInjection.Tests
git commit -m "feat(di-extensions): add IHtmlToPdfConverter"
```

---

### Task 4: IDocxEditor

**Files:**
- Create: `src/DocToolkit.Extensions.DependencyInjection/IDocxEditor.cs`
- Create: `src/DocToolkit.Extensions.DependencyInjection/DocxEditorService.cs`
- Create: `tests/DocToolkit.Extensions.DependencyInjection.Tests/DocxEditorServiceTests.cs`

**Interfaces:**
- Consumes: `DocToolkit.DocxEditor.ReplaceText(byte[] docx, IReadOnlyDictionary<string, string> replacements) -> byte[]`, `DocToolkit.DocxEditor.ExtractText(byte[] docx) -> string`, `DocToolkit.DocxEditor.ExtractText(byte[] docx, bool includeHeadersAndFooters) -> string`.
- Produces: `DocToolkit.Extensions.DependencyInjection.IDocxEditor` with `byte[] ReplaceText(byte[] docx, IReadOnlyDictionary<string, string> replacements)`, `string ExtractText(byte[] docx)`, `string ExtractText(byte[] docx, bool includeHeadersAndFooters)`.

**Note — a small, deliberate completion of the design:** the approved design's `IDocxEditor` listed only `ExtractText(byte[])`, but §4 of the design states the goal is to mirror each static class "1:1", and the static `DocxEditor` has always had a second `ExtractText(byte[], bool includeHeadersAndFooters)` overload too (confirmed present in the published 0.1.0). Add both overloads here so the interface actually mirrors the static class it wraps.

- [ ] **Step 1: Write the failing test**

Create `tests/DocToolkit.Extensions.DependencyInjection.Tests/DocxEditorServiceTests.cs`:

```csharp
using System.Collections.Generic;
using DocToolkit;
using DocToolkit.Extensions.DependencyInjection;
using Xunit;

namespace DocToolkit.Extensions.DependencyInjection.Tests;

public class DocxEditorServiceTests
{
    [Fact]
    public async Task ReplaceText_SubstitutesPlaceholders()
    {
        var sut = new DocxEditorService();
        var docx = await HtmlToDocxConverter.ConvertAsync(
            "<p>Dear {{name}}, your balance is {{balance}}.</p>");

        var edited = sut.ReplaceText(docx, new Dictionary<string, string>
        {
            ["{{name}}"] = "Contoso Ltd",
            ["{{balance}}"] = "4,250.00",
        });

        var text = sut.ExtractText(edited);
        Assert.Contains("Contoso Ltd", text);
        Assert.Contains("4,250.00", text);
        Assert.DoesNotContain("{{name}}", text);
    }

    [Fact]
    public async Task ExtractText_WithHeadersAndFooters_MatchesTheStaticMethod()
    {
        var docx = await HtmlToDocxConverter.ConvertAsync("<p>Body text.</p>");
        var sut = new DocxEditorService();

        Assert.Equal(
            DocxEditor.ExtractText(docx, includeHeadersAndFooters: true),
            sut.ExtractText(docx, includeHeadersAndFooters: true));
    }

    [Fact]
    public void ReplaceText_RejectsNullReplacements()
    {
        var sut = new DocxEditorService();

        Assert.Throws<ArgumentNullException>(() => sut.ReplaceText(Array.Empty<byte>(), null!));
    }
}
```

- [ ] **Step 2: Run to verify it fails**

Run: `dotnet test tests/DocToolkit.Extensions.DependencyInjection.Tests --filter FullyQualifiedName~DocxEditorServiceTests`
Expected: **build failure** — `DocxEditorService` does not exist.

- [ ] **Step 3: Implement**

Create `src/DocToolkit.Extensions.DependencyInjection/IDocxEditor.cs`:

```csharp
namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>Opens and edits an existing .docx package. Registered by <c>AddDocToolkit</c> (added in Task 7 of this plan).</summary>
public interface IDocxEditor
{
    /// <summary>Replaces every key with its value across the document body, headers, footers, footnotes and endnotes.</summary>
    /// <exception cref="ArgumentNullException">Either argument is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The package could not be opened or edited.</exception>
    byte[] ReplaceText(byte[] docx, IReadOnlyDictionary<string, string> replacements);

    /// <summary>Returns the plain text of the document body. Headers, footers, footnotes and endnotes are not included.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="docx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The package could not be opened or read.</exception>
    string ExtractText(byte[] docx);

    /// <summary>Returns the plain text of the document. When <paramref name="includeHeadersAndFooters"/> is true, headers and footers follow the body text.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="docx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The package could not be opened or read.</exception>
    string ExtractText(byte[] docx, bool includeHeadersAndFooters);
}
```

Create `src/DocToolkit.Extensions.DependencyInjection/DocxEditorService.cs`:

```csharp
namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>Default <see cref="IDocxEditor"/>, delegating to <see cref="DocToolkit.DocxEditor"/>.</summary>
internal sealed class DocxEditorService : IDocxEditor
{
    public byte[] ReplaceText(byte[] docx, IReadOnlyDictionary<string, string> replacements)
        => DocToolkit.DocxEditor.ReplaceText(docx, replacements);

    public string ExtractText(byte[] docx) => DocToolkit.DocxEditor.ExtractText(docx);

    public string ExtractText(byte[] docx, bool includeHeadersAndFooters)
        => DocToolkit.DocxEditor.ExtractText(docx, includeHeadersAndFooters);
}
```

- [ ] **Step 4: Run to verify it passes**

Run: `dotnet test tests/DocToolkit.Extensions.DependencyInjection.Tests --filter FullyQualifiedName~DocxEditorServiceTests`
Expected: **PASS**, 6 results (3 tests × 2 TFMs).

- [ ] **Step 5: Commit**

```bash
git add src/DocToolkit.Extensions.DependencyInjection tests/DocToolkit.Extensions.DependencyInjection.Tests
git commit -m "feat(di-extensions): add IDocxEditor"
```

---

### Task 5: IWorkbookEditor

**Files:**
- Create: `src/DocToolkit.Extensions.DependencyInjection/IWorkbookEditor.cs`
- Create: `src/DocToolkit.Extensions.DependencyInjection/WorkbookEditorService.cs`
- Create: `tests/DocToolkit.Extensions.DependencyInjection.Tests/WorkbookEditorServiceTests.cs`

**Interfaces:**
- Consumes: `DocToolkit.WorkbookEditor.Create(string sheetName, IEnumerable<IEnumerable<object?>> rows) -> byte[]`, `.ReadCell(byte[] xlsx, string sheetName, string cellRef) -> string`, `.SetCell(byte[] xlsx, string sheetName, string cellRef, object? value) -> byte[]`.
- Produces: `DocToolkit.Extensions.DependencyInjection.IWorkbookEditor` mirroring those three signatures.

- [ ] **Step 1: Write the failing test**

Create `tests/DocToolkit.Extensions.DependencyInjection.Tests/WorkbookEditorServiceTests.cs`:

```csharp
using System.Linq;
using DocToolkit.Extensions.DependencyInjection;
using Xunit;

namespace DocToolkit.Extensions.DependencyInjection.Tests;

public class WorkbookEditorServiceTests
{
    [Fact]
    public void Create_ReadCell_SetCell_RoundTripCorrectly()
    {
        var sut = new WorkbookEditorService();

        var xlsx = sut.Create("Sales", new object?[][]
        {
            new object?[] { "Region", "Total" },
            new object?[] { "North", 1200 },
        });

        Assert.Equal(new byte[] { 0x50, 0x4B, 0x03, 0x04 }, xlsx.Take(4).ToArray());
        Assert.Equal("Region", sut.ReadCell(xlsx, "Sales", "A1"));
        Assert.Equal("1200", sut.ReadCell(xlsx, "Sales", "B2"));

        var updated = sut.SetCell(xlsx, "Sales", "B2", 1500);
        Assert.Equal("1500", sut.ReadCell(updated, "Sales", "B2"));
    }

    [Fact]
    public void Create_RejectsABlankSheetName()
    {
        var sut = new WorkbookEditorService();

        Assert.Throws<ArgumentException>(() => sut.Create(" ", new object?[][] { }));
    }
}
```

- [ ] **Step 2: Run to verify it fails**

Run: `dotnet test tests/DocToolkit.Extensions.DependencyInjection.Tests --filter FullyQualifiedName~WorkbookEditorServiceTests`
Expected: **build failure** — `WorkbookEditorService` does not exist.

- [ ] **Step 3: Implement**

Create `src/DocToolkit.Extensions.DependencyInjection/IWorkbookEditor.cs`:

```csharp
namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>Creates, reads and edits Excel (.xlsx) workbooks. Registered by <c>AddDocToolkit</c> (added in Task 7 of this plan).</summary>
public interface IWorkbookEditor
{
    /// <summary>Creates a workbook with one sheet populated from <paramref name="rows"/>.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="rows"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="sheetName"/> is blank, or a row is null.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The workbook could not be built.</exception>
    byte[] Create(string sheetName, IEnumerable<IEnumerable<object?>> rows);

    /// <summary>Reads a cell as a string. <paramref name="cellRef"/> is an A1-style reference.</summary>
    /// <exception cref="ArgumentNullException">Any argument is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="xlsx"/> is empty, or a name is blank.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The workbook could not be opened, the sheet does not exist, or the reference is not valid.</exception>
    string ReadCell(byte[] xlsx, string sheetName, string cellRef);

    /// <summary>Sets a cell and returns the updated workbook bytes.</summary>
    /// <exception cref="ArgumentNullException">Any argument other than <paramref name="value"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="xlsx"/> is empty, or a name is blank.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The workbook could not be opened, the sheet does not exist, or the reference is not valid.</exception>
    byte[] SetCell(byte[] xlsx, string sheetName, string cellRef, object? value);
}
```

Create `src/DocToolkit.Extensions.DependencyInjection/WorkbookEditorService.cs`:

```csharp
namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>Default <see cref="IWorkbookEditor"/>, delegating to <see cref="DocToolkit.WorkbookEditor"/>.</summary>
internal sealed class WorkbookEditorService : IWorkbookEditor
{
    public byte[] Create(string sheetName, IEnumerable<IEnumerable<object?>> rows)
        => DocToolkit.WorkbookEditor.Create(sheetName, rows);

    public string ReadCell(byte[] xlsx, string sheetName, string cellRef)
        => DocToolkit.WorkbookEditor.ReadCell(xlsx, sheetName, cellRef);

    public byte[] SetCell(byte[] xlsx, string sheetName, string cellRef, object? value)
        => DocToolkit.WorkbookEditor.SetCell(xlsx, sheetName, cellRef, value);
}
```

- [ ] **Step 4: Run to verify it passes**

Run: `dotnet test tests/DocToolkit.Extensions.DependencyInjection.Tests --filter FullyQualifiedName~WorkbookEditorServiceTests`
Expected: **PASS**, 4 results (2 tests × 2 TFMs).

- [ ] **Step 5: Commit**

```bash
git add src/DocToolkit.Extensions.DependencyInjection tests/DocToolkit.Extensions.DependencyInjection.Tests
git commit -m "feat(di-extensions): add IWorkbookEditor"
```

---

### Task 6: IPresentationEditor

**Files:**
- Create: `src/DocToolkit.Extensions.DependencyInjection/IPresentationEditor.cs`
- Create: `src/DocToolkit.Extensions.DependencyInjection/PresentationEditorService.cs`
- Create: `tests/DocToolkit.Extensions.DependencyInjection.Tests/PresentationEditorServiceTests.cs`
- Modify: `tests/DocToolkit.Extensions.DependencyInjection.Tests/DocToolkit.Extensions.DependencyInjection.Tests.csproj`

**Interfaces:**
- Consumes: `DocToolkit.PresentationEditor.SlideCount(byte[] pptx) -> int`, `.ExtractText(byte[] pptx) -> IReadOnlyList<string>`, `.ReplaceText(byte[] pptx, IReadOnlyDictionary<string, string> replacements) -> byte[]`.
- Produces: `DocToolkit.Extensions.DependencyInjection.IPresentationEditor` mirroring those three signatures.

**Fixture note:** a real one-slide `.pptx` (`"Hello {{who}}"`) already exists at `tests/DocToolkit.Tests/assets/sample.pptx`. Link it into this test project rather than committing a second copy of the same binary.

- [ ] **Step 1: Link the existing PPTX fixture**

Add to `tests/DocToolkit.Extensions.DependencyInjection.Tests/DocToolkit.Extensions.DependencyInjection.Tests.csproj`, inside the `<Project>` element:

```xml
  <ItemGroup>
    <Content Include="..\DocToolkit.Tests\assets\sample.pptx" Link="assets\sample.pptx" CopyToOutputDirectory="PreserveNewest" />
  </ItemGroup>
```

- [ ] **Step 2: Write the failing test**

Create `tests/DocToolkit.Extensions.DependencyInjection.Tests/PresentationEditorServiceTests.cs`:

```csharp
using System.Collections.Generic;
using System.IO;
using DocToolkit;
using DocToolkit.Extensions.DependencyInjection;
using Xunit;

namespace DocToolkit.Extensions.DependencyInjection.Tests;

public class PresentationEditorServiceTests
{
    private static byte[] SamplePptx() => File.ReadAllBytes(Path.Combine("assets", "sample.pptx"));

    [Fact]
    public void SlideCount_ExtractText_MatchTheStaticMethods()
    {
        var pptx = SamplePptx();
        var sut = new PresentationEditorService();

        Assert.Equal(PresentationEditor.SlideCount(pptx), sut.SlideCount(pptx));
        Assert.Equal(PresentationEditor.ExtractText(pptx), sut.ExtractText(pptx));
    }

    [Fact]
    public void ReplaceText_SubstitutesPlaceholders()
    {
        var pptx = SamplePptx();
        var sut = new PresentationEditorService();
        var replacements = new Dictionary<string, string> { ["{{who}}"] = "World" };

        var edited = sut.ReplaceText(pptx, replacements);

        var text = sut.ExtractText(edited);
        Assert.Contains(text, t => t.Contains("Hello World"));
        Assert.DoesNotContain(text, t => t.Contains("{{who}}"));
    }
}
```

- [ ] **Step 3: Run to verify it fails**

Run: `dotnet test tests/DocToolkit.Extensions.DependencyInjection.Tests --filter FullyQualifiedName~PresentationEditorServiceTests`
Expected: **build failure** — `PresentationEditorService` does not exist.

- [ ] **Step 4: Implement**

Create `src/DocToolkit.Extensions.DependencyInjection/IPresentationEditor.cs`:

```csharp
namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>Opens and edits PowerPoint (.pptx) presentations. Registered by <c>AddDocToolkit</c> (added in Task 7 of this plan).</summary>
public interface IPresentationEditor
{
    /// <summary>Number of slides in the deck, as counted from the deck's slide list.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="pptx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="pptx"/> is empty.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The package could not be opened or read.</exception>
    int SlideCount(byte[] pptx);

    /// <summary>All text found on every slide, one entry per text-bearing body, in deck order.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="pptx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="pptx"/> is empty.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The package could not be opened or read.</exception>
    IReadOnlyList<string> ExtractText(byte[] pptx);

    /// <summary>Replaces every key with its value across all slide text, returning updated bytes.</summary>
    /// <exception cref="ArgumentNullException">Either argument is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="pptx"/> is empty.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The package could not be opened or edited.</exception>
    byte[] ReplaceText(byte[] pptx, IReadOnlyDictionary<string, string> replacements);
}
```

Create `src/DocToolkit.Extensions.DependencyInjection/PresentationEditorService.cs`:

```csharp
namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>Default <see cref="IPresentationEditor"/>, delegating to <see cref="DocToolkit.PresentationEditor"/>.</summary>
internal sealed class PresentationEditorService : IPresentationEditor
{
    public int SlideCount(byte[] pptx) => DocToolkit.PresentationEditor.SlideCount(pptx);

    public IReadOnlyList<string> ExtractText(byte[] pptx) => DocToolkit.PresentationEditor.ExtractText(pptx);

    public byte[] ReplaceText(byte[] pptx, IReadOnlyDictionary<string, string> replacements)
        => DocToolkit.PresentationEditor.ReplaceText(pptx, replacements);
}
```

- [ ] **Step 5: Run to verify it passes**

Run: `dotnet test tests/DocToolkit.Extensions.DependencyInjection.Tests --filter FullyQualifiedName~PresentationEditorServiceTests`
Expected: **PASS**, 4 results (2 tests × 2 TFMs).

- [ ] **Step 6: Commit**

```bash
git add src/DocToolkit.Extensions.DependencyInjection tests/DocToolkit.Extensions.DependencyInjection.Tests
git commit -m "feat(di-extensions): add IPresentationEditor"
```

---

### Task 7: ServiceCollectionExtensions.AddDocToolkit()

**Files:**
- Create: `src/DocToolkit.Extensions.DependencyInjection/ServiceCollectionExtensions.cs`
- Create: `tests/DocToolkit.Extensions.DependencyInjection.Tests/LoopbackProbe.cs`
- Create: `tests/DocToolkit.Extensions.DependencyInjection.Tests/ServiceCollectionExtensionsTests.cs`
- Modify: `tests/DocToolkit.Extensions.DependencyInjection.Tests/DocToolkit.Extensions.DependencyInjection.Tests.csproj`

**Interfaces:**
- Consumes: all six interfaces and their services from Tasks 1–6, plus `DocToolkit.HtmlToDocxConverter` (indirectly, through `IHtmlToDocxConverter`) for the network-wiring proof.
- Produces: `DocToolkit.Extensions.DependencyInjection.ServiceCollectionExtensions.AddDocToolkit(this IServiceCollection services, Action<DocToolkitOptions>? configure = null) -> IServiceCollection`.

This is the task the design's own registration/lifetime/options tests belong to (§7 of the design doc): resolving all six, singleton lifetime, and proving `AllowRemoteImageDownload` genuinely changes runtime behaviour rather than just being an inert property — the same category of proof the core project's `AirGapGuardTests` makes for the static API, applied here to prove the **wiring**, not re-prove the underlying conversion behaviour (which is already covered exhaustively in `DocToolkit.Tests`).

- [ ] **Step 1: Add the full DI package to the test project**

The library only references `.Abstractions` (keeping the shipped package's footprint minimal); the *test* project needs the full package for `ServiceCollection`/`BuildServiceProvider()`:

```bash
dotnet add tests/DocToolkit.Extensions.DependencyInjection.Tests package Microsoft.Extensions.DependencyInjection
```

- [ ] **Step 2: Write the loopback probe test helper**

A minimal, self-contained proof that a connection was or wasn't attempted — the same raw-`TcpListener` technique `DocToolkit.Tests/AirGapGuardTests.cs` already uses successfully in this repo's CI, reimplemented here at the smaller scope this test actually needs (this project does not reference `DocToolkit.Tests`, so nothing is shared — the *conversion* behaviour behind the flag is already proven there; this only has to prove the one line of wiring from `DocToolkitOptions` to the static method's parameter).

Create `tests/DocToolkit.Extensions.DependencyInjection.Tests/LoopbackProbe.cs`:

```csharp
using System.Net;
using System.Net.Sockets;
using System.Text;

namespace DocToolkit.Extensions.DependencyInjection.Tests;

/// <summary>
/// A minimal loopback TCP listener, proving whether <see cref="DocToolkitOptions.AllowRemoteImageDownload"/>
/// registered through <see cref="ServiceCollectionExtensions.AddDocToolkit"/> actually reached the
/// converter. Answers every connection with a tiny valid image so a fetch completes cleanly
/// instead of hanging or erroring - the assertion only cares whether a connection was accepted.
/// </summary>
internal sealed class LoopbackProbe : IDisposable
{
    private readonly TcpListener _listener;
    private readonly CancellationTokenSource _stopping = new();
    private int _connections;

    public LoopbackProbe()
    {
        _listener = new TcpListener(IPAddress.Loopback, 0);
        _listener.Start();
        Port = ((IPEndPoint)_listener.LocalEndpoint).Port;
        _ = AcceptLoopAsync();
    }

    public int Port { get; }

    public string ImageUrl => $"http://127.0.0.1:{Port}/x.gif";

    public int Connections => Volatile.Read(ref _connections);

    public async Task<bool> WaitForConnectionAsync(TimeSpan timeout)
    {
        var deadline = DateTime.UtcNow + timeout;
        while (DateTime.UtcNow < deadline)
        {
            if (Connections > 0) return true;
            await Task.Delay(25);
        }

        return Connections > 0;
    }

    private async Task AcceptLoopAsync()
    {
        while (!_stopping.IsCancellationRequested)
        {
            TcpClient client;
            try
            {
                client = await _listener.AcceptTcpClientAsync(_stopping.Token);
            }
            catch
            {
                return; // Listener stopped, or the test finished.
            }

            Interlocked.Increment(ref _connections);
            _ = RespondAsync(client);
        }
    }

    private static async Task RespondAsync(TcpClient client)
    {
        using (client)
        {
            try
            {
                var stream = client.GetStream();
                using var reader = new StreamReader(stream, leaveOpen: true);
                while (!string.IsNullOrEmpty(await reader.ReadLineAsync())) { }

                var body = System.Convert.FromHexString(
                    "47494638396101000100800000000000ffffff21f90401000000002c00000000010001000002024401003b");
                var header = Encoding.ASCII.GetBytes(
                    $"HTTP/1.1 200 OK\r\nContent-Type: image/gif\r\nContent-Length: {body.Length}\r\nConnection: close\r\n\r\n");

                await stream.WriteAsync(header);
                await stream.WriteAsync(body);
            }
            catch
            {
                // Best effort - the assertion only needs Connections to have been incremented.
            }
        }
    }

    public void Dispose()
    {
        _stopping.Cancel();
        _listener.Stop();
    }
}
```

- [ ] **Step 3: Write the failing tests**

Create `tests/DocToolkit.Extensions.DependencyInjection.Tests/ServiceCollectionExtensionsTests.cs`:

```csharp
using DocToolkit.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Options;
using Xunit;

namespace DocToolkit.Extensions.DependencyInjection.Tests;

public class ServiceCollectionExtensionsTests
{
    [Fact]
    public void AddDocToolkit_ResolvesAllSixInterfaces()
    {
        var provider = new ServiceCollection().AddDocToolkit().BuildServiceProvider();

        Assert.NotNull(provider.GetRequiredService<IHtmlToDocxConverter>());
        Assert.NotNull(provider.GetRequiredService<IDocxToPdfConverter>());
        Assert.NotNull(provider.GetRequiredService<IHtmlToPdfConverter>());
        Assert.NotNull(provider.GetRequiredService<IDocxEditor>());
        Assert.NotNull(provider.GetRequiredService<IWorkbookEditor>());
        Assert.NotNull(provider.GetRequiredService<IPresentationEditor>());
    }

    [Fact]
    public void AddDocToolkit_RegistersEachInterfaceAsASingleton()
    {
        var provider = new ServiceCollection().AddDocToolkit().BuildServiceProvider();

        Assert.Same(
            provider.GetRequiredService<IHtmlToDocxConverter>(),
            provider.GetRequiredService<IHtmlToDocxConverter>());
        Assert.Same(
            provider.GetRequiredService<IWorkbookEditor>(),
            provider.GetRequiredService<IWorkbookEditor>());
    }

    [Fact]
    public void AddDocToolkit_WithNoConfigure_DefaultsToNoRemoteImageDownload()
    {
        var provider = new ServiceCollection().AddDocToolkit().BuildServiceProvider();

        Assert.False(provider.GetRequiredService<IOptions<DocToolkitOptions>>().Value.AllowRemoteImageDownload);
    }

    [Fact]
    public void AddDocToolkit_WithConfigure_MakesTheValueObservableViaIOptions()
    {
        var provider = new ServiceCollection()
            .AddDocToolkit(o => o.AllowRemoteImageDownload = true)
            .BuildServiceProvider();

        Assert.True(provider.GetRequiredService<IOptions<DocToolkitOptions>>().Value.AllowRemoteImageDownload);
    }

    [Fact]
    public async Task AddDocToolkit_WithAllowRemoteImageDownloadFalse_NeverConnectsOutbound()
    {
        using var probe = new LoopbackProbe();
        var provider = new ServiceCollection().AddDocToolkit().BuildServiceProvider();
        var sut = provider.GetRequiredService<IHtmlToDocxConverter>();

        await sut.ConvertAsync($"<img src=\"{probe.ImageUrl}\">");
        await Task.Delay(300);

        Assert.Equal(0, probe.Connections);
    }

    [Fact]
    public async Task AddDocToolkit_WithAllowRemoteImageDownloadTrue_DoesConnectOutbound()
    {
        using var probe = new LoopbackProbe();
        var provider = new ServiceCollection()
            .AddDocToolkit(o => o.AllowRemoteImageDownload = true)
            .BuildServiceProvider();
        var sut = provider.GetRequiredService<IHtmlToDocxConverter>();

        await sut.ConvertAsync($"<img src=\"{probe.ImageUrl}\">");

        Assert.True(await probe.WaitForConnectionAsync(TimeSpan.FromSeconds(5)));
    }
}
```

- [ ] **Step 4: Run to verify it fails**

Run: `dotnet test tests/DocToolkit.Extensions.DependencyInjection.Tests --filter FullyQualifiedName~ServiceCollectionExtensionsTests`
Expected: **build failure** — `AddDocToolkit` does not exist.

- [ ] **Step 5: Implement**

Create `src/DocToolkit.Extensions.DependencyInjection/ServiceCollectionExtensions.cs`:

```csharp
using Microsoft.Extensions.DependencyInjection;

namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>Registers DocToolkit's DI-friendly services.</summary>
public static class ServiceCollectionExtensions
{
    /// <summary>
    /// Registers <see cref="IHtmlToDocxConverter"/>, <see cref="IDocxToPdfConverter"/>,
    /// <see cref="IHtmlToPdfConverter"/>, <see cref="IDocxEditor"/>, <see cref="IWorkbookEditor"/>
    /// and <see cref="IPresentationEditor"/> as singletons - each wraps a stateless static class,
    /// so one shared instance is safe under concurrent use.
    /// </summary>
    /// <param name="services">The service collection to add to.</param>
    /// <param name="configure">
    /// Configures <see cref="DocToolkitOptions"/>. Leave null to keep every default -
    /// <see cref="DocToolkitOptions.AllowRemoteImageDownload"/> stays <c>false</c>.
    /// </param>
    public static IServiceCollection AddDocToolkit(
        this IServiceCollection services, Action<DocToolkitOptions>? configure = null)
    {
        services.AddOptions<DocToolkitOptions>();
        if (configure is not null) services.Configure(configure);

        services.AddSingleton<IHtmlToDocxConverter, HtmlToDocxConverterService>();
        services.AddSingleton<IDocxToPdfConverter, DocxToPdfConverterService>();
        services.AddSingleton<IHtmlToPdfConverter, HtmlToPdfConverterService>();
        services.AddSingleton<IDocxEditor, DocxEditorService>();
        services.AddSingleton<IWorkbookEditor, WorkbookEditorService>();
        services.AddSingleton<IPresentationEditor, PresentationEditorService>();

        return services;
    }
}
```

- [ ] **Step 6: Run to verify it passes**

Run: `dotnet test tests/DocToolkit.Extensions.DependencyInjection.Tests --filter FullyQualifiedName~ServiceCollectionExtensionsTests`
Expected: **PASS**, 12 results (6 tests × 2 TFMs).

- [ ] **Step 7: Run the whole new test project**

Run: `dotnet build DocToolkit.sln -c Release -warnaserror && dotnet test tests/DocToolkit.Extensions.DependencyInjection.Tests`
Expected: **0 warnings**, all pass — 42 results total (21 tests × 2 TFMs: 2 options + 2 HtmlToDocx + 2 DocxToPdf + 2 HtmlToPdf + 3 DocxEditor + 2 WorkbookEditor + 2 PresentationEditor + 6 ServiceCollectionExtensions = 21).

- [ ] **Step 8: Run the ENTIRE solution's test suite**

Run: `dotnet test DocToolkit.sln`
Expected: **0 failures** — the existing `DocToolkit.Tests` project (182 tests × 2 TFMs = 364 results) is completely untouched, plus the 21×2 = 42 new results.

- [ ] **Step 9: Commit**

```bash
git add src/DocToolkit.Extensions.DependencyInjection tests/DocToolkit.Extensions.DependencyInjection.Tests
git commit -m "feat(di-extensions): add AddDocToolkit() service registration"
```

---

### Task 8: Extend CI to cover both packages

**Files:**
- Modify: `.github/workflows/ci.yml`

**Interfaces:**
- Consumes: nothing new — this wires existing CI jobs to the new project's path.
- Produces: nothing consumed by later tasks.

The `build-test` job already runs `dotnet build DocToolkit.sln` / `dotnet test DocToolkit.sln`, and the native-binary check already scans the whole checkout (`find .`) — both automatically cover the new project once it's part of the solution, with **no change needed**. Two things do need to change: the banned-package check only inspects `env.PROJECT` (the core csproj), and the package job only packs/verifies the core project.

- [ ] **Step 1: Confirm what already "just works" without changes**

Run locally: `dotnet build DocToolkit.sln -c Release && find . -path '*/bin/*' \( -name '*.so' -o -name '*.dylib' \) | wc -l`
Expected: build succeeds for both projects; native-file count is `0`.

- [ ] **Step 2: Add an env var for the extensions project path**

In `.github/workflows/ci.yml`, in the top-level `env:` block, add one line after `PROJECT`:

```yaml
env:
  DOTNET_NOLOGO: true
  DOTNET_SKIP_FIRST_TIME_EXPERIENCE: true
  DOTNET_CLI_TELEMETRY_OPTOUT: true
  SOLUTION: DocToolkit.sln
  PROJECT: src/DocToolkit/DocToolkit.csproj
  EXTENSIONS_PROJECT: src/DocToolkit.Extensions.DependencyInjection/DocToolkit.Extensions.DependencyInjection.csproj
```

- [ ] **Step 3: Extend the banned-packages check to cover both projects**

Replace the `premise-guard` job's "Assert no banned packages in the resolved graph" step:

```yaml
      - name: Assert no banned packages in the resolved graph
        run: |
          banned='EPPlus|NPOI|Spire\.|Syncfusion|QuestPDF|IronPDF|ShapeCrawler|SkiaSharp|Magick\.NET|System\.Drawing\.Common'
          for project in "$PROJECT" "$EXTENSIONS_PROJECT"; do
            echo "::group::$project"
            dotnet list "$project" package --include-transitive > graph.txt
            cat graph.txt
            if grep -Eiq "$banned" graph.txt; then
              grep -Ei "$banned" graph.txt
              echo "::error::A banned package is in the dependency graph of $project."
              exit 1
            fi
            echo "no banned packages"
            echo "::endgroup::"
          done
```

The "Assert zero native binaries" step and the "Assert SixLabors.Fonts stayed on 1.x" step need no change — the former already scans the whole checkout, and the latter only makes sense for the core project (the extensions project never references OfficeIMO/SixLabors.Fonts at all).

- [ ] **Step 4: Extend the package job to pack and verify both packages**

Replace the entire `package` job:

```yaml
  # Proves the artifacts consumers actually install are correct - not just that
  # the solution compiles. Catches missing TFMs and metadata regressions.
  package:
    name: pack & verify .nupkg (${{ matrix.name }})
    runs-on: ubuntu-latest
    needs: build-test
    strategy:
      fail-fast: false
      matrix:
        include:
          - name: core
            project: src/DocToolkit/DocToolkit.csproj
            deps: "DocumentFormat.OpenXml,HtmlToOpenXml.dll,OfficeIMO.Word.Pdf,ClosedXML"
          - name: extensions
            project: src/DocToolkit.Extensions.DependencyInjection/DocToolkit.Extensions.DependencyInjection.csproj
            deps: "Ank.DocToolkit,Microsoft.Extensions.DependencyInjection.Abstractions,Microsoft.Extensions.Options"
    steps:
      - uses: actions/checkout@v4

      - uses: actions/setup-dotnet@v4
        with:
          dotnet-version: |
            8.0.x
            10.0.x

      - name: Pack
        run: dotnet pack ${{ matrix.project }} -c Release

      - name: Verify package contents
        env:
          EXPECTED_DEPS: ${{ matrix.deps }}
        run: |
          python3 - <<'PY'
          import glob, os, sys, zipfile
          deps = os.environ['EXPECTED_DEPS'].split(',')
          hits = glob.glob('artifacts/*.nupkg')
          if not hits:
              sys.exit('no .nupkg produced')
          pkg = hits[0]
          print('package:', pkg)
          z = zipfile.ZipFile(pkg)
          names = z.namelist()
          nuspec = z.read([n for n in names if n.endswith('.nuspec')][0]).decode('utf-8')

          failures = []
          for required in ['README.md', 'THIRD-PARTY-NOTICES.txt']:
              ok = required in names
              print(('OK   ' if ok else 'MISS ') + required)
              if not ok:
                  failures.append(required)

          has_net8 = any(n.startswith('lib/net8.0/') for n in names)
          has_net10 = any(n.startswith('lib/net10.0/') for n in names)
          print(('OK   ' if has_net8 else 'MISS ') + 'lib/net8.0/')
          print(('OK   ' if has_net10 else 'MISS ') + 'lib/net10.0/')
          if not has_net8: failures.append('lib/net8.0/')
          if not has_net10: failures.append('lib/net10.0/')

          for dep in deps:
              ok = dep in nuspec
              print(('OK   dep ' if ok else 'MISS dep ') + dep)
              if not ok:
                  failures.append('dep ' + dep)

          if '<license type="expression">MIT' not in nuspec:
              failures.append('MIT licence expression')
          if any(n.startswith('runtimes/') for n in names):
              failures.append('runtimes/ present - native payload in package')

          if failures:
              sys.exit('package verification failed: ' + ', '.join(failures))
          print('\npackage verified')
          PY

      - name: Upload package
        uses: actions/upload-artifact@v4
        with:
          name: nupkg-${{ matrix.name }}
          path: artifacts/*.nupkg
          if-no-files-found: error
```

- [ ] **Step 5: Push to a branch and confirm CI passes on both matrix legs**

```bash
git checkout -b ci/cover-di-extensions
git add .github/workflows/ci.yml
git commit -m "ci: cover the DI extensions package in build, guard and package jobs"
git push -u origin ci/cover-di-extensions
```

Watch the run (e.g. `gh run watch`). Expected: `build-test` (2 OS legs), `premise-guard`, and **`package` running twice** — `pack & verify .nupkg (core)` and `pack & verify .nupkg (extensions)` — all green.

- [ ] **Step 6: Merge to main**

Open and merge a PR (or fast-forward merge locally, matching however this repo has merged prior branches), once CI is green.

---

### Task 9: Release workflow, package metadata, and documentation

**Files:**
- Create: `.github/workflows/release-extensions.yml`
- Modify: `src/DocToolkit.Extensions.DependencyInjection/README.md`
- Modify: `src/DocToolkit.Extensions.DependencyInjection/THIRD-PARTY-NOTICES.txt`
- Modify: `CLAUDE.md`
- Modify: `README.md` (repo root)

**Interfaces:**
- Consumes: nothing new.
- Produces: nothing consumed by later tasks — this is the final task.

The two packages are versioned independently (per the design), so they need independent release triggers. Rather than parameterising the existing `release.yml` (which would need conditional logic threaded through nearly every step — version resolution, guards, verification, the nuget.org Trusted Publishing policy lookup), this creates a **second, standalone workflow** on its own tag prefix, mirroring `release.yml`'s structure exactly. Both approaches were considered in the design (§9); a separate file is simpler to read and cannot accidentally publish the wrong package from the wrong step.

- [ ] **Step 1: Create the release workflow**

Create `.github/workflows/release-extensions.yml`:

```yaml
name: Release Extensions

# Sibling of release.yml, for Ank.DocToolkit.Extensions.DependencyInjection specifically. A
# separate file rather than parameterising release.yml: the two packages have different expected
# dependencies to verify and different premise guards (this one has no SixLabors.Fonts check -
# it never references OfficeIMO), and nuget.org's Trusted Publishing policy is keyed to an exact
# workflow FILENAME, so each package needs its own policy entry regardless.
#
# Trigger by pushing a tag:   git tag ext-v1.0.0 && git push origin ext-v1.0.0
# Tag prefix is "ext-v", not "v" - the core package already owns "v*" in release.yml, and the two
# packages version independently.
#
# AUTHENTICATION: Trusted Publishing (OIDC), same mechanism as release.yml. Requires ITS OWN
# nuget.org policy (your name > Trusted Publishing):
#   Repository Owner : Ank-KhoaHo
#   Repository       : DocToolkit
#   Workflow File    : release-extensions.yml     <- filename only, no path
# and the same NUGET_USER repository variable/secret release.yml already uses (shared, since it
# is just your nuget.org profile name).

on:
  push:
    tags: ["ext-v*"]
  workflow_dispatch:
    inputs:
      version:
        description: "Version to publish (e.g. 0.2.0). Dry-run unless 'publish' is ticked."
        required: true
      publish:
        description: "Actually push to nuget.org"
        type: boolean
        default: false

permissions:
  contents: write
  id-token: write

env:
  DOTNET_NOLOGO: true
  DOTNET_CLI_TELEMETRY_OPTOUT: true
  PROJECT: src/DocToolkit.Extensions.DependencyInjection/DocToolkit.Extensions.DependencyInjection.csproj
  SOLUTION: DocToolkit.sln

jobs:
  release:
    name: verify, pack and publish
    runs-on: ubuntu-latest

    steps:
      - uses: actions/checkout@v4
        with:
          fetch-depth: 0

      - uses: actions/setup-dotnet@v4
        with:
          dotnet-version: |
            8.0.x
            10.0.x

      - name: Resolve version
        id: v
        run: |
          if [ "${{ github.event_name }}" = "workflow_dispatch" ]; then
            VERSION="${{ inputs.version }}"
            PUBLISH="${{ inputs.publish }}"
          else
            VERSION="${GITHUB_REF_NAME#ext-v}"   # ext-v1.2.3 -> 1.2.3
            PUBLISH="true"
          fi
          if ! printf '%s' "$VERSION" | grep -Eq '^[0-9]+\.[0-9]+\.[0-9]+(-[0-9A-Za-z.-]+)?$'; then
            echo "::error::'$VERSION' is not a valid SemVer version. Tag as ext-v1.2.3 or ext-v1.2.3-beta.1."
            exit 1
          fi
          echo "version=$VERSION" >> "$GITHUB_OUTPUT"
          echo "publish=$PUBLISH" >> "$GITHUB_OUTPUT"
          case "$VERSION" in *-*) echo "prerelease=true"  >> "$GITHUB_OUTPUT";;
                             *)   echo "prerelease=false" >> "$GITHUB_OUTPUT";; esac
          echo "Releasing version $VERSION (publish=$PUBLISH)"

      - name: Fail early if the nuget.org username is missing
        if: steps.v.outputs.publish == 'true'
        run: |
          USER="${{ secrets.NUGET_USER || vars.NUGET_USER }}"
          if [ -z "$USER" ]; then
            echo "::error::NUGET_USER is not set. Add your nuget.org PROFILE NAME (not your email) as a repository variable or secret under Settings > Secrets and variables > Actions."
            exit 1
          fi
          echo "nuget.org user configured"

      - name: Build
        run: dotnet build ${{ env.SOLUTION }} -c Release -warnaserror

      - name: Test
        run: dotnet test ${{ env.SOLUTION }} -c Release --no-build

      - name: Guard - no native binaries
        run: |
          mapfile -t native < <(find . -path '*/bin/*' \
            \( -name '*.so' -o -name '*.so.*' -o -name '*.dylib' \) -print)
          echo "native binaries found: ${#native[@]}"
          if [ "${#native[@]}" -ne 0 ]; then
            printf '  %s\n' "${native[@]}"
            echo "::error::Refusing to publish - native binaries in build output."
            exit 1
          fi

      - name: Guard - no banned packages
        run: |
          dotnet list ${{ env.PROJECT }} package --include-transitive > graph.txt
          banned='EPPlus|NPOI|Spire\.|Syncfusion|QuestPDF|IronPDF|ShapeCrawler|SkiaSharp|Magick\.NET|System\.Drawing\.Common'
          if grep -Eiq "$banned" graph.txt; then
            grep -Ei "$banned" graph.txt
            echo "::error::Refusing to publish - a package that is not free for commercial use, or carries native payload, is in the graph."
            exit 1
          fi
          echo "no banned packages"

      - name: Pack
        run: |
          dotnet pack ${{ env.PROJECT }} -c Release --no-build \
            -p:Version=${{ steps.v.outputs.version }} \
            -p:PackageVersion=${{ steps.v.outputs.version }}
          ls -la artifacts/

      - name: Verify package contents
        run: |
          python3 - "${{ steps.v.outputs.version }}" <<'PY'
          import glob, sys, zipfile
          version = sys.argv[1]
          hits = sorted(glob.glob('artifacts/*.nupkg'))
          matching = [h for h in hits if version in h]
          if not matching:
              sys.exit(f'no .nupkg found carrying version {version} in {hits}')
          pkg = matching[0]
          z, failures = zipfile.ZipFile(pkg), []
          names = z.namelist()
          nuspec = [n for n in names if n.endswith('.nuspec')][0]
          spec = z.read(nuspec).decode('utf-8')

          for required in ['lib/net8.0/', 'lib/net10.0/', 'README.md', 'THIRD-PARTY-NOTICES.txt']:
              ok = any(n.startswith(required) or n == required for n in names)
              print(('OK   ' if ok else 'MISS ') + required)
              if not ok: failures.append(required)

          for dep in ['Ank.DocToolkit', 'Microsoft.Extensions.DependencyInjection.Abstractions', 'Microsoft.Extensions.Options']:
              ok = dep in spec
              print(('OK   dep ' if ok else 'MISS dep ') + dep)
              if not ok: failures.append('dep ' + dep)

          if '<license type="expression">MIT' not in spec:
              failures.append('MIT licence expression')
          if any(n.startswith('runtimes/') for n in names):
              failures.append('runtimes/ - native payload in package')
          if f'<version>{version}</version>' not in spec:
              failures.append(f'nuspec version is not {version}')

          if failures:
              sys.exit('package verification failed: ' + ', '.join(failures))
          print(f'\n{pkg} verified')
          PY

      - name: NuGet login (OIDC exchange for a short-lived key)
        id: login
        uses: NuGet/login@v1
        with:
          user: ${{ secrets.NUGET_USER || vars.NUGET_USER }}

      - name: Confirm the OIDC exchange produced a key
        run: |
          if [ -z "${{ steps.login.outputs.NUGET_API_KEY }}" ]; then
            echo "::error::OIDC exchange returned no key. Check: (a) permissions.id-token is 'write', (b) the nuget.org Trusted Publishing policy names owner '${{ github.repository_owner }}', repo '${{ github.event.repository.name }}' and workflow file 'release-extensions.yml', (c) NUGET_USER is your nuget.org profile name, not an email."
            exit 1
          fi
          echo "Trusted Publishing OK - short-lived key obtained (valid 1 hour)"

      - name: Push to nuget.org
        if: steps.v.outputs.publish == 'true'
        run: |
          dotnet nuget push "artifacts/Ank.DocToolkit.Extensions.DependencyInjection.${{ steps.v.outputs.version }}.nupkg" \
            --source https://api.nuget.org/v3/index.json \
            --api-key "${{ steps.login.outputs.NUGET_API_KEY }}" \
            --skip-duplicate
          echo "published ${{ steps.v.outputs.version }}"

      - name: Create GitHub Release
        if: steps.v.outputs.publish == 'true' && github.event_name == 'push'
        env:
          GH_TOKEN: ${{ github.token }}
        run: |
          gh release create "${GITHUB_REF_NAME}" \
            --title "${GITHUB_REF_NAME}" \
            --generate-notes \
            ${{ steps.v.outputs.prerelease == 'true' && '--prerelease' || '' }} \
            "artifacts/Ank.DocToolkit.Extensions.DependencyInjection.${{ steps.v.outputs.version }}.nupkg" \
            "artifacts/Ank.DocToolkit.Extensions.DependencyInjection.${{ steps.v.outputs.version }}.snupkg"

      - name: Upload package as a build artifact
        uses: actions/upload-artifact@v4
        with:
          name: nupkg-extensions-${{ steps.v.outputs.version }}
          path: "artifacts/Ank.DocToolkit.Extensions.DependencyInjection.${{ steps.v.outputs.version }}.*"
          if-no-files-found: error
```

**Note:** the `Push to nuget.org` / `Create GitHub Release` / `Upload package` steps reference the extensions package's `.nupkg` by exact filename rather than a bare `artifacts/*.nupkg` glob — unlike `release.yml`, which packs only one project. This CI job only packs the extensions project too (nothing else writes into `artifacts/` during this run), so a glob would in fact be equally safe, but naming the file explicitly makes it unambiguous to a reader that this workflow only ever touches the extensions package.

- [ ] **Step 2: Finalise the package README**

Replace `src/DocToolkit.Extensions.DependencyInjection/README.md`:

````markdown
# Ank.DocToolkit.Extensions.DependencyInjection

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

All six interfaces — `IHtmlToDocxConverter`, `IDocxToPdfConverter`, `IHtmlToPdfConverter`,
`IDocxEditor`, `IWorkbookEditor`, `IPresentationEditor` — mirror
[`Ank.DocToolkit`](https://www.nuget.org/packages/Ank.DocToolkit)'s static API one-for-one, are
registered as singletons (each wraps stateless logic), and are safe to inject and call
concurrently. See the core package's README for what each one does and the offline/licensing
guarantees behind them.

## Why a separate package

A console app, Lambda or simple script that only wants the static `byte[]`-based API installs
just `Ank.DocToolkit`, with zero DI dependencies. ASP.NET Core and worker-service consumers add
this package too.

## Licence

MIT — see the parent repository's [LICENSE](https://github.com/Ank-KhoaHo/DocToolkit/blob/main/LICENSE).
````

- [ ] **Step 3: Finalise the third-party notices**

Replace `src/DocToolkit.Extensions.DependencyInjection/THIRD-PARTY-NOTICES.txt`:

```text
Ank.DocToolkit.Extensions.DependencyInjection third-party notices
===================================================================

This package depends on the following. All are permissively licensed and free for commercial use.

1. Ank.DocToolkit - MIT License
   Copyright (c) Khoa Ho
   https://github.com/Ank-KhoaHo/DocToolkit
   (See that package's own THIRD-PARTY-NOTICES.txt for its further dependencies.)

2. Microsoft.Extensions.DependencyInjection.Abstractions - MIT License
   Copyright (c) .NET Foundation and Contributors
   https://github.com/dotnet/runtime

3. Microsoft.Extensions.Options - MIT License
   Copyright (c) .NET Foundation and Contributors
   https://github.com/dotnet/runtime

The full MIT licence text is reproduced in the LICENSE file in the parent repository.
```

- [ ] **Step 4: Rebuild and re-verify the package**

```bash
dotnet pack src/DocToolkit.Extensions.DependencyInjection -c Release
```

Expected: succeeds; `artifacts/Ank.DocToolkit.Extensions.DependencyInjection.0.1.0.nupkg` now embeds the finalised README and notices.

- [ ] **Step 5: Update CLAUDE.md**

Add a new section to `CLAUDE.md`, after the existing `## Conventions` section:

```markdown
## The DI extensions package

`src/DocToolkit.Extensions.DependencyInjection/` ships as its own NuGet package,
`Ank.DocToolkit.Extensions.DependencyInjection`, versioned and released independently of the core
package (tag prefix `ext-v*`, via `.github/workflows/release-extensions.yml` — see that file's
header comment for the matching nuget.org Trusted Publishing policy it needs).

It references `Ank.DocToolkit` as a real `PackageReference`, never a `ProjectReference` — the
whole point is to prove the extensions package works the way an external consumer's restore
would, against the *published* core package, not against whatever is currently on `main`. Before
changing an interface here, confirm the byte[] method it wraps actually exists in the core
version this project's `Ank.DocToolkit` reference floor requires.

Six interfaces mirror the six static classes 1:1 (`byte[]` in, `byte[]`/`string`/`int` out — no
`Stream` overloads here; that was a deliberate scope decision, not an oversight, since the DI
layer was designed before the static API's `Stream` overloads existed). Service implementations
are `internal sealed` — never `public` — and are pure delegation, one line per method, to the
matching static method. If a service method does anything more than call through, that logic
belongs in the core static method instead.

`DocToolkitOptions.AllowRemoteImageDownload` replaces the static API's per-call
`allowRemoteImageDownload` bool: configured once at `AddDocToolkit(configure)`, not re-decided per
call. `ServiceCollectionExtensionsTests` proves the wiring with a small self-contained loopback
listener (not a copy of the core project's `AirGapGuardTests` — that already proves the
*conversion* behaviour exhaustively; this only has to prove the option value reaches the static
method's parameter).
```

- [ ] **Step 6: Update the repository root README**

Add a short section to `README.md`, after the main usage section:

```markdown
## Dependency injection

For ASP.NET Core / worker-service consumers:

```bash
dotnet add package Ank.DocToolkit.Extensions.DependencyInjection
```

See that package's own README for `AddDocToolkit()` usage.
```

- [ ] **Step 7: Full-solution final check**

```bash
dotnet build DocToolkit.sln -c Release -warnaserror
dotnet test DocToolkit.sln
dotnet pack src/DocToolkit/DocToolkit.csproj -c Release
dotnet pack src/DocToolkit.Extensions.DependencyInjection -c Release
```

Expected: 0 warnings; every test passes (core 364 results + extensions 42 results = 406); both
packages build.

- [ ] **Step 8: Commit and push**

```bash
git add .github/workflows/release-extensions.yml src/DocToolkit.Extensions.DependencyInjection CLAUDE.md README.md
git commit -m "docs(di-extensions): finalise package README, notices, and add the release workflow"
git push origin main
```

- [ ] **Step 9: Manual, user-side setup (not automatable by an implementer)**

Before the first `ext-v*` tag can publish:

1. On nuget.org → *your name → Trusted Publishing* → add a **second** policy: Repository Owner
   `Ank-KhoaHo`, Repository `DocToolkit`, Workflow File `release-extensions.yml` (filename only).
2. `NUGET_USER` is already set from the core package's release setup — nothing new needed there.
3. Recommended first release: a **dry run** first (`workflow_dispatch`, version `0.1.0`, `publish`
   left unticked) to prove the new policy resolves, before tagging `ext-v0.1.0` for real.

---

## Spec coverage

| Design requirement (§) | Task |
|---|---|
| §3 package structure, PackageReference not ProjectReference | Task 1 |
| §4 six interfaces mirroring the static classes | Tasks 1–6 (plus the `IDocxEditor` overload completion, noted in Task 4) |
| §5 `DocToolkitOptions.AllowRemoteImageDownload` | Task 1 (class), Tasks 1 & 3 (consumed by the two converters that need it) |
| §6 `AddDocToolkit()` registration, singleton lifetime | Task 7 |
| §7 parity tests, registration tests, options-actually-takes-effect test | Tasks 1–6 (parity), Task 7 (registration + options wiring) |
| §8 core package untouched; CI extended, not exempted | Task 8 |
| §9 CI mechanics decision | Task 8 (extend existing jobs, matrix for `package`) |
| §9 tagging scheme decision | Task 9 (`ext-v*`, separate workflow file) |

## Known risks

1. **Task 7's loopback-listener test opens a real (loopback-only) socket.** This is the same
   technique already proven reliable in this repo's CI (`DocToolkit.Tests/AirGapGuardTests.cs`),
   reimplemented at a smaller scope rather than shared, so it carries the same low, already-accepted
   risk profile (no known flakiness in this repo's CI history).
2. **The `IDocxEditor` interface gained a method the original written design didn't list**
   (`ExtractText(byte[], bool)`) — a deliberate completion toward the design's own stated "mirror
   1:1" goal, flagged explicitly in Task 4 rather than silently added.
3. **`release-extensions.yml` needs a second, separately-configured Trusted Publishing policy**
   on nuget.org before its first real publish — this is manual, user-side setup (Task 9, Step 9),
   not something an implementer can do from the repository.
