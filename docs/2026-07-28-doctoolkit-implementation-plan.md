# DocToolkit Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Ship `DocToolkit` as a **redistributable NuGet package** that converts HTML to DOCX and PDF and opens/edits DOCX, XLSX and PPTX files, so any other project can `dotnet add package DocToolkit` and reuse it — using only permissive-licensed, pure-managed dependencies that run on Linux.

**Architecture:** A thin, stateless facade over the stack verified in `implementation/dotnet-doc-libs/spike/`, packaged for distribution. The core API is byte-array in / byte-array out so it works server-side with no temp files; path overloads are convenience wrappers. HTML→PDF is *composed* from HTML→DOCX plus DOCX→PDF rather than reimplemented — DOCX is the pivot format, because no permissive NuGet-only library renders HTML to PDF directly on Linux. The five underlying packages flow to consumers as transitive dependencies; consumers never reference them directly.

**Tech Stack:** .NET 8 + .NET 10 · `DocumentFormat.OpenXml` (MIT) · `HtmlToOpenXml.dll` (MIT) · `OfficeIMO.Word.Pdf` (MIT) · `ClosedXML` (MIT) · `ShapeCrawler` (MIT) · xUnit

**Source spec:** [`learning-docs/dotnet-doc-libs/report.html`](../../learning-docs/dotnet-doc-libs/report.html) — the research and licence verification behind every package choice.

## Global Constraints

- **Target frameworks:** `net8.0;net10.0` for both the library and the test project. Tests therefore run twice — once per framework — and both must pass.
- **Package identity:** `PackageId` is `DocToolkit`, starting at version `0.1.0`, licensed `MIT`.
- **Distribution:** a **local folder feed** at `implementation/dotnet-doc-libs/local-feed/`. No credentials, no publishing. Do not run `dotnet nuget push` — adding a real feed later requires no code change.
- **The package must carry third-party licence notices.** Every dependency is permissive but each requires attribution; `THIRD-PARTY-NOTICES.txt` ships inside the `.nupkg`.
- **Licences:** only MIT / Apache-2.0 / BSD packages. **Never add** `EPPlus` (≥5 is Polyform Noncommercial), `NPOI` (≥2.8.0 requires a paid maintenance fee), `Spire.*`, `Syncfusion.*`, `QuestPDF`, or `IronPDF`.
- **No `System.Drawing.Common`,** directly or transitively. It resolves and builds fine, then throws `PlatformNotSupportedException` at runtime on Linux. Task 8 enforces this automatically.
- **No native binaries** — no Chromium, no LibreOffice, no SkiaSharp. Everything must work after `dotnet restore` alone.
- **Legacy `.xls` is out of scope** (descoped 2026-07-28). Do not add a package for it.
- **Repo is not yet a git repository.** Run `git init` at `E:\PJ\LnDPrj` once before Task 1, otherwise every commit step fails.
- **Working directory** for all commands is `implementation/dotnet-doc-libs/` unless stated otherwise.

## File Structure

```
implementation/dotnet-doc-libs/
├── DocToolkit.sln
├── LICENSE                         # MIT, ships in the .nupkg
├── src/DocToolkit/
│   ├── DocToolkit.csproj           # multi-targeted AND packable
│   ├── README.md                   # package landing page on the feed
│   ├── THIRD-PARTY-NOTICES.txt     # attribution for the 5 bundled deps
│   ├── HtmlToDocxConverter.cs      # HTML  -> DOCX bytes
│   ├── DocxToPdfConverter.cs       # DOCX  -> PDF bytes
│   ├── HtmlToPdfConverter.cs       # composes the two above
│   ├── DocxEditor.cs               # open/edit DOCX (placeholder replacement)
│   ├── WorkbookEditor.cs           # open/edit XLSX
│   ├── PresentationEditor.cs       # open/edit PPTX
│   └── DocumentConversionException.cs
├── tests/DocToolkit.Tests/
│   ├── DocToolkit.Tests.csproj
│   ├── PdfProbe.cs                 # test helper: decode PDF text/pages
│   ├── HtmlToDocxConverterTests.cs
│   ├── DocxToPdfConverterTests.cs
│   ├── HtmlToPdfConverterTests.cs
│   ├── DocxEditorTests.cs
│   ├── WorkbookEditorTests.cs
│   ├── PresentationEditorTests.cs
│   └── DependencyGuardTests.cs
├── artifacts/                      # .nupkg output (gitignored)
├── local-feed/                     # the folder feed consumers restore from
├── Dockerfile.linux-test
└── spike/                          # existing proof-of-concept, leave untouched
```

**Why `PdfProbe` is its own file:** OfficeIMO writes **uncompressed** content streams and emits text as **hex-string** operators (`<41636D65> Tj` == `"Acme"`). Searching the raw PDF bytes for `"Acme"` finds nothing, and inflating the streams finds nothing either — both fail *silently* and look like a broken converter. This gotcha cost real debugging time during the spike. Encapsulate it once.

---

### Task 1: Solution scaffold and HTML → DOCX

**Files:**
- Create: `DocToolkit.sln`, `src/DocToolkit/DocToolkit.csproj`, `src/DocToolkit/HtmlToDocxConverter.cs`, `src/DocToolkit/DocumentConversionException.cs`
- Create: `src/DocToolkit/README.md`, `src/DocToolkit/THIRD-PARTY-NOTICES.txt` (stubs; finalised in Task 9)
- Create: `tests/DocToolkit.Tests/DocToolkit.Tests.csproj`, `tests/DocToolkit.Tests/HtmlToDocxConverterTests.cs`

**Interfaces:**
- Consumes: nothing.
- Produces: `DocToolkit.HtmlToDocxConverter.ConvertAsync(string html, CancellationToken ct = default) -> Task<byte[]>` and `DocToolkit.DocumentConversionException`.

- [ ] **Step 1: Scaffold the solution and projects**

```bash
cd implementation/dotnet-doc-libs
dotnet new sln -n DocToolkit
dotnet new classlib -f net8.0 -o src/DocToolkit -n DocToolkit
dotnet new xunit  -f net8.0 -o tests/DocToolkit.Tests -n DocToolkit.Tests
rm src/DocToolkit/Class1.cs tests/DocToolkit.Tests/UnitTest1.cs
dotnet sln add src/DocToolkit/DocToolkit.csproj tests/DocToolkit.Tests/DocToolkit.Tests.csproj
dotnet add tests/DocToolkit.Tests reference src/DocToolkit
dotnet add src/DocToolkit package DocumentFormat.OpenXml
dotnet add src/DocToolkit package HtmlToOpenXml.dll
```

- [ ] **Step 2: Make the library multi-targeted and packable**

The package metadata goes in from the start — bolting it on at the end tends to mean discovering a TFM incompatibility after all the code is written.

Replace `src/DocToolkit/DocToolkit.csproj` with:

```xml
<Project Sdk="Microsoft.NET.Sdk">

  <PropertyGroup>
    <TargetFrameworks>net8.0;net10.0</TargetFrameworks>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
    <LangVersion>latest</LangVersion>
    <GenerateDocumentationFile>true</GenerateDocumentationFile>
  </PropertyGroup>

  <!-- NuGet package identity -->
  <PropertyGroup>
    <PackageId>DocToolkit</PackageId>
    <Version>0.1.0</Version>
    <Authors>Khoa Ho</Authors>
    <Description>
      Convert HTML to DOCX and PDF, and open/edit DOCX, XLSX and PPTX - pure managed,
      no native binaries, no browser, no LibreOffice. Runs on Linux.
    </Description>
    <PackageTags>docx;pdf;xlsx;pptx;html;openxml;document;conversion</PackageTags>
    <PackageLicenseExpression>MIT</PackageLicenseExpression>
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

</Project>
```

Then multi-target the test project so both frameworks are actually exercised. In
`tests/DocToolkit.Tests/DocToolkit.Tests.csproj`, change:

```xml
    <TargetFramework>net8.0</TargetFramework>
```

to:

```xml
    <TargetFrameworks>net8.0;net10.0</TargetFrameworks>
```

Create the two files the package references so `dotnet build` does not fail on the missing
`None Include` items. Content is finalised in Task 9 — these are real, working placeholders
for the metadata only, not unfinished work.

`src/DocToolkit/README.md`:

```markdown
# DocToolkit

Convert HTML to DOCX and PDF, and open/edit DOCX, XLSX and PPTX. Pure managed - no native
binaries, no browser, no LibreOffice. Runs on Linux.
```

`src/DocToolkit/THIRD-PARTY-NOTICES.txt`:

```text
DocToolkit bundles the following third-party packages. Full notices are added in Task 9.
```

- [ ] **Step 3: Verify both frameworks build**

Run: `dotnet build`
Expected: **succeeds**, producing `bin/Debug/net8.0/DocToolkit.dll` and `bin/Debug/net10.0/DocToolkit.dll`.

- [ ] **Step 4: Write the failing test**

Create `tests/DocToolkit.Tests/HtmlToDocxConverterTests.cs`:

```csharp
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocToolkit;
using Xunit;

namespace DocToolkit.Tests;

public class HtmlToDocxConverterTests
{
    private const string Html = """
        <h1>Quarterly Report</h1>
        <p>Revenue was <strong>up 12%</strong> and costs were <em>flat</em>.</p>
        <table border="1"><tr><th>Region</th><th>Total</th></tr>
        <tr><td>North</td><td>1200</td></tr></table>
        <ul><li>First</li><li>Second</li></ul>
        <p><a href="https://example.com/report">Full report</a></p>
        """;

    [Fact]
    public async Task ConvertAsync_ProducesAValidDocxPackage()
    {
        var bytes = await HtmlToDocxConverter.ConvertAsync(Html);

        Assert.NotEmpty(bytes);
        // A .docx is a ZIP: it must start with the local file header magic "PK\x03\x04".
        Assert.Equal(new byte[] { 0x50, 0x4B, 0x03, 0x04 }, bytes.Take(4).ToArray());
    }

    [Fact]
    public async Task ConvertAsync_PreservesStructureAndFormatting()
    {
        var bytes = await HtmlToDocxConverter.ConvertAsync(Html);

        using var ms = new MemoryStream(bytes);
        using var doc = WordprocessingDocument.Open(ms, false);
        var body = doc.MainDocumentPart!.Document.Body!;

        Assert.True(body.Descendants<Paragraph>().Count() >= 4);
        Assert.Single(body.Descendants<Table>());
        Assert.Equal(2, body.Descendants<TableRow>().Count());
        Assert.NotEmpty(body.Descendants<Bold>());
        Assert.NotEmpty(body.Descendants<Italic>());
        Assert.NotEmpty(body.Descendants<Hyperlink>());
        Assert.Contains("Quarterly Report", body.InnerText);
    }

    [Fact]
    public async Task ConvertAsync_RejectsNullHtml()
    {
        await Assert.ThrowsAsync<ArgumentNullException>(
            () => HtmlToDocxConverter.ConvertAsync(null!));
    }
}
```

- [ ] **Step 5: Run the test to verify it fails**

Run: `dotnet test --filter FullyQualifiedName~HtmlToDocxConverterTests`
Expected: **build failure** — `The name 'HtmlToDocxConverter' does not exist in the current context`.

- [ ] **Step 6: Write the exception type**

Create `src/DocToolkit/DocumentConversionException.cs`:

```csharp
namespace DocToolkit;

/// <summary>Thrown when a document conversion fails.</summary>
public sealed class DocumentConversionException : Exception
{
    public DocumentConversionException(string message) : base(message) { }
    public DocumentConversionException(string message, Exception inner) : base(message, inner) { }
}
```

- [ ] **Step 7: Write the minimal implementation**

Create `src/DocToolkit/HtmlToDocxConverter.cs`:

```csharp
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlToOpenXml;

namespace DocToolkit;

/// <summary>Converts an HTML fragment into a Word (.docx) package.</summary>
public static class HtmlToDocxConverter
{
    /// <summary>Converts <paramref name="html"/> to the bytes of a .docx package.</summary>
    public static async Task<byte[]> ConvertAsync(string html, CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(html);
        ct.ThrowIfCancellationRequested();

        using var ms = new MemoryStream();
        try
        {
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());
                var converter = new HtmlConverter(mainPart);
                await converter.ParseBody(html);
                mainPart.Document.Save();
            }
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            throw new DocumentConversionException("Failed to convert HTML to DOCX.", ex);
        }

        // ToArray() is valid after the package is disposed - MemoryStream keeps its buffer.
        return ms.ToArray();
    }

    /// <summary>Converts <paramref name="html"/> and writes the .docx to <paramref name="outputPath"/>.</summary>
    public static async Task ConvertToFileAsync(string html, string outputPath, CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);
        var bytes = await ConvertAsync(html, ct);
        await File.WriteAllBytesAsync(outputPath, bytes, ct);
    }
}
```

- [ ] **Step 8: Run the tests to verify they pass**

Run: `dotnet test --filter FullyQualifiedName~HtmlToDocxConverterTests`
Expected: **PASS**, 6 results — 3 tests × 2 target frameworks.

- [ ] **Step 9: Commit**

```bash
git add implementation/dotnet-doc-libs
git commit -m "feat(doctoolkit): scaffold packable multi-targeted library with HTML to DOCX"
```

---

### Task 2: PdfProbe test helper

Nothing downstream can be tested without this. It must exist before any PDF assertion.

**Files:**
- Create: `tests/DocToolkit.Tests/PdfProbe.cs`
- Modify: `tests/DocToolkit.Tests/DocToolkit.Tests.csproj` (no package changes; reference only)

**Interfaces:**
- Consumes: nothing.
- Produces: `DocToolkit.Tests.PdfProbe.ExtractText(byte[]) -> string`, `PdfProbe.PageCount(byte[]) -> int`, `PdfProbe.TextYPositions(byte[]) -> IReadOnlyList<double>`, `PdfProbe.IsPdf(byte[]) -> bool`.

- [ ] **Step 1: Write the failing test**

Add to a new file `tests/DocToolkit.Tests/PdfProbeTests.cs`:

```csharp
using Xunit;

namespace DocToolkit.Tests;

public class PdfProbeTests
{
    [Fact]
    public void ExtractText_DecodesHexStringTextOperators()
    {
        // "Acme" == 41 63 6D 65 ; "Corp" == 43 6F 72 70
        var pdf = System.Text.Encoding.Latin1.GetBytes(
            "%PDF-1.4\n5 0 obj\n<< /Length 40 >>\nstream\nBT\n<41636D65> Tj\n<436F7270> Tj\nET\nendstream\nendobj\n");

        Assert.Equal("AcmeCorp", PdfProbe.ExtractText(pdf));
    }

    [Fact]
    public void PageCount_ReadsTheCountFromThePageTree()
    {
        var pdf = System.Text.Encoding.Latin1.GetBytes(
            "%PDF-1.4\n1 0 obj\n<< /Type /Pages /Count 7 /Kids [ 7 0 R ] >>\nendobj\n");

        Assert.Equal(7, PdfProbe.PageCount(pdf));
    }

    [Fact]
    public void IsPdf_ChecksTheHeaderMagic()
    {
        Assert.True(PdfProbe.IsPdf(System.Text.Encoding.Latin1.GetBytes("%PDF-1.4\n")));
        Assert.False(PdfProbe.IsPdf(new byte[] { 0x50, 0x4B, 0x03, 0x04 }));
    }
}
```

- [ ] **Step 2: Run to verify it fails**

Run: `dotnet test --filter FullyQualifiedName~PdfProbeTests`
Expected: **build failure** — `The name 'PdfProbe' does not exist in the current context`.

- [ ] **Step 3: Implement the probe**

Create `tests/DocToolkit.Tests/PdfProbe.cs`:

```csharp
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace DocToolkit.Tests;

/// <summary>
/// Reads facts out of a generated PDF for assertions.
///
/// IMPORTANT: OfficeIMO writes content streams UNCOMPRESSED (no /Filter) and emits text as
/// hex-string operators, e.g. "&lt;41636D65&gt; Tj" == "Acme". So neither inflating the streams
/// nor substring-searching the raw bytes finds any text - both return nothing and look exactly
/// like a broken converter. Always go through this helper.
/// </summary>
public static class PdfProbe
{
    private static readonly Regex HexText = new(@"<([0-9A-Fa-f]+)>\s*Tj", RegexOptions.Compiled);
    private static readonly Regex PageTree = new(@"/Type\s*/Pages.*?/Count\s+(\d+)",
                                                 RegexOptions.Compiled | RegexOptions.Singleline);
    private static readonly Regex TextMatrix = new(@"1 0 0 1 [-\d.]+ ([-\d.]+) Tm", RegexOptions.Compiled);

    private static string Raw(byte[] pdf) => Encoding.Latin1.GetString(pdf);

    public static bool IsPdf(byte[] pdf) =>
        pdf.Length >= 5 && Encoding.ASCII.GetString(pdf, 0, 5) == "%PDF-";

    /// <summary>All visible text, in content-stream order.</summary>
    public static string ExtractText(byte[] pdf)
    {
        var sb = new StringBuilder();
        foreach (Match m in HexText.Matches(Raw(pdf)))
        {
            var hex = m.Groups[1].Value;
            if (hex.Length % 2 != 0) continue;
            for (var i = 0; i < hex.Length; i += 2)
                sb.Append((char)Convert.ToInt32(hex.Substring(i, 2), 16));
        }
        return sb.ToString();
    }

    /// <summary>Page count taken from the /Pages tree node.</summary>
    public static int PageCount(byte[] pdf)
    {
        var m = PageTree.Match(Raw(pdf));
        return m.Success ? int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture) : 0;
    }

    /// <summary>Y coordinate of every text-positioning operator. Negative values are drawn off-page.</summary>
    public static IReadOnlyList<double> TextYPositions(byte[] pdf) =>
        TextMatrix.Matches(Raw(pdf))
                  .Select(m => double.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture))
                  .ToList();
}
```

- [ ] **Step 4: Run to verify it passes**

Run: `dotnet test --filter FullyQualifiedName~PdfProbeTests`
Expected: **PASS**, 6 results — 3 tests × 2 target frameworks.

- [ ] **Step 5: Commit**

```bash
git add implementation/dotnet-doc-libs/tests
git commit -m "test(doctoolkit): add PdfProbe helper for PDF assertions"
```

---

### Task 3: DOCX → PDF

**Files:**
- Create: `src/DocToolkit/DocxToPdfConverter.cs`, `tests/DocToolkit.Tests/DocxToPdfConverterTests.cs`
- Modify: `src/DocToolkit/DocToolkit.csproj` (add `OfficeIMO.Word.Pdf`)

**Interfaces:**
- Consumes: `HtmlToDocxConverter.ConvertAsync` (to build test input), `PdfProbe`.
- Produces: `DocToolkit.DocxToPdfConverter.Convert(byte[] docx) -> byte[]` and `ConvertFile(string inputPath, string outputPath) -> void`.

- [ ] **Step 1: Add the package**

```bash
dotnet add src/DocToolkit package OfficeIMO.Word.Pdf
```

- [ ] **Step 2: Write the failing test**

Create `tests/DocToolkit.Tests/DocxToPdfConverterTests.cs`:

```csharp
using DocToolkit;
using Xunit;

namespace DocToolkit.Tests;

public class DocxToPdfConverterTests
{
    [Fact]
    public async Task Convert_ProducesAPdfContainingTheSourceText()
    {
        var docx = await HtmlToDocxConverter.ConvertAsync(
            "<h1>Invoice INV-42</h1><p>Total due: 18,100.00</p>");

        var pdf = DocxToPdfConverter.Convert(docx);

        Assert.True(PdfProbe.IsPdf(pdf));
        var text = PdfProbe.ExtractText(pdf);
        Assert.Contains("Invoice INV-42", text);
        Assert.Contains("18,100.00", text);
    }

    [Fact]
    public async Task Convert_PaginatesLongDocuments()
    {
        var rows = string.Concat(Enumerable.Range(1, 60).Select(i =>
            $"<tr><td>Line item {i} with a reasonably long description</td><td>{i * 950}</td></tr>"));
        var html = $"<h1>Big</h1><table border=\"1\">{rows}</table><p>END-MARKER</p>";

        var pdf = DocxToPdfConverter.Convert(await HtmlToDocxConverter.ConvertAsync(html));

        Assert.True(PdfProbe.PageCount(pdf) > 1, "expected the document to span multiple pages");
        Assert.Contains("END-MARKER", PdfProbe.ExtractText(pdf));
        Assert.DoesNotContain(PdfProbe.TextYPositions(pdf), y => y < 0);
    }

    [Fact]
    public void Convert_RejectsEmptyInput()
    {
        Assert.Throws<ArgumentException>(() => DocxToPdfConverter.Convert(Array.Empty<byte>()));
    }
}
```

- [ ] **Step 3: Run to verify it fails**

Run: `dotnet test --filter FullyQualifiedName~DocxToPdfConverterTests`
Expected: **build failure** — `The name 'DocxToPdfConverter' does not exist in the current context`.

- [ ] **Step 4: Write the implementation**

Create `src/DocToolkit/DocxToPdfConverter.cs`:

```csharp
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

namespace DocToolkit;

/// <summary>Renders a Word (.docx) package to PDF. Pure managed - no browser, no LibreOffice.</summary>
public static class DocxToPdfConverter
{
    /// <summary>Renders the .docx in <paramref name="docx"/> and returns PDF bytes.</summary>
    public static byte[] Convert(byte[] docx)
    {
        ArgumentNullException.ThrowIfNull(docx);
        if (docx.Length == 0)
            throw new ArgumentException("DOCX content was empty.", nameof(docx));

        try
        {
            // Copy into an expandable stream: OfficeIMO opens the package read/write.
            using var input = new MemoryStream();
            input.Write(docx, 0, docx.Length);
            input.Position = 0;

            using var word = WordDocument.Load(input);
            return word.ToPdf();
        }
        catch (Exception ex)
        {
            throw new DocumentConversionException("Failed to render DOCX to PDF.", ex);
        }
    }

    /// <summary>Renders <paramref name="inputPath"/> to a PDF at <paramref name="outputPath"/>.</summary>
    public static void ConvertFile(string inputPath, string outputPath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(inputPath);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);
        File.WriteAllBytes(outputPath, Convert(File.ReadAllBytes(inputPath)));
    }
}
```

**If `WordDocument.Load(Stream)` or `word.ToPdf()` does not resolve,** the stream overload or the parameterless extension may differ in your installed version. Use this file-based form instead — it is the exact call proven in `spike/Program.cs`:

```csharp
    public static byte[] Convert(byte[] docx)
    {
        ArgumentNullException.ThrowIfNull(docx);
        if (docx.Length == 0)
            throw new ArgumentException("DOCX content was empty.", nameof(docx));

        var tempDocx = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.docx");
        var tempPdf = Path.ChangeExtension(tempDocx, ".pdf");
        try
        {
            File.WriteAllBytes(tempDocx, docx);
            using (var word = WordDocument.Load(tempDocx))
            {
                var result = word.SaveAsPdf(tempPdf);
                if (!result.Succeeded)
                    throw new DocumentConversionException("PDF rendering reported failure.");
            }
            return File.ReadAllBytes(tempPdf);
        }
        catch (Exception ex) when (ex is not DocumentConversionException)
        {
            throw new DocumentConversionException("Failed to render DOCX to PDF.", ex);
        }
        finally
        {
            if (File.Exists(tempDocx)) File.Delete(tempDocx);
            if (File.Exists(tempPdf)) File.Delete(tempPdf);
        }
    }
```

- [ ] **Step 5: Run to verify it passes**

Run: `dotnet test --filter FullyQualifiedName~DocxToPdfConverterTests`
Expected: **PASS**, 6 results — 3 tests × 2 target frameworks.

- [ ] **Step 6: Commit**

```bash
git add implementation/dotnet-doc-libs
git commit -m "feat(doctoolkit): add pure-managed DOCX to PDF rendering"
```

---

### Task 4: HTML → PDF (composition)

**Files:**
- Create: `src/DocToolkit/HtmlToPdfConverter.cs`, `tests/DocToolkit.Tests/HtmlToPdfConverterTests.cs`

**Interfaces:**
- Consumes: `HtmlToDocxConverter.ConvertAsync(string, CancellationToken) -> Task<byte[]>`, `DocxToPdfConverter.Convert(byte[]) -> byte[]`, `PdfProbe`.
- Produces: `DocToolkit.HtmlToPdfConverter.ConvertAsync(string html, CancellationToken ct = default) -> Task<byte[]>`.

- [ ] **Step 1: Write the failing test**

Create `tests/DocToolkit.Tests/HtmlToPdfConverterTests.cs`:

```csharp
using DocToolkit;
using Xunit;

namespace DocToolkit.Tests;

public class HtmlToPdfConverterTests
{
    [Fact]
    public async Task ConvertAsync_ProducesAPdfFromHtmlInOneCall()
    {
        var pdf = await HtmlToPdfConverter.ConvertAsync(
            "<h1>Statement</h1><p>Balance: 4,250.00</p>");

        Assert.True(PdfProbe.IsPdf(pdf));
        var text = PdfProbe.ExtractText(pdf);
        Assert.Contains("Statement", text);
        Assert.Contains("4,250.00", text);
    }

    [Fact]
    public async Task ConvertAsync_MatchesTheTwoStepPipeline()
    {
        const string html = "<h1>Same Input</h1><p>Same output text.</p>";

        var direct = await HtmlToPdfConverter.ConvertAsync(html);
        var stepwise = DocxToPdfConverter.Convert(await HtmlToDocxConverter.ConvertAsync(html));

        // Byte equality is not guaranteed (timestamps/ids), but the rendered text must match.
        Assert.Equal(PdfProbe.ExtractText(stepwise), PdfProbe.ExtractText(direct));
    }

    [Fact]
    public async Task ConvertAsync_RejectsNullHtml()
    {
        await Assert.ThrowsAsync<ArgumentNullException>(
            () => HtmlToPdfConverter.ConvertAsync(null!));
    }
}
```

- [ ] **Step 2: Run to verify it fails**

Run: `dotnet test --filter FullyQualifiedName~HtmlToPdfConverterTests`
Expected: **build failure** — `The name 'HtmlToPdfConverter' does not exist in the current context`.

- [ ] **Step 3: Write the implementation**

Create `src/DocToolkit/HtmlToPdfConverter.cs`:

```csharp
namespace DocToolkit;

/// <summary>
/// Converts HTML to PDF by pivoting through DOCX.
///
/// There is no permissive, NuGet-only, Linux-safe library that renders HTML to PDF directly:
/// the only free renderers are browsers, and a browser is a native binary. Pivoting through
/// DOCX keeps the whole chain pure managed. See learning-docs/dotnet-doc-libs/report.html.
/// </summary>
public static class HtmlToPdfConverter
{
    /// <summary>Converts <paramref name="html"/> straight to PDF bytes.</summary>
    public static async Task<byte[]> ConvertAsync(string html, CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(html);
        var docx = await HtmlToDocxConverter.ConvertAsync(html, ct);
        ct.ThrowIfCancellationRequested();
        return DocxToPdfConverter.Convert(docx);
    }

    /// <summary>Converts <paramref name="html"/> and writes the PDF to <paramref name="outputPath"/>.</summary>
    public static async Task ConvertToFileAsync(string html, string outputPath, CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);
        var pdf = await ConvertAsync(html, ct);
        await File.WriteAllBytesAsync(outputPath, pdf, ct);
    }
}
```

- [ ] **Step 4: Run to verify it passes**

Run: `dotnet test --filter FullyQualifiedName~HtmlToPdfConverterTests`
Expected: **PASS**, 6 results — 3 tests × 2 target frameworks.

- [ ] **Step 5: Commit**

```bash
git add implementation/dotnet-doc-libs
git commit -m "feat(doctoolkit): add HTML to PDF via the DOCX pivot"
```

---

### Task 5: Open and edit DOCX

Covers the "open/edit docx" requirement: load an existing template and replace `{{placeholder}}` tokens.

**Files:**
- Create: `src/DocToolkit/DocxEditor.cs`, `tests/DocToolkit.Tests/DocxEditorTests.cs`

**Interfaces:**
- Consumes: `HtmlToDocxConverter.ConvertAsync` (to build test input).
- Produces: `DocToolkit.DocxEditor.ReplaceText(byte[] docx, IReadOnlyDictionary<string, string> replacements) -> byte[]`, `DocxEditor.ExtractText(byte[] docx) -> string`.

- [ ] **Step 1: Write the failing test**

Create `tests/DocToolkit.Tests/DocxEditorTests.cs`:

```csharp
using DocToolkit;
using Xunit;

namespace DocToolkit.Tests;

public class DocxEditorTests
{
    [Fact]
    public async Task ReplaceText_SubstitutesPlaceholders()
    {
        var docx = await HtmlToDocxConverter.ConvertAsync(
            "<p>Dear {{name}}, your balance is {{balance}}.</p>");

        var edited = DocxEditor.ReplaceText(docx, new Dictionary<string, string>
        {
            ["{{name}}"] = "Contoso Ltd",
            ["{{balance}}"] = "4,250.00",
        });

        var text = DocxEditor.ExtractText(edited);
        Assert.Contains("Contoso Ltd", text);
        Assert.Contains("4,250.00", text);
        Assert.DoesNotContain("{{name}}", text);
        Assert.DoesNotContain("{{balance}}", text);
    }

    [Fact]
    public async Task ReplaceText_LeavesTheDocumentOpenable()
    {
        var docx = await HtmlToDocxConverter.ConvertAsync("<p>Hello {{who}}</p>");

        var edited = DocxEditor.ReplaceText(docx,
            new Dictionary<string, string> { ["{{who}}"] = "world" });

        // Still a valid package, and still renders.
        Assert.Equal(new byte[] { 0x50, 0x4B, 0x03, 0x04 }, edited.Take(4).ToArray());
        Assert.Contains("world", PdfProbe.ExtractText(DocxToPdfConverter.Convert(edited)));
    }

    [Fact]
    public async Task ExtractText_ReturnsDocumentText()
    {
        var docx = await HtmlToDocxConverter.ConvertAsync("<h1>Title</h1><p>Body copy.</p>");
        var text = DocxEditor.ExtractText(docx);

        Assert.Contains("Title", text);
        Assert.Contains("Body copy.", text);
    }
}
```

- [ ] **Step 2: Run to verify it fails**

Run: `dotnet test --filter FullyQualifiedName~DocxEditorTests`
Expected: **build failure** — `The name 'DocxEditor' does not exist in the current context`.

- [ ] **Step 3: Write the implementation**

Create `src/DocToolkit/DocxEditor.cs`:

```csharp
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocToolkit;

/// <summary>Opens and edits an existing .docx package.</summary>
public static class DocxEditor
{
    /// <summary>
    /// Replaces every key with its value across the document body.
    ///
    /// Word often splits a single visible word across several &lt;w:t&gt; runs (spell-check state,
    /// formatting changes), so a naive per-run replace misses placeholders. This merges the runs
    /// of each paragraph before substituting.
    /// </summary>
    public static byte[] ReplaceText(byte[] docx, IReadOnlyDictionary<string, string> replacements)
    {
        ArgumentNullException.ThrowIfNull(docx);
        ArgumentNullException.ThrowIfNull(replacements);
        if (docx.Length == 0)
            throw new ArgumentException("DOCX content was empty.", nameof(docx));

        using var ms = new MemoryStream();
        ms.Write(docx, 0, docx.Length);
        ms.Position = 0;

        try
        {
            using (var doc = WordprocessingDocument.Open(ms, true))
            {
                var body = doc.MainDocumentPart?.Document?.Body
                           ?? throw new DocumentConversionException("Document has no body.");

                foreach (var paragraph in body.Descendants<Paragraph>())
                {
                    var texts = paragraph.Descendants<Text>().ToList();
                    if (texts.Count == 0) continue;

                    var merged = string.Concat(texts.Select(t => t.Text));
                    var updated = merged;
                    foreach (var (key, value) in replacements)
                        updated = updated.Replace(key, value ?? string.Empty);

                    if (updated == merged) continue;

                    // Put all text on the first run and blank the rest, preserving its formatting.
                    texts[0].Text = updated;
                    texts[0].Space = SpaceProcessingModeValues.Preserve;
                    for (var i = 1; i < texts.Count; i++) texts[i].Text = string.Empty;
                }

                doc.MainDocumentPart!.Document.Save();
            }
        }
        catch (Exception ex) when (ex is not DocumentConversionException)
        {
            throw new DocumentConversionException("Failed to edit DOCX.", ex);
        }

        return ms.ToArray();
    }

    /// <summary>Returns the plain text of the document body.</summary>
    public static string ExtractText(byte[] docx)
    {
        ArgumentNullException.ThrowIfNull(docx);
        using var ms = new MemoryStream(docx);
        using var doc = WordprocessingDocument.Open(ms, false);
        return doc.MainDocumentPart?.Document?.Body?.InnerText ?? string.Empty;
    }
}
```

- [ ] **Step 4: Run to verify it passes**

Run: `dotnet test --filter FullyQualifiedName~DocxEditorTests`
Expected: **PASS**, 6 results — 3 tests × 2 target frameworks.

- [ ] **Step 5: Commit**

```bash
git add implementation/dotnet-doc-libs
git commit -m "feat(doctoolkit): add DOCX open/edit with placeholder replacement"
```

---

### Task 6: Open and edit XLSX

**Files:**
- Create: `src/DocToolkit/WorkbookEditor.cs`, `tests/DocToolkit.Tests/WorkbookEditorTests.cs`
- Modify: `src/DocToolkit/DocToolkit.csproj` (add `ClosedXML`)

**Interfaces:**
- Consumes: nothing from earlier tasks.
- Produces: `DocToolkit.WorkbookEditor.Create(string sheetName, IEnumerable<IEnumerable<object?>> rows) -> byte[]`, `WorkbookEditor.ReadCell(byte[] xlsx, string sheetName, string cellRef) -> string`, `WorkbookEditor.SetCell(byte[] xlsx, string sheetName, string cellRef, object? value) -> byte[]`.

- [ ] **Step 1: Add the package**

```bash
dotnet add src/DocToolkit package ClosedXML
```

- [ ] **Step 2: Write the failing test**

Create `tests/DocToolkit.Tests/WorkbookEditorTests.cs`:

```csharp
using DocToolkit;
using Xunit;

namespace DocToolkit.Tests;

public class WorkbookEditorTests
{
    private static byte[] SampleWorkbook() => WorkbookEditor.Create("Sales", new[]
    {
        new object?[] { "Region", "Total" },
        new object?[] { "North", 1200 },
        new object?[] { "South", 950 },
    });

    [Fact]
    public void Create_ProducesAReadableWorkbook()
    {
        var xlsx = SampleWorkbook();

        Assert.Equal(new byte[] { 0x50, 0x4B, 0x03, 0x04 }, xlsx.Take(4).ToArray());
        Assert.Equal("Region", WorkbookEditor.ReadCell(xlsx, "Sales", "A1"));
        Assert.Equal("North", WorkbookEditor.ReadCell(xlsx, "Sales", "A2"));
        Assert.Equal("1200", WorkbookEditor.ReadCell(xlsx, "Sales", "B2"));
    }

    [Fact]
    public void SetCell_EditsAnExistingWorkbook()
    {
        var edited = WorkbookEditor.SetCell(SampleWorkbook(), "Sales", "B2", 1500);

        Assert.Equal("1500", WorkbookEditor.ReadCell(edited, "Sales", "B2"));
        Assert.Equal("South", WorkbookEditor.ReadCell(edited, "Sales", "A3"));
    }

    [Fact]
    public void ReadCell_ThrowsForAMissingSheet()
    {
        Assert.Throws<DocumentConversionException>(
            () => WorkbookEditor.ReadCell(SampleWorkbook(), "NoSuchSheet", "A1"));
    }
}
```

- [ ] **Step 3: Run to verify it fails**

Run: `dotnet test --filter FullyQualifiedName~WorkbookEditorTests`
Expected: **build failure** — `The name 'WorkbookEditor' does not exist in the current context`.

- [ ] **Step 4: Write the implementation**

Create `src/DocToolkit/WorkbookEditor.cs`:

```csharp
using ClosedXML.Excel;

namespace DocToolkit;

/// <summary>Creates, reads and edits Excel (.xlsx) workbooks. Legacy .xls is not supported.</summary>
public static class WorkbookEditor
{
    /// <summary>Creates a workbook with one sheet populated from <paramref name="rows"/>.</summary>
    public static byte[] Create(string sheetName, IEnumerable<IEnumerable<object?>> rows)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        ArgumentNullException.ThrowIfNull(rows);

        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add(sheetName);

        var r = 1;
        foreach (var row in rows)
        {
            var c = 1;
            foreach (var value in row)
                SetCellValue(sheet.Cell(r, c++), value);
            r++;
        }

        using var ms = new MemoryStream();
        workbook.SaveAs(ms);
        return ms.ToArray();
    }

    /// <summary>Reads a cell as a string. <paramref name="cellRef"/> is an A1-style reference.</summary>
    public static string ReadCell(byte[] xlsx, string sheetName, string cellRef)
    {
        ArgumentNullException.ThrowIfNull(xlsx);
        using var workbook = Open(xlsx);
        return Sheet(workbook, sheetName).Cell(cellRef).GetString();
    }

    /// <summary>Sets a cell and returns the updated workbook bytes.</summary>
    public static byte[] SetCell(byte[] xlsx, string sheetName, string cellRef, object? value)
    {
        ArgumentNullException.ThrowIfNull(xlsx);
        using var workbook = Open(xlsx);
        SetCellValue(Sheet(workbook, sheetName).Cell(cellRef), value);

        using var ms = new MemoryStream();
        workbook.SaveAs(ms);
        return ms.ToArray();
    }

    private static XLWorkbook Open(byte[] xlsx)
    {
        if (xlsx.Length == 0)
            throw new ArgumentException("Workbook content was empty.", nameof(xlsx));
        var ms = new MemoryStream();
        ms.Write(xlsx, 0, xlsx.Length);
        ms.Position = 0;
        return new XLWorkbook(ms);
    }

    private static IXLWorksheet Sheet(XLWorkbook workbook, string sheetName)
    {
        if (!workbook.Worksheets.TryGetWorksheet(sheetName, out var sheet))
            throw new DocumentConversionException($"Worksheet '{sheetName}' was not found.");
        return sheet;
    }

    private static void SetCellValue(IXLCell cell, object? value)
    {
        switch (value)
        {
            case null: cell.Clear(XLClearOptions.Contents); break;
            case string s: cell.Value = s; break;
            case bool b: cell.Value = b; break;
            case DateTime d: cell.Value = d; break;
            case int or long or short or byte or double or float or decimal:
                cell.Value = Convert.ToDouble(value); break;
            default: cell.Value = value.ToString(); break;
        }
    }
}
```

- [ ] **Step 5: Run to verify it passes**

Run: `dotnet test --filter FullyQualifiedName~WorkbookEditorTests`
Expected: **PASS**, 6 results — 3 tests × 2 target frameworks.

- [ ] **Step 6: Commit**

```bash
git add implementation/dotnet-doc-libs
git commit -m "feat(doctoolkit): add XLSX create/read/edit via ClosedXML"
```

---

### Task 7: Open and edit PPTX

**Files:**
- Create: `src/DocToolkit/PresentationEditor.cs`, `tests/DocToolkit.Tests/PresentationEditorTests.cs`
- Modify: `src/DocToolkit/DocToolkit.csproj` (add `ShapeCrawler`)

**Interfaces:**
- Consumes: nothing from earlier tasks.
- Produces: `DocToolkit.PresentationEditor.SlideCount(byte[] pptx) -> int`, `PresentationEditor.ExtractText(byte[] pptx) -> IReadOnlyList<string>`, `PresentationEditor.ReplaceText(byte[] pptx, IReadOnlyDictionary<string, string> replacements) -> byte[]`.

**API note:** ShapeCrawler 0.79.4 — verified by reflection, do not guess. `new Presentation(Stream)`; `pres.Slides` is an `ISlideCollection`; `pres.Slide(int)` is **1-based** and returns `IUserSlide`; `slide.Shapes` is an `IUserSlideShapeCollection`; `shape.TextBox` is `ITextBox?` with `.Text` and `.SetText(string)`; `slide.GetTexts()` returns `IList<string>`; save with `pres.Save(Stream)`. There is no `ISlide` type.

- [ ] **Step 1: Add the package**

```bash
dotnet add src/DocToolkit package ShapeCrawler
```

- [ ] **Step 2: Write the failing test**

Create `tests/DocToolkit.Tests/PresentationEditorTests.cs`:

```csharp
using DocToolkit;
using ShapeCrawler;
using Xunit;

namespace DocToolkit.Tests;

public class PresentationEditorTests
{
    /// <summary>Builds a one-slide deck with a single text box reading "Hello {{who}}".</summary>
    private static byte[] SampleDeck()
    {
        using var pres = new Presentation();
        pres.Slides.Add(1);
        var slide = pres.Slide(1);
        slide.Shapes.AddText(50, 50, 400, 100, "Hello {{who}}");

        using var ms = new MemoryStream();
        pres.Save(ms);
        return ms.ToArray();
    }

    [Fact]
    public void SlideCount_CountsSlides()
    {
        Assert.Equal(1, PresentationEditor.SlideCount(SampleDeck()));
    }

    [Fact]
    public void ExtractText_ReturnsSlideText()
    {
        var texts = PresentationEditor.ExtractText(SampleDeck());
        Assert.Contains(texts, t => t.Contains("Hello {{who}}"));
    }

    [Fact]
    public void ReplaceText_SubstitutesPlaceholders()
    {
        var edited = PresentationEditor.ReplaceText(SampleDeck(),
            new Dictionary<string, string> { ["{{who}}"] = "world" });

        var texts = PresentationEditor.ExtractText(edited);
        Assert.Contains(texts, t => t.Contains("Hello world"));
        Assert.DoesNotContain(texts, t => t.Contains("{{who}}"));
    }
}
```

**If `Shapes.AddText(...)` does not resolve** in your installed version, build the fixture from a file instead: save any one-slide `.pptx` containing the literal text `Hello {{who}}` to `tests/DocToolkit.Tests/assets/sample.pptx`, mark it `<None Update="assets\sample.pptx" CopyToOutputDirectory="PreserveNewest" />` in the test `.csproj`, and replace `SampleDeck()` with `File.ReadAllBytes("assets/sample.pptx")`.

- [ ] **Step 3: Run to verify it fails**

Run: `dotnet test --filter FullyQualifiedName~PresentationEditorTests`
Expected: **build failure** — `The name 'PresentationEditor' does not exist in the current context`.

- [ ] **Step 4: Write the implementation**

Create `src/DocToolkit/PresentationEditor.cs`:

```csharp
using ShapeCrawler;

namespace DocToolkit;

/// <summary>Opens and edits PowerPoint (.pptx) presentations.</summary>
public static class PresentationEditor
{
    /// <summary>Number of slides in the deck.</summary>
    public static int SlideCount(byte[] pptx)
    {
        using var pres = Open(pptx);
        return pres.Slides.Count;
    }

    /// <summary>All text found on every slide, one entry per text-bearing shape.</summary>
    public static IReadOnlyList<string> ExtractText(byte[] pptx)
    {
        using var pres = Open(pptx);
        var results = new List<string>();

        for (var n = 1; n <= pres.Slides.Count; n++)
            results.AddRange(pres.Slide(n).GetTexts());

        return results;
    }

    /// <summary>Replaces every key with its value in all text boxes, returning updated bytes.</summary>
    public static byte[] ReplaceText(byte[] pptx, IReadOnlyDictionary<string, string> replacements)
    {
        ArgumentNullException.ThrowIfNull(replacements);

        using var pres = Open(pptx);
        for (var n = 1; n <= pres.Slides.Count; n++)
        {
            foreach (var shape in pres.Slide(n).Shapes)
            {
                var textBox = shape.TextBox;
                if (textBox is null) continue;

                var original = textBox.Text;
                var updated = original;
                foreach (var (key, value) in replacements)
                    updated = updated.Replace(key, value ?? string.Empty);

                if (updated != original) textBox.SetText(updated);
            }
        }

        using var ms = new MemoryStream();
        pres.Save(ms);
        return ms.ToArray();
    }

    private static Presentation Open(byte[] pptx)
    {
        ArgumentNullException.ThrowIfNull(pptx);
        if (pptx.Length == 0)
            throw new ArgumentException("Presentation content was empty.", nameof(pptx));

        try
        {
            var ms = new MemoryStream();
            ms.Write(pptx, 0, pptx.Length);
            ms.Position = 0;
            return new Presentation(ms);
        }
        catch (Exception ex)
        {
            throw new DocumentConversionException("Failed to open PPTX.", ex);
        }
    }
}
```

- [ ] **Step 5: Run to verify it passes**

Run: `dotnet test --filter FullyQualifiedName~PresentationEditorTests`
Expected: **PASS**, 6 results — 3 tests × 2 target frameworks.

- [ ] **Step 6: Commit**

```bash
git add implementation/dotnet-doc-libs
git commit -m "feat(doctoolkit): add PPTX open/edit via ShapeCrawler"
```

---

### Task 8: Dependency guard and Linux verification

Turns the two standing risks — a Windows-only package sneaking in, and Linux being *inferred* rather than tested — into automated checks.

**Files:**
- Create: `tests/DocToolkit.Tests/DependencyGuardTests.cs`, `Dockerfile.linux-test`

**Interfaces:**
- Consumes: every public type built so far.
- Produces: nothing consumed downstream.

- [ ] **Step 1: Write the failing test**

Create `tests/DocToolkit.Tests/DependencyGuardTests.cs`:

```csharp
using System.Reflection;
using Xunit;

namespace DocToolkit.Tests;

/// <summary>
/// Guards the licence and platform constraints from the research spec. These are the two
/// mistakes that are cheap to make and expensive to discover in production.
/// </summary>
public class DependencyGuardTests
{
    private static readonly string[] BannedAssemblies =
    {
        "System.Drawing.Common",  // throws PlatformNotSupportedException on Linux (.NET 7+)
        "SkiaSharp",              // pulls native binaries
        "EPPlus",                 // Polyform Noncommercial - not free for commercial use
        "NPOI",                   // >= 2.8.0 requires a paid maintenance fee
    };

    [Fact]
    public void DocToolkit_DoesNotReferenceBannedAssemblies()
    {
        var toolkit = typeof(HtmlToDocxConverter).Assembly;
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var queue = new Queue<Assembly>();
        queue.Enqueue(toolkit);

        while (queue.Count > 0)
        {
            foreach (var reference in queue.Dequeue().GetReferencedAssemblies())
            {
                if (!seen.Add(reference.Name!)) continue;
                try { queue.Enqueue(Assembly.Load(reference)); }
                catch { /* not all references load standalone; the name check below still applies */ }
            }
        }

        var violations = seen.Where(name =>
            BannedAssemblies.Any(b => name.Equals(b, StringComparison.OrdinalIgnoreCase))).ToList();

        Assert.True(violations.Count == 0,
            $"Banned assemblies referenced: {string.Join(", ", violations)}");
    }

    [Fact]
    public void NoNativeBinariesAreCopiedToOutput()
    {
        var outputDir = Path.GetDirectoryName(typeof(DependencyGuardTests).Assembly.Location)!;
        var native = Directory.EnumerateFiles(outputDir, "*", SearchOption.AllDirectories)
                              .Where(f => f.EndsWith(".so", StringComparison.OrdinalIgnoreCase)
                                       || f.EndsWith(".dylib", StringComparison.OrdinalIgnoreCase))
                              .ToList();

        Assert.True(native.Count == 0,
            $"Unexpected native binaries in output: {string.Join(", ", native.Select(Path.GetFileName))}");
    }
}
```

- [ ] **Step 2: Run the guard**

Run: `dotnet test --filter FullyQualifiedName~DependencyGuardTests`
Expected: **PASS**, 4 results — 2 tests × 2 target frameworks.

This is a regression guard, not a TDD cycle — it is expected to pass the moment it is written, because the stack chosen in the research is already clean. Its value is failing *later*, if someone adds a banned package. **If it fails now, a banned dependency is already present — remove the package rather than relaxing the test.**

- [ ] **Step 3: Add the Linux verification image**

Create `implementation/dotnet-doc-libs/Dockerfile.linux-test`:

```dockerfile
# Verifies the whole stack really is Linux-safe with no native dependencies.
# Build context is implementation/dotnet-doc-libs/
FROM mcr.microsoft.com/dotnet/sdk:8.0

WORKDIR /src
COPY DocToolkit.sln ./
COPY src/ ./src/
COPY tests/ ./tests/

RUN dotnet restore DocToolkit.sln
CMD ["dotnet", "test", "DocToolkit.sln", "--logger", "console;verbosity=normal"]
```

- [ ] **Step 4: Run the full suite on Linux**

```bash
cd implementation/dotnet-doc-libs
docker build -f Dockerfile.linux-test -t doctoolkit-linux-test .
docker run --rm doctoolkit-linux-test
```

Expected: every test **PASSES** on Linux. A `PlatformNotSupportedException` mentioning `System.Drawing.Common` means a Windows-only dependency crept in.

> If Docker is unavailable, mark this step skipped **and say so** — do not record the Linux constraint as verified when it was not.

- [ ] **Step 5: Run the whole suite once more**

Run: `dotnet test`
Expected: **PASS** — 46 results: 23 tests (3 each for Tasks 1–7, plus 2 guards) × 2 target frameworks.

- [ ] **Step 6: Commit**

```bash
git add implementation/dotnet-doc-libs
git commit -m "test(doctoolkit): guard banned dependencies and verify on Linux"
```

---

### Task 9: Package metadata, licence and third-party notices

Every dependency is permissive, but permissive is not the same as attribution-free — MIT and Apache-2.0 both require the notice to travel with redistributed binaries. Shipping a `.nupkg` *is* redistribution.

**Files:**
- Create: `LICENSE` (repo-relative: `implementation/dotnet-doc-libs/LICENSE`)
- Modify: `src/DocToolkit/README.md`, `src/DocToolkit/THIRD-PARTY-NOTICES.txt` (replace the Task 1 stubs)
- Modify: `.gitignore` (repo root — ignore build artifacts and the local feed)

**Interfaces:**
- Consumes: the finished library from Tasks 1–8.
- Produces: `artifacts/DocToolkit.0.1.0.nupkg` and `artifacts/DocToolkit.0.1.0.snupkg`.

- [ ] **Step 1: Write the licence**

Create `implementation/dotnet-doc-libs/LICENSE` (MIT, matching `PackageLicenseExpression`):

```text
MIT License

Copyright (c) 2026 Khoa Ho

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

- [ ] **Step 2: Write the third-party notices**

Replace `src/DocToolkit/THIRD-PARTY-NOTICES.txt`:

```text
DocToolkit third-party notices
==============================

DocToolkit depends on the following packages. All are permissively licensed and free for
commercial use. Each is redistributed as a transitive NuGet dependency, not vendored.

1. DocumentFormat.OpenXml - MIT License
   Copyright (c) Microsoft Corporation
   https://github.com/dotnet/Open-XML-SDK

2. HtmlToOpenXml.dll - MIT License
   Copyright (c) Olivier Nizet
   https://github.com/onizet/html2openxml

3. OfficeIMO.Word.Pdf (and OfficeIMO.Word, .Pdf, .Drawing, .Security) - MIT License
   Copyright (c) Przemyslaw Klys / Evotec
   https://github.com/EvotecIT/OfficeIMO

4. ClosedXML - MIT License
   Copyright (c) ClosedXML contributors
   https://github.com/ClosedXML/ClosedXML

5. ShapeCrawler - MIT License
   Copyright (c) ShapeCrawler contributors
   https://github.com/ShapeCrawler/ShapeCrawler

Transitively: AngleSharp (MIT), BouncyCastle.Cryptography (MIT),
System.IO.Packaging (MIT), Microsoft.Extensions.Logging.Abstractions (MIT).

The full MIT licence text is reproduced in the LICENSE file shipped with this package.

DELIBERATELY EXCLUDED
---------------------
The following are commonly used for these tasks but are NOT free for commercial use and
must never be added: EPPlus >= 5 (Polyform Noncommercial), NPOI >= 2.8.0 (paid maintenance
fee for revenue-generating users), Spire.* (feature-capped free editions),
Syncfusion.* and QuestPDF (revenue-gated community licences), IronPDF (commercial).
```

- [ ] **Step 3: Write the package README**

Replace `src/DocToolkit/README.md`:

````markdown
# DocToolkit

Convert HTML to DOCX and PDF, and open/edit DOCX, XLSX and PPTX from .NET.

**Pure managed.** No native binaries, no browser, no LibreOffice, no Office interop.
Works after `dotnet restore` alone, and runs on Linux.

## Install

```bash
dotnet add package DocToolkit
```

Targets `net8.0` and `net10.0`.

## Usage

```csharp
using DocToolkit;

// HTML -> DOCX
byte[] docx = await HtmlToDocxConverter.ConvertAsync("<h1>Invoice</h1><p>Total: 18,100.00</p>");

// HTML -> PDF (pivots through DOCX internally)
byte[] pdf = await HtmlToPdfConverter.ConvertAsync("<h1>Invoice</h1><p>Total: 18,100.00</p>");

// DOCX -> PDF
byte[] rendered = DocxToPdfConverter.Convert(docx);

// Fill a DOCX template
byte[] filled = DocxEditor.ReplaceText(docx, new Dictionary<string, string>
{
    ["{{customer}}"] = "Contoso Ltd",
});

// Spreadsheets
byte[] xlsx = WorkbookEditor.Create("Sales", new[]
{
    new object?[] { "Region", "Total" },
    new object?[] { "North", 1200 },
});
string cell = WorkbookEditor.ReadCell(xlsx, "Sales", "B2");

// Presentations
IReadOnlyList<string> slideText = PresentationEditor.ExtractText(pptx);
```

## Why HTML to PDF goes through DOCX

No permissively-licensed, NuGet-only library renders HTML to PDF on Linux: the only free
renderers are browsers, and a browser is a native binary. Pivoting through DOCX keeps the
whole chain pure managed.

## Licence

MIT. See `THIRD-PARTY-NOTICES.txt` for dependency attribution.
````

- [ ] **Step 4: Ignore build artifacts**

Append to the repo-root `.gitignore`:

```gitignore
# ─── NuGet package output & local feed ───────────────────────────────
implementation/**/artifacts/
implementation/**/local-feed/
```

- [ ] **Step 5: Pack**

```bash
cd implementation/dotnet-doc-libs
dotnet pack src/DocToolkit/DocToolkit.csproj -c Release
```

Expected: `artifacts/DocToolkit.0.1.0.nupkg` and `artifacts/DocToolkit.0.1.0.snupkg` exist.

- [ ] **Step 6: Verify the package contents**

A `.nupkg` is a ZIP. Check it contains both frameworks and the metadata files:

```bash
python - <<'PY'
import zipfile
z = zipfile.ZipFile('artifacts/DocToolkit.0.1.0.nupkg')
names = z.namelist()
for required in ['lib/net8.0/DocToolkit.dll', 'lib/net10.0/DocToolkit.dll',
                 'README.md', 'THIRD-PARTY-NOTICES.txt', 'DocToolkit.nuspec']:
    print(('OK   ' if required in names else 'MISS ') + required)
nuspec = z.read('DocToolkit.nuspec').decode('utf-8')
for dep in ['DocumentFormat.OpenXml', 'HtmlToOpenXml.dll', 'OfficeIMO.Word.Pdf',
            'ClosedXML', 'ShapeCrawler']:
    print(('OK   dep ' if dep in nuspec else 'MISS dep ') + dep)
print(('OK   ' if '<license type="expression">MIT' in nuspec else 'MISS ') + 'MIT licence expression')
PY
```

Expected: every line starts with `OK`. A missing `lib/net10.0/` means multi-targeting did not take effect; a missing dependency means a `PackageReference` was added to the *test* project instead of the library.

- [ ] **Step 7: Commit**

```bash
git add implementation/dotnet-doc-libs .gitignore
git commit -m "build(doctoolkit): add package metadata, licence and third-party notices"
```

---

### Task 10: Prove the package works when consumed

Referencing a project and restoring a package are different things. This task catches the failures that only appear on the consumer side: a missing dependency, a TFM that does not apply, an internal type that should have been public.

**Files:**
- Create: `local-feed/` (populated from `artifacts/`)
- Create: a throwaway consumer project **outside the repo** (do not commit it)

**Interfaces:**
- Consumes: `artifacts/DocToolkit.0.1.0.nupkg`.
- Produces: nothing — this is a verification gate.

- [ ] **Step 1: Publish to the local folder feed**

```bash
cd implementation/dotnet-doc-libs
mkdir -p local-feed
cp artifacts/DocToolkit.0.1.0.nupkg local-feed/
```

- [ ] **Step 2: Clear any cached copy**

**This step is not optional.** NuGet caches by id+version in the global packages folder. If a
`DocToolkit 0.1.0` was ever restored before, the *old* copy is reused and your changes appear
to have no effect — a genuinely confusing failure.

```bash
dotnet nuget locals http-cache --clear
rm -rf ~/.nuget/packages/doctoolkit
```

On Windows PowerShell the second command is:

```powershell
Remove-Item -Recurse -Force "$env:USERPROFILE\.nuget\packages\doctoolkit" -ErrorAction Ignore
```

- [ ] **Step 3: Create a consumer project and install the package**

```bash
cd /tmp
rm -rf DocToolkitConsumer
dotnet new console -f net8.0 -n DocToolkitConsumer -o DocToolkitConsumer
cd DocToolkitConsumer
dotnet nuget add source "<ABSOLUTE-PATH>/implementation/dotnet-doc-libs/local-feed" -n doctoolkit-local
dotnet add package DocToolkit --version 0.1.0
```

Replace `<ABSOLUTE-PATH>` with the full path to the repo (e.g. `E:/PJ/LnDPrj`). If the source
name already exists, run `dotnet nuget remove source doctoolkit-local` first.

Expected: restore succeeds and pulls in all five dependencies transitively.

- [ ] **Step 4: Write a consumer smoke test**

Replace `/tmp/DocToolkitConsumer/Program.cs`:

```csharp
using DocToolkit;

// Exercise every public entry point exactly as a downstream consumer would.
const string html = "<h1>Consumer Smoke Test</h1><p>Total: <strong>18,100.00</strong></p>";

byte[] docx = await HtmlToDocxConverter.ConvertAsync(html);
Console.WriteLine($"docx           : {docx.Length,8:N0} bytes");

byte[] pdf = await HtmlToPdfConverter.ConvertAsync(html);
Console.WriteLine($"pdf            : {pdf.Length,8:N0} bytes");

byte[] filled = DocxEditor.ReplaceText(
    await HtmlToDocxConverter.ConvertAsync("<p>Dear {{who}}</p>"),
    new Dictionary<string, string> { ["{{who}}"] = "Contoso" });
Console.WriteLine($"docx edited    : {DocxEditor.ExtractText(filled)}");

byte[] xlsx = WorkbookEditor.Create("Sales", new[]
{
    new object?[] { "Region", "Total" },
    new object?[] { "North", 1200 },
});
Console.WriteLine($"xlsx B2        : {WorkbookEditor.ReadCell(xlsx, "Sales", "B2")}");

// Assertions - non-zero exit means the package is broken for consumers.
if (pdf.Length == 0 || System.Text.Encoding.ASCII.GetString(pdf, 0, 5) != "%PDF-")
{
    Console.Error.WriteLine("FAIL: PDF output is not a valid PDF");
    return 1;
}
if (!DocxEditor.ExtractText(filled).Contains("Contoso"))
{
    Console.Error.WriteLine("FAIL: placeholder replacement did not apply");
    return 1;
}
if (WorkbookEditor.ReadCell(xlsx, "Sales", "B2") != "1200")
{
    Console.Error.WriteLine("FAIL: workbook cell mismatch");
    return 1;
}

Console.WriteLine("\nCONSUMER SMOKE TEST OK");
return 0;
```

- [ ] **Step 5: Run it**

```bash
cd /tmp/DocToolkitConsumer
dotnet run
echo "exit code: $?"
```

Expected: prints byte counts, then `CONSUMER SMOKE TEST OK`, exit code `0`.

If a type is reported as inaccessible, it was left `internal` — make it `public` in the library,
re-pack (Task 9 Step 5), and repeat from Step 2 of this task.

- [ ] **Step 6: Clean up the consumer and the feed source**

```bash
cd /tmp/DocToolkitConsumer && dotnet nuget remove source doctoolkit-local
cd /tmp && rm -rf DocToolkitConsumer
```

- [ ] **Step 7: Commit**

```bash
cd <repo root>
git add implementation/dotnet-doc-libs
git commit -m "build(doctoolkit): verify package installs and works from a local feed"
```

> **Do not run `dotnet nuget push`.** Distribution is a local folder feed by decision; publishing
> to a remote feed is irreversible and is yours to trigger.

---

## Spec coverage

| Requirement (from the research report) | Task |
|---|---|
| Free for commercial use — permissive only | Global constraints + Task 8 guard |
| HTML → DOCX | Task 1 |
| HTML → PDF | Task 4 (via Task 1 + Task 3) |
| Open / edit DOCX | Task 5 |
| Open / edit XLSX | Task 6 |
| Open / edit PPTX | Task 7 |
| DOCX → PDF (optional #4) | Task 3 |
| NuGet-only, no native installs | Task 8 (`NoNativeBinariesAreCopiedToOutput`) |
| Runs on Linux | Task 8 (Docker run) |
| **Redistributable as a NuGet package** | Tasks 1 (packable csproj) + 9 (pack) |
| **Reusable from other projects** | Task 10 (consumer smoke test from a real feed) |
| Third-party licence attribution | Task 9 (`THIRD-PARTY-NOTICES.txt` ships in the package) |
| Legacy `.xls` | **Descoped** 2026-07-28 — deliberately no task |

## Known risks

1. **`OfficeIMO.Word.Pdf` is young** (v3.0.3, published 2026-07-27). The spike proved text, tables, colours, links and pagination are correct, but *visual polish is unverified*. If output looks wrong, the fallback is MigraDoc (MIT) for the PDF layer only — Tasks 1, 5, 6, 7 are unaffected because they never touch PDF rendering.
2. **Task 3 and Task 7 each carry one API uncertainty**, flagged inline with complete fallback code. Neither blocks the task.
3. **Docker may be unavailable** for Task 8 step 4. Report it skipped rather than claiming Linux verification.
4. **The package id `DocToolkit` is a common word** and is very likely already taken on nuget.org. This is harmless for a local folder feed — ids only need to be unique per feed. But if you later publish publicly, expect to rename to something owner-prefixed (`KhoaHo.DocToolkit`). Renaming means a new `PackageId`, a new package page, and consumers updating their `PackageReference`, so decide before the first public push, not after.
5. **`net10.0` requires the .NET 10 SDK** (confirmed installed: 10.0.101). Anyone building this on a machine with only the .NET 8 SDK will fail at restore. Drop `net10.0` from `TargetFrameworks` rather than fighting it.
