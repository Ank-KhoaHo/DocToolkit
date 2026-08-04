# Samples Split Per Capability — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development
> (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use
> checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the single eight-section `samples/ConsoleSample` with five focused samples, one per
capability, each with its own `README.md`.

**Architecture:** Six sample projects under `samples/`, sharing a `Directory.Build.props` for common
properties. Each references the **published** NuGet package, never this source. No tests — a sample
is verified by *running* it.

**Tech Stack:** .NET 8, `Ank.DocToolkit` 0.6.0+ (referenced as `Version="*"`).

**Spec:** `docs/2026-08-04-samples-per-capability-design.md`

## Global Constraints

- **Samples reference the published package via `Version="*"`.** Never a `ProjectReference`, never
  a pinned or floored version. NuGet resolves a minimum-version range to the **lowest** satisfying
  version, so a floor silently pins the sample to an old release.
- **`samples/Directory.Build.props` must NOT declare a `PackageReference`.** `MinimalApi`
  deliberately references only the Extensions package, proving the core arrives transitively. A
  shared core reference would destroy that proof. Each project declares its own.
- **Target framework is `net8.0`** for every sample (single, not multi-targeted).
- **`samples/` carries no lockfile.** Do not add one, and do not run `dotnet restore --locked-mode`
  in these projects.
- The build must stay at **0 warnings**: `dotnet build DocToolkit.sln -c Release -warnaserror
  --no-incremental`. `--no-incremental` is mandatory — MSBuild skips unchanged projects and a
  skipped project emits no diagnostics.
- Commit messages follow Conventional Commits (`type(scope)?: description`). Scope these `samples`.
  **Never add a `Co-Authored-By` trailer.**
- **`main` cannot be pushed directly.** Work on `docs/split-samples-per-capability`, one PR.
- **Every sample must be RUN, not merely built.** CI only builds them, so a sample that compiles and
  throws at runtime would ship green. This is the check CI cannot do.

---

### Task 1: `Directory.Build.props` and the `HtmlConversion` sample

**Files:**
- Create: `samples/Directory.Build.props`
- Create: `samples/HtmlConversion/HtmlConversion.csproj`
- Create: `samples/HtmlConversion/Program.cs`
- Create: `samples/HtmlConversion/README.md`
- Modify: `DocToolkit.sln`

**Interfaces:**
- Produces: the `samples/Directory.Build.props` every later sample project relies on, supplying
  `TargetFramework`, `ImplicitUsings`, `Nullable` and `IsPackable`. Later tasks' `.csproj` files
  therefore declare **only** their `PackageReference` and any `Content` items.

- [ ] **Step 1: Create the shared props file**

`samples/Directory.Build.props`:

```xml
<Project>

  <PropertyGroup>
    <TargetFramework>net8.0</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
    <IsPackable>false</IsPackable>
  </PropertyGroup>

  <!--
    Every sample references the PUBLISHED package with Version="*". The reasoning is in
    samples/README.md; the short version is that "*" resolves to the newest published stable
    release, and a version floor does the opposite - NuGet resolves a minimum-version range to
    the LOWEST satisfying version, so a floor pins the sample to an old release while newer ones
    ship past it unexercised.

    Deliberately NOT declared here: the Ank.DocToolkit PackageReference itself. MinimalApi
    references only Ank.DocToolkit.Extensions.DependencyInjection, which proves the core package
    arrives transitively the way a consumer's restore delivers it. Adding an explicit core
    reference to every sample from here would still build if that transitive dependency broke,
    silently destroying the proof. Each project declares its own reference.
  -->

</Project>
```

- [ ] **Step 2: Create the project file**

`samples/HtmlConversion/HtmlConversion.csproj`:

```xml
<Project Sdk="Microsoft.NET.Sdk">

  <PropertyGroup>
    <OutputType>Exe</OutputType>
  </PropertyGroup>

  <ItemGroup>
    <PackageReference Include="Ank.DocToolkit" Version="*" />
  </ItemGroup>

</Project>
```

- [ ] **Step 3: Write the sample**

`samples/HtmlConversion/Program.cs`:

```csharp
using DocToolkit;

// All three conversions in one file on purpose: HTML -> PDF has no direct renderer under this
// package's constraints (the only free ones are browsers, and a browser is a native binary), so it
// pivots through DOCX internally. Seeing the three side by side is what makes that visible.

const string Html = "<h1>Invoice</h1><p>Total: 18,100.00</p>";

Console.WriteLine("HTML conversion");
Console.WriteLine("===============");

byte[] docx = await HtmlToDocxConverter.ConvertAsync(Html);
Console.WriteLine($"\nHTML -> DOCX : {docx.Length,7:N0} bytes");

byte[] pdf = await HtmlToPdfConverter.ConvertAsync(Html);
Console.WriteLine($"HTML -> PDF  : {pdf.Length,7:N0} bytes  (pivots through DOCX internally)");

byte[] rendered = DocxToPdfConverter.Convert(docx);
Console.WriteLine($"DOCX -> PDF  : {rendered.Length,7:N0} bytes  (from the DOCX above)");

Console.WriteLine("\nDone.");
```

- [ ] **Step 4: Write the README**

`samples/HtmlConversion/README.md`:

```markdown
# HTML conversion

Converting **HTML to DOCX**, **HTML to PDF**, and **DOCX to PDF**.

```bash
dotnet run --project samples/HtmlConversion
```

Prints the byte count each conversion produced.

## The non-obvious part

**HTML → PDF pivots through DOCX.** No permissively-licensed, NuGet-only, Linux-safe library
renders HTML to PDF directly — the only free renderers *are* browsers, and a browser is a native
binary this package will not take on. `HtmlToPdfConverter` therefore composes the other two
converters rather than doing anything of its own.

That is why all three conversions live in one sample: split across three folders, the relationship
between them is invisible.
```

- [ ] **Step 5: Add to the solution and build**

```bash
dotnet sln DocToolkit.sln add samples/HtmlConversion/HtmlConversion.csproj
dotnet build DocToolkit.sln -c Release -warnaserror --no-incremental
```

Expected: `0 Warning(s)`, `0 Error(s)`.

- [ ] **Step 6: RUN it — this is the real check**

```bash
dotnet run --project samples/HtmlConversion -c Release
```

Expected: three non-zero byte counts and `Done.`. If it throws, fix it before committing — CI only
builds samples, so a runtime failure here would never be caught again.

- [ ] **Step 7: Commit**

```bash
git add samples/Directory.Build.props samples/HtmlConversion DocToolkit.sln
git commit -m "docs(samples): add the HtmlConversion sample"
```

---

### Task 2: The `DocxTemplating` sample

**Files:**
- Create: `samples/DocxTemplating/DocxTemplating.csproj`
- Create: `samples/DocxTemplating/Program.cs`
- Create: `samples/DocxTemplating/README.md`
- Modify: `DocToolkit.sln`

**Interfaces:**
- Consumes: `samples/Directory.Build.props` from Task 1, which supplies `TargetFramework`,
  `ImplicitUsings`, `Nullable` and `IsPackable`. Do not repeat them.

- [ ] **Step 1: Create the project file**

`samples/DocxTemplating/DocxTemplating.csproj`:

```xml
<Project Sdk="Microsoft.NET.Sdk">

  <PropertyGroup>
    <OutputType>Exe</OutputType>
  </PropertyGroup>

  <ItemGroup>
    <PackageReference Include="Ank.DocToolkit" Version="*" />
  </ItemGroup>

</Project>
```

- [ ] **Step 2: Write the sample**

`samples/DocxTemplating/Program.cs`:

```csharp
using DocToolkit;

Console.WriteLine("DOCX templating");
Console.WriteLine("===============");

// --- Scalars ------------------------------------------------------------------------------
// ReplaceText handles a placeholder even when Word has split it across several runs, which it
// routinely does - {{customer}} is often three separate <w:t> elements.

byte[] template = await HtmlToDocxConverter.ConvertAsync("<p>Customer: {{customer}}</p>");
byte[] filled = DocxEditor.ReplaceText(template, new Dictionary<string, string>
{
    ["{{customer}}"] = "Contoso Ltd",
});

Console.WriteLine($"\nScalar fill  : \"{DocxEditor.ExtractText(filled).Trim()}\"");

// --- Repeating rows -----------------------------------------------------------------------
// A whole invoice: one template row becomes one row per line item, each keeping the template
// row's formatting.

byte[] invoiceTemplate = await HtmlToDocxConverter.ConvertAsync(
    """
    <h1>Invoice for {{customer}}</h1>
    <table border="1">
      <tr><th>Description</th><th>Qty</th><th>Total</th></tr>
      <tr><td>{{item.Desc}}</td><td>{{item.Qty}}</td><td>{{item.Total}}</td></tr>
    </table>
    """);

// ROWS FIRST, THEN SCALARS. Expanding clones the template row, so any scalar substituted
// beforehand is duplicated into every line. This ordering is the whole reason these two
// operations are demonstrated together - see README.md.
byte[] withRows = DocxEditor.FillRows(invoiceTemplate, "item", new[]
{
    new Dictionary<string, string> { ["Desc"] = "Widget",    ["Qty"] = "2", ["Total"] = "19.98" },
    new Dictionary<string, string> { ["Desc"] = "Gadget",    ["Qty"] = "5", ["Total"] = "45.00" },
    new Dictionary<string, string> { ["Desc"] = "Doohickey", ["Qty"] = "1", ["Total"] = "7.50" },
});

byte[] invoice = DocxEditor.ReplaceText(withRows, new Dictionary<string, string>
{
    ["{{customer}}"] = "Contoso Ltd",
});

string invoiceText = DocxEditor.ExtractText(invoice);
string[] descriptions = { "Widget", "Gadget", "Doohickey" };
int lineCount = descriptions.Count(invoiceText.Contains);

Console.WriteLine($"Line items   : {lineCount} rows from one template row");
Console.WriteLine($"Customer set : {invoiceText.Contains("Contoso Ltd")}");
Console.WriteLine($"Placeholders left over: {invoiceText.Contains("{{item.")}");

Console.WriteLine("\nDone.");
```

- [ ] **Step 3: Write the README**

`samples/DocxTemplating/README.md`:

```markdown
# DOCX templating

Filling a Word template: **scalar placeholders** with `ReplaceText`, and **one table row per
record** with `FillRows`.

```bash
dotnet run --project samples/DocxTemplating
```

Prints the extracted text of the filled document, how many line items one template row became, and
whether any placeholder survived.

## The non-obvious part

**Call `FillRows` before `ReplaceText`, not the other way round.**

`FillRows` clones the template row once per record. Any document-level scalar you substituted
*first* gets cloned along with it, so `{{customer}}` ends up repeated inside every line item
instead of appearing once in the heading. Rows first, then scalars.

That ordering rule is why both operations are shown in one sample: separated, the rule would live
in a folder that does not contain the code it constrains.

Worth knowing: a placeholder is often several `<w:t>` runs in the underlying XML, because Word
splits text as you type. Both methods handle that — a naive per-run `string.Replace` would not.
```

- [ ] **Step 4: Add to the solution and build**

```bash
dotnet sln DocToolkit.sln add samples/DocxTemplating/DocxTemplating.csproj
dotnet build DocToolkit.sln -c Release -warnaserror --no-incremental
```

Expected: `0 Warning(s)`, `0 Error(s)`.

- [ ] **Step 5: RUN it**

```bash
dotnet run --project samples/DocxTemplating -c Release
```

Expected: `Line items   : 3`, `Customer set : True`, `Placeholders left over: False`. **If line
items is not 3, or placeholders is True, the sample is wrong** — do not commit it.

- [ ] **Step 6: Commit**

```bash
git add samples/DocxTemplating DocToolkit.sln
git commit -m "docs(samples): add the DocxTemplating sample"
```

---

### Task 3: The `DocxImages` sample

**Files:**
- Create: `samples/DocxImages/DocxImages.csproj`
- Create: `samples/DocxImages/Program.cs`
- Create: `samples/DocxImages/README.md`
- Modify: `DocToolkit.sln`

**Interfaces:**
- Consumes: `samples/Directory.Build.props` from Task 1.

The old `ConsoleSample` pulled its image bytes out of `sample.pptx` — clever while the PPTX demo
lived in the same project, confusing now that it does not. This sample embeds a tiny PNG as a
base64 constant instead, so it carries no binary asset and needs no fixture from `tests/`.

- [ ] **Step 1: Create the project file**

`samples/DocxImages/DocxImages.csproj`:

```xml
<Project Sdk="Microsoft.NET.Sdk">

  <PropertyGroup>
    <OutputType>Exe</OutputType>
  </PropertyGroup>

  <ItemGroup>
    <PackageReference Include="Ank.DocToolkit" Version="*" />
  </ItemGroup>

</Project>
```

- [ ] **Step 2: Write the sample**

`samples/DocxImages/Program.cs`. The base64 string is a real 137-byte 64x64 solid-colour PNG —
use it exactly as written:

```csharp
using DocToolkit;

Console.WriteLine("DOCX images");
Console.WriteLine("===========");

// A real 64x64 PNG, 137 bytes, inline as base64 so this sample carries no binary asset and needs
// no fixture from anywhere else in the repo. Any PNG or JPEG works the same way.
const string LogoBase64 =
    "iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAIAAAAlC+aJAAAAUElEQVR42u3PQQkAAAgEsOtlFCsZ2gi+hcEKLNXzWgQEBA" +
    "QEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQErsACwGghD5ay/wAAAAAASUVORK5CYII=";

byte[] logo = Convert.FromBase64String(LogoBase64);

byte[] letterhead = await HtmlToDocxConverter.ConvertAsync(
    "<p>{{logo}}</p><p>Dear {{customer}}, please find your invoice attached.</p>");

// Size is in points. Give one dimension and the other scales to preserve the aspect ratio; give
// neither and the image's own header decides, read at 96 DPI.
byte[] branded = DocxEditor.ReplaceImage(letterhead, "{{logo}}", logo, widthPoints: 96);

branded = DocxEditor.ReplaceText(branded, new Dictionary<string, string>
{
    ["{{customer}}"] = "Contoso Ltd",
});

string text = DocxEditor.ExtractText(branded);

Console.WriteLine($"\nLogo         : {logo.Length:N0}-byte PNG, placed 96pt wide");
Console.WriteLine($"Document grew: {letterhead.Length:N0} -> {branded.Length:N0} bytes");
Console.WriteLine($"Placeholder replaced: {!text.Contains("{{logo}}")}");
Console.WriteLine($"Customer set : {text.Contains("Contoso Ltd")}");

Console.WriteLine("\nDone.");
```

- [ ] **Step 3: Write the README**

`samples/DocxImages/README.md`:

```markdown
# DOCX images

Replacing a placeholder with a picture — a logo, a signature, a QR code — using `ReplaceImage`.

```bash
dotnet run --project samples/DocxImages
```

Prints the image size, how much the document grew, and whether the placeholder is gone.

## The non-obvious part

**Sizing is in points, and you usually want to give only one dimension.** Supply `widthPoints` and
the height scales to preserve the aspect ratio. Supply neither and the image's own header decides,
read at 96 DPI.

**The format is decided by the image's magic bytes, never by a filename.** PNG and JPEG are
supported. A file that claims to be a PNG while holding JPEG bytes renders as a blank frame in
Word, silently — so the bytes are what count.

The logo here is a 137-byte PNG inlined as base64, so this sample carries no binary file. Any real
PNG or JPEG works identically.
```

- [ ] **Step 4: Add to the solution and build**

```bash
dotnet sln DocToolkit.sln add samples/DocxImages/DocxImages.csproj
dotnet build DocToolkit.sln -c Release -warnaserror --no-incremental
```

Expected: `0 Warning(s)`, `0 Error(s)`.

- [ ] **Step 5: RUN it**

```bash
dotnet run --project samples/DocxImages -c Release
```

Expected: `Logo         : 137-byte PNG`, a document that grew, `Placeholder replaced: True`,
`Customer set : True`. **If `Placeholder replaced` is False the image was not embedded** — fix
before committing.

- [ ] **Step 6: Commit**

```bash
git add samples/DocxImages DocToolkit.sln
git commit -m "docs(samples): add the DocxImages sample"
```

---

### Task 4: The `Spreadsheets` sample

**Files:**
- Create: `samples/Spreadsheets/Spreadsheets.csproj`
- Create: `samples/Spreadsheets/Program.cs`
- Create: `samples/Spreadsheets/README.md`
- Modify: `DocToolkit.sln`

**Interfaces:**
- Consumes: `samples/Directory.Build.props` from Task 1.

This sample also picks up `SheetNames` and `ReadSheet`, which shipped in **v0.6.0**. They could not
be demonstrated when they were written, because samples reference the published package and the
methods had not been released yet.

- [ ] **Step 1: Create the project file**

`samples/Spreadsheets/Spreadsheets.csproj`:

```xml
<Project Sdk="Microsoft.NET.Sdk">

  <PropertyGroup>
    <OutputType>Exe</OutputType>
  </PropertyGroup>

  <ItemGroup>
    <PackageReference Include="Ank.DocToolkit" Version="*" />
  </ItemGroup>

</Project>
```

- [ ] **Step 2: Write the sample**

`samples/Spreadsheets/Program.cs`:

```csharp
using DocToolkit;

Console.WriteLine("Spreadsheets");
Console.WriteLine("============");

// --- Create, read one cell, edit one cell -------------------------------------------------

byte[] xlsx = WorkbookEditor.Create("Sales", new object?[][]
{
    new object?[] { "Region", "Total" },
    new object?[] { "North",  1200 },
    new object?[] { "South",  950 },
});

string before = WorkbookEditor.ReadCell(xlsx, "Sales", "B2");
byte[] updated = WorkbookEditor.SetCell(xlsx, "Sales", "B2", 1500);
string after = WorkbookEditor.ReadCell(updated, "Sales", "B2");

Console.WriteLine($"\nB2 before {before}, after SetCell {after}");

// --- Reading a workbook you were handed ---------------------------------------------------
// The point of these two: you do not need to know the workbook's shape in advance.

IReadOnlyList<string> sheets = WorkbookEditor.SheetNames(updated);
IReadOnlyList<IReadOnlyList<string>> grid = WorkbookEditor.ReadSheet(updated, sheets[0]);

Console.WriteLine($"Sheets       : {string.Join(", ", sheets)}");
Console.WriteLine($"Shape        : {grid.Count} rows x {grid[0].Count} columns");

foreach (var row in grid)
    Console.WriteLine($"  | {string.Join(" | ", row)}");

Console.WriteLine("\nDone.");
```

- [ ] **Step 3: Write the README**

`samples/Spreadsheets/README.md`:

```markdown
# Spreadsheets

Creating an XLSX, reading and writing single cells, and reading a whole sheet you know nothing
about in advance.

```bash
dotnet run --project samples/Spreadsheets
```

Prints a cell before and after an edit, then the sheet names and the full grid.

## The non-obvious part

**`ReadSheet` anchors its result at A1, not at the first cell containing data.** If a sheet's data
starts at C3, that value is at `rows[2][2]` and everything before it is `""`. That keeps
`rows[r][c]` meaning what it looks like it means. Rows are padded to a uniform width, and entirely
blank rows inside the range are kept rather than dropped — dropping them would shift every later
index.

**Cells come back as strings**, produced the same way `ReadCell` produces them, so the two can
never disagree. A formula cell yields its **cached** value: nothing in this package evaluates
formulas.

**`ReadSheet` refuses a sheet spanning more than 2,000,000 cells.** The result is materialised
whole, so its cost tracks the *rectangle*, not how much of it holds data — one stray value in a far
corner of a sheet describes an enormous grid from a file only a few KB on disk. It throws
`DocumentConversionException` naming the actual extent rather than exhausting memory.

**`SheetNames` includes hidden sheets**, in tab order. Hiding a sheet is a presentation choice, not
a privacy boundary.
```

- [ ] **Step 4: Add to the solution and build**

```bash
dotnet sln DocToolkit.sln add samples/Spreadsheets/Spreadsheets.csproj
dotnet build DocToolkit.sln -c Release -warnaserror --no-incremental
```

Expected: `0 Warning(s)`, `0 Error(s)`.

- [ ] **Step 5: RUN it**

```bash
dotnet run --project samples/Spreadsheets -c Release
```

Expected: `B2 before 1200, after SetCell 1500`; `Sheets       : Sales`; `Shape        : 3 rows x 2
columns`; then three printed rows.

**If this fails with a compile error naming `SheetNames` or `ReadSheet`, the published package on
nuget.org is older than 0.6.0.** Check with `dotnet list samples/Spreadsheets package`. Do not
switch to a `ProjectReference` to work around it — that would defeat the point of the sample. Stop
and report it instead.

- [ ] **Step 6: Commit**

```bash
git add samples/Spreadsheets DocToolkit.sln
git commit -m "docs(samples): add the Spreadsheets sample"
```

---

### Task 5: The `Presentations` sample

**Files:**
- Create: `samples/Presentations/Presentations.csproj`
- Create: `samples/Presentations/Program.cs`
- Create: `samples/Presentations/README.md`
- Modify: `DocToolkit.sln`

**Interfaces:**
- Consumes: `samples/Directory.Build.props` from Task 1.

This is the **only** sample that links `tests/DocToolkit.Tests/assets/sample.pptx`. There is no
"create a PPTX from scratch" method in the public API, so a real fixture is the only way to
demonstrate reading one.

- [ ] **Step 1: Create the project file**

`samples/Presentations/Presentations.csproj`:

```xml
<Project Sdk="Microsoft.NET.Sdk">

  <PropertyGroup>
    <OutputType>Exe</OutputType>
  </PropertyGroup>

  <ItemGroup>
    <PackageReference Include="Ank.DocToolkit" Version="*" />
  </ItemGroup>

  <!--
    There is no "create a PPTX from scratch" method in the public API, so demonstrating PPTX
    reading needs a real file. This borrows the test project's fixture rather than committing a
    second copy. If that fixture ever moves, this fails with an opaque MSBuild copy error rather
    than a message naming the real cause.
  -->
  <ItemGroup>
    <Content Include="..\..\tests\DocToolkit.Tests\assets\sample.pptx"
             Link="assets\sample.pptx"
             CopyToOutputDirectory="PreserveNewest" />
  </ItemGroup>

</Project>
```

- [ ] **Step 2: Write the sample**

`samples/Presentations/Program.cs`:

```csharp
using DocToolkit;

Console.WriteLine("Presentations");
Console.WriteLine("=============");

string path = Path.Combine(AppContext.BaseDirectory, "assets", "sample.pptx");
byte[] pptx = await File.ReadAllBytesAsync(path);

int slides = PresentationEditor.SlideCount(pptx);
IReadOnlyList<string> text = PresentationEditor.ExtractText(pptx);

Console.WriteLine($"\nSlides       : {slides}");
Console.WriteLine($"First slide  : \"{(text.Count > 0 ? text[0] : "(empty)")}\"");

byte[] edited = PresentationEditor.ReplaceText(pptx, new Dictionary<string, string>
{
    ["{{who}}"] = "World",
});

IReadOnlyList<string> editedText = PresentationEditor.ExtractText(edited);
Console.WriteLine($"After replace: \"{(editedText.Count > 0 ? editedText[0] : "(empty)")}\"");

Console.WriteLine("\nDone.");
```

- [ ] **Step 3: Write the README**

`samples/Presentations/README.md`:

```markdown
# Presentations

Reading a PowerPoint file: slide count, text extraction, and placeholder replacement.

```bash
dotnet run --project samples/Presentations
```

Prints the slide count and the first slide's text, before and after a replacement.

## The non-obvious part

**`ExtractText` returns one entry per slide, in deck order** — not one blob of text.

**PowerPoint splits words across runs**, exactly as Word does. A single visible `{{who}}` is often
several `<a:t>` elements in the underlying XML, so a naive per-run `string.Replace` would miss it.
`ReplaceText` maps matches back onto the individual runs they overlap, which is what preserves
per-run formatting.

**There is no "create a PPTX from scratch" method**, so this sample reads a committed fixture —
the test project's `sample.pptx`, borrowed rather than duplicated.
```

- [ ] **Step 4: Add to the solution and build**

```bash
dotnet sln DocToolkit.sln add samples/Presentations/Presentations.csproj
dotnet build DocToolkit.sln -c Release -warnaserror --no-incremental
```

Expected: `0 Warning(s)`, `0 Error(s)`.

- [ ] **Step 5: RUN it**

```bash
dotnet run --project samples/Presentations -c Release
```

Expected: a non-zero slide count and first-slide text. The fixture's text is `Hello {{who}}`, so
`After replace` should read `Hello World`. **If it still shows `{{who}}`, replacement did not
work** — fix before committing.

- [ ] **Step 6: Commit**

```bash
git add samples/Presentations DocToolkit.sln
git commit -m "docs(samples): add the Presentations sample"
```

---

### Task 6: Rename `MinimalApiSample` to `MinimalApi`

**Files:**
- Rename: `samples/MinimalApiSample/` → `samples/MinimalApi/`
- Rename: `samples/MinimalApi/MinimalApiSample.csproj` → `samples/MinimalApi/MinimalApi.csproj`
- Modify: `samples/MinimalApi/MinimalApi.csproj`
- Create: `samples/MinimalApi/README.md`
- Modify: `DocToolkit.sln`

**Interfaces:**
- Consumes: `samples/Directory.Build.props` from Task 1, which now supplies `TargetFramework`,
  `ImplicitUsings`, `Nullable` and `IsPackable` — delete those from the csproj so they are not
  declared twice.

- [ ] **Step 1: Rename with git so history follows**

```bash
git mv samples/MinimalApiSample samples/MinimalApi
git mv samples/MinimalApi/MinimalApiSample.csproj samples/MinimalApi/MinimalApi.csproj
```

- [ ] **Step 2: Slim the project file**

Replace the whole contents of `samples/MinimalApi/MinimalApi.csproj` with:

```xml
<Project Sdk="Microsoft.NET.Sdk.Web">

  <!--
    Deliberately references ONLY the extensions package. That proves Ank.DocToolkit arrives
    transitively, the way a consumer's restore delivers it - an explicit core reference here
    would still build if that transitive dependency broke. Version="*" is explained in
    samples/README.md.
  -->
  <ItemGroup>
    <PackageReference Include="Ank.DocToolkit.Extensions.DependencyInjection" Version="*" />
  </ItemGroup>

</Project>
```

Note what changed: the `PropertyGroup` is gone entirely (now inherited from
`Directory.Build.props`), and the old comment's dangling *"see ConsoleSample.csproj"* pointer —
that file is deleted in Task 7 — is replaced.

The `Sdk="Microsoft.NET.Sdk.Web"` attribute stays. It is on the `Project` element, which
`Directory.Build.props` does not touch.

- [ ] **Step 3: Write the README**

`samples/MinimalApi/README.md`. Move the content from `samples/README.md`'s *MinimalApiSample*
section, updating the paths:

```markdown
# Minimal API

An ASP.NET Core minimal API demonstrating `services.AddDocToolkit()` — one endpoint per injected
interface.

```bash
dotnet run --project samples/MinimalApi --urls http://127.0.0.1:5299
```

Then, in another terminal:

```bash
curl -X POST http://127.0.0.1:5299/html-to-docx \
  -H "Content-Type: application/json" \
  -d '{"html":"<h1>Hello</h1>"}' \
  -o output.docx
```

`/html-to-pdf` takes the same `{"html":"..."}` body. The remaining endpoints (`/docx-to-pdf`,
`/docx/extract-text`, `/xlsx/read-cell`, `/pptx/slide-count`) take `{"bytes":"<base64>"}` instead —
`/xlsx/read-cell` also takes `sheet` and `cell`. See `Program.cs` for each endpoint's exact shape.

## The non-obvious part

**`byte[]` fields are base64-encoded JSON strings**, using ASP.NET Core's built-in handling. No
custom serialization is needed in either direction.

**This project references only `Ank.DocToolkit.Extensions.DependencyInjection`**, never the core
package directly. That is deliberate: it proves the core package arrives transitively, exactly as
it would in a consumer's project. Adding an explicit core reference would make the build pass even
if that transitive dependency broke.

**`AllowRemoteImageDownload` is configured once** at `AddDocToolkit(...)`, rather than being
decided per call as the static API does.
```

- [ ] **Step 4: Fix the solution entry**

```bash
dotnet sln DocToolkit.sln remove samples/MinimalApiSample/MinimalApiSample.csproj
dotnet sln DocToolkit.sln add samples/MinimalApi/MinimalApi.csproj
```

If the `remove` errors because the old path is already gone from disk, edit `DocToolkit.sln` by
hand to drop the stale `MinimalApiSample` project block and its `GlobalSection` GUID lines, then
run the `add`.

- [ ] **Step 5: Build and run**

```bash
dotnet build DocToolkit.sln -c Release -warnaserror --no-incremental
```

Expected: `0 Warning(s)`, `0 Error(s)`.

Then prove the DI wiring still resolves, by starting the API in the background, calling it, and
stopping it — all in one command so nothing is left listening:

```bash
dotnet run --project samples/MinimalApi -c Release --urls http://127.0.0.1:5299 &
API_PID=$!
sleep 8
curl -sS -X POST http://127.0.0.1:5299/html-to-docx \
  -H "Content-Type: application/json" -d '{"html":"<h1>Hello</h1>"}' \
  -o samples/MinimalApi/out.docx
ls -l samples/MinimalApi/out.docx
kill $API_PID
rm samples/MinimalApi/out.docx
```

Expected: a non-empty `out.docx` (a DOCX is a ZIP, so a few KB, not a few bytes). A build alone
proves nothing here — `AddDocToolkit()` resolving its six interfaces is a runtime fact.

Write the output next to the project rather than `/tmp`: this repo is worked on from both Git Bash
and Windows tooling, where `/tmp` resolves to two different directories, and a file written to one
is invisible to the other. Delete it afterwards — it must not be committed.

- [ ] **Step 6: Commit**

`git mv` already staged both renames, so this only needs the new README, the edited csproj and the
solution:

```bash
git add samples/MinimalApi DocToolkit.sln
git status --short          # expect R (renamed) lines, no stray out.docx
git commit -m "docs(samples): rename MinimalApiSample to MinimalApi"
```

Check `git status --short` shows the moves as `R`, not as a delete plus an add — that is what keeps
`git log --follow` working on `Program.cs`.

---

### Task 7: Retire `ConsoleSample` and repoint every reference to it

**Files:**
- Delete: `samples/ConsoleSample/`
- Rewrite: `samples/README.md`
- Modify: `README.md`
- Modify: `CLAUDE.md`
- Modify: `DocToolkit.sln`

**Interfaces:**
- Consumes: all five samples from Tasks 1–5 and the rename from Task 6.

Nothing may reference `ConsoleSample` after this task. It is the last task deliberately: until now
the old sample still worked, so no intermediate commit had a dangling path.

- [ ] **Step 1: Delete the project and its solution entry**

```bash
dotnet sln DocToolkit.sln remove samples/ConsoleSample/ConsoleSample.csproj
git rm -r samples/ConsoleSample
```

- [ ] **Step 2: Rewrite `samples/README.md` as an index**

Replace the whole file with:

```markdown
# Samples

Six runnable projects, each answering one question. Every one references the **published** NuGet
packages rather than this repo's source — the same restore an external consumer gets.

| Sample | Answers |
|---|---|
| [HtmlConversion](HtmlConversion/) | How do I turn HTML into a DOCX or a PDF? |
| [DocxTemplating](DocxTemplating/) | How do I fill a Word template, including one row per record? |
| [DocxImages](DocxImages/) | How do I drop a logo or signature into a placeholder? |
| [Spreadsheets](Spreadsheets/) | How do I create, edit and read an XLSX? |
| [Presentations](Presentations/) | How do I read text out of a PowerPoint file? |
| [MinimalApi](MinimalApi/) | How do I wire this into ASP.NET Core dependency injection? |

Each folder has its own `README.md` with the command to run it and the one thing about that
capability that is not obvious.

## Why they reference the published packages

Every sample uses `<PackageReference ... Version="*" />`, never a `ProjectReference`. That makes
them a canary: they exercise the artifact a consumer would actually restore, not whatever happens
to be on `main`.

`*` resolves to the newest published stable release, and that non-determinism is the point. A
version **floor** does the opposite — NuGet resolves a minimum-version range to the **lowest**
satisfying version, so `[0.2.1, )` once kept the samples building against 0.2.1 while three
releases shipped past them unexercised. Bumping the floor fixed it for exactly one release and
went stale on the next.

The trade-off: a capability that has been merged but not yet released **cannot appear here** until
it ships. That is not a gap to work around — it is the guarantee working.

Shared build settings live in `Directory.Build.props`. It deliberately declares no package
reference: `MinimalApi` references only the extensions package, which proves the core package
arrives transitively.
```

- [ ] **Step 3: Fix `README.md`**

In the *Build and test* fenced block, replace:

```bash
dotnet run --project samples/ConsoleSample
dotnet run --project samples/MinimalApiSample
```

with:

```bash
dotnet run --project samples/HtmlConversion      # one folder per capability - see samples/README.md
dotnet run --project samples/MinimalApi
```

Then in the *Repository layout* block, replace the `samples/` line with:

```
samples/                                                six runnable samples, one per capability, on the published packages
```

- [ ] **Step 4: Fix `CLAUDE.md`**

Three edits in the *Samples and docs site* section:

1. Replace the opening sentence `samples/ConsoleSample and samples/MinimalApiSample are runnable,
   added to DocToolkit.sln, ...` so it names the new layout instead:

   ```
   The six projects under `samples/` are runnable, added to `DocToolkit.sln`, and reference the
   published packages via `PackageReference` (never `ProjectReference`) — same reasoning as the
   extensions package itself: they prove the real published artifact works, not whatever is
   currently on `main`. They're built by the existing CI `dotnet build` step with no special
   handling; a breaking API change fails the next sample build. Shared build properties live in
   `samples/Directory.Build.props`, which deliberately declares no `PackageReference` — `MinimalApi`
   references only the extensions package, which is what proves the core package arrives
   transitively.
   ```

2. **Delete the final paragraph entirely** — the one beginning `ConsoleSample reaches into
   tests/DocToolkit.Tests/assets/sample.pptx for its PPTX demo`. Replace it with the same fact
   about the sample that now does it:

   ```
   `samples/Presentations` reaches into `tests/DocToolkit.Tests/assets/sample.pptx` — there's no
   "create a PPTX from scratch" method in the public API, so this is a deliberate trade-off. If that
   fixture ever moves, the sample fails with an opaque MSBuild copy error, not a message pointing at
   the real cause. It is the only sample that needs a fixture: `DocxImages` inlines a 137-byte PNG
   as base64 rather than borrowing one.
   ```

3. In the **Layout** block at the bottom, replace the two `samples/...` lines with:

   ```
   samples/                                                six runnable samples, one per capability
   ```

- [ ] **Step 5: Verify nothing references the old paths**

```bash
grep -rn "ConsoleSample\|MinimalApiSample" --exclude-dir=.git --exclude-dir=obj --exclude-dir=bin . || echo "CLEAN"
```

Expected: `CLEAN`, apart from hits inside `docs/` (historical design documents, which describe what
was true when written and are not navigation) and `CHANGELOG.md` (immutable release history). **A
hit in `README.md`, `CLAUDE.md`, `samples/`, or any `.csproj` is a bug** — fix it.

- [ ] **Step 6: Full build, then run every sample**

```bash
dotnet build DocToolkit.sln -c Release -warnaserror --no-incremental
dotnet test DocToolkit.sln -c Release
for s in HtmlConversion DocxTemplating DocxImages Spreadsheets Presentations; do
  echo "=== $s ==="
  dotnet run --project samples/$s -c Release || echo "FAILED: $s"
done
```

Expected: `0 Warning(s)`; the full suite green; and all five samples printing `Done.` with no
`FAILED` line. The test count is unchanged by this work — samples carry no tests — so do not touch
the counts in `README.md` or `CLAUDE.md`.

- [ ] **Step 7: Commit and open the pull request**

```bash
git add -A
git commit -m "docs(samples): retire ConsoleSample in favour of per-capability samples"
git push -u origin docs/split-samples-per-capability
gh pr create --base main --title "docs(samples): split the samples per capability"
```

Write the PR body to say: six folders replacing two, one README each, `ConsoleSample` retired,
`MinimalApiSample` renamed, and that every sample was run rather than only built.

---

## Notes for the implementer

**Do not switch any sample to a `ProjectReference`**, however tempting. The whole value of these
projects is that they restore the published artifact. If a sample cannot compile against the
published package, that is information — report it rather than working around it.

**Do not add a lockfile under `samples/`.** The premise guard in CI checks lockfiles for the two
packable projects only, and `samples/` deliberately has none.

**Run every sample you touch.** CI builds samples but never runs them, so a runtime failure
introduced here would not be caught by anything else, ever.

**The `samples/README.md` rewrite in Task 7 deletes the MinimalApi section** that Task 6 moved into
`samples/MinimalApi/README.md`. If Task 6 was skipped, that content would be lost — do the tasks in
order.
