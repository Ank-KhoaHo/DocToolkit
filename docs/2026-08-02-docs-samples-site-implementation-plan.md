# Docs/Samples Site Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add two runnable sample projects and a DocFX-generated API-reference site, published to
GitHub Pages automatically whenever a release actually ships.

**Architecture:** `samples/ConsoleSample` and `samples/MinimalApiSample` join `DocToolkit.sln`,
referencing the published packages via `PackageReference` (never `ProjectReference`), so
`ci.yml`'s existing build step verifies them with zero CI file changes. A new `docfx/` folder
holds a DocFX site scaffolded via `docfx init`, configured to read both projects' XML doc
comments. A new `.github/workflows/docs.yml`, triggered by `workflow_run` on `Release` completing
successfully, builds the site and deploys it to GitHub Pages via `actions/deploy-pages`.

**Tech Stack:** .NET 8/10, ASP.NET Core Minimal APIs, DocFX 2.78+, GitHub Actions
(`actions/deploy-pages`, `actions/upload-pages-artifact`).

## Global Constraints

- Both packages are currently published at **0.2.1** on nuget.org. Sample projects reference them
  via `PackageReference Version="[0.2.1, )"` — an open floor range, never a pin, never
  `ProjectReference` — matching how `Ank.DocToolkit.Extensions.DependencyInjection` already
  references `Ank.DocToolkit`.
- **No `Co-Authored-By` trailer in any commit message.** Repo-wide convention, no exceptions.
- Planning/spec docs live at `docs/YYYY-MM-DD-<topic>-*.md`. This plan's own spec is
  `docs/2026-08-01-docs-samples-site-design.md`. The new `docfx/` folder is unrelated to `docs/`
  — one is this project's planning history, the other is published site source. Don't conflate
  them or move one into the other.
- `dotnet build DocToolkit.sln -c Release -warnaserror` must stay at **0 warnings**. Both new
  sample projects join the `.sln`, so this existing gate (already run by `ci.yml`) covers them
  automatically — no `ci.yml` edits are needed for this plan.
- Sample projects target **`net8.0` only** (not the library's `net8.0;net10.0` multi-target — a
  typical consumer app targets one framework) and are `<IsPackable>false</IsPackable>`.
- **DocFX's `globalMetadata` must set `"_enableSearch": false` and `"pdf": false`.** Verified
  empirically while writing this plan: the default scaffold's "modern" template silently
  downloads a ~109 MB headless-Chromium binary via Playwright/Node.js during the build step (for
  client-side search-index extraction and for PDF export), even on a from-scratch `docfx init`.
  Leaving either flag at its scaffolded default (`true`) makes every cold `docs.yml` run slow and
  pulls in a heavyweight native browser dependency for documentation tooling — exactly the kind
  of dependency this repo's premise guards exist to keep out of the shipped packages, and out of
  character even though it wouldn't touch the packages themselves. Do not remove these two flags
  to "restore search" without re-verifying the browser download doesn't come back.
- `docs.yml` triggers via `workflow_run` on `Release` completing, gated on
  `github.event.workflow_run.conclusion == 'success'` — never independently on the same tag push.
  A version that fails `release.yml`'s guards must never get a docs site describing it as shipped.
- GitHub Pages must be enabled once, manually, in the repo's Settings → Pages (source: "GitHub
  Actions"). No task in this plan can do this — it's called out explicitly in Task 4.

---

### Task 1: ConsoleSample project

**Files:**
- Create: `samples/ConsoleSample/ConsoleSample.csproj`
- Create: `samples/ConsoleSample/Program.cs`
- Modify: `DocToolkit.sln` (add project via `dotnet sln add`)

**Interfaces:**
- Consumes: `Ank.DocToolkit` 0.2.1 from nuget.org (public static API: `HtmlToDocxConverter.ConvertAsync`, `HtmlToPdfConverter.ConvertAsync`, `DocxToPdfConverter.Convert`, `DocxEditor.ReplaceText`/`ExtractText`, `WorkbookEditor.Create`/`ReadCell`/`SetCell`, `PresentationEditor.SlideCount`/`ExtractText`). Also reads the existing test fixture `tests/DocToolkit.Tests/assets/sample.pptx` (already present in the repo) via a linked `Content` item, since the public API has no "create a PPTX from scratch" method.
- Produces: nothing consumed by later tasks — this is a leaf project.

- [ ] **Step 1: Create the project directory and csproj**

Create `samples/ConsoleSample/ConsoleSample.csproj`:

```xml
<Project Sdk="Microsoft.NET.Sdk">

  <PropertyGroup>
    <OutputType>Exe</OutputType>
    <TargetFramework>net8.0</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
    <IsPackable>false</IsPackable>
  </PropertyGroup>

  <ItemGroup>
    <PackageReference Include="Ank.DocToolkit" Version="[0.2.1, )" />
  </ItemGroup>

  <ItemGroup>
    <Content Include="..\..\tests\DocToolkit.Tests\assets\sample.pptx" Link="assets\sample.pptx" CopyToOutputDirectory="PreserveNewest" />
  </ItemGroup>

</Project>
```

- [ ] **Step 2: Write Program.cs**

Create `samples/ConsoleSample/Program.cs`:

```csharp
using DocToolkit;

Console.WriteLine("DocToolkit console sample");
Console.WriteLine("==========================");

// 1. HTML -> DOCX
Console.WriteLine("\n1. HTML -> DOCX");
byte[] docx = await HtmlToDocxConverter.ConvertAsync("<h1>Invoice</h1><p>Total: 18,100.00</p>");
Console.WriteLine($"   Generated {docx.Length} bytes of DOCX.");

// 2. HTML -> PDF (pivots through DOCX internally)
Console.WriteLine("\n2. HTML -> PDF");
byte[] pdf = await HtmlToPdfConverter.ConvertAsync("<h1>Invoice</h1><p>Total: 18,100.00</p>");
Console.WriteLine($"   Generated {pdf.Length} bytes of PDF.");

// 3. DOCX -> PDF
Console.WriteLine("\n3. DOCX -> PDF");
byte[] rendered = DocxToPdfConverter.Convert(docx);
Console.WriteLine($"   Rendered {rendered.Length} bytes of PDF from the DOCX above.");

// 4. Fill a DOCX template, then extract text back out
Console.WriteLine("\n4. DOCX template fill + text extraction");
byte[] template = await HtmlToDocxConverter.ConvertAsync("<p>Customer: {{customer}}</p>");
byte[] filled = DocxEditor.ReplaceText(template, new Dictionary<string, string>
{
    ["{{customer}}"] = "Contoso Ltd",
});
string text = DocxEditor.ExtractText(filled);
Console.WriteLine($"   Extracted text: \"{text.Trim()}\"");

// 5. Spreadsheets: create, read a cell, update it, read again
Console.WriteLine("\n5. XLSX create/read/edit");
byte[] xlsx = WorkbookEditor.Create("Sales", new object?[][]
{
    new object?[] { "Region", "Total" },
    new object?[] { "North", 1200 },
});
string cellBefore = WorkbookEditor.ReadCell(xlsx, "Sales", "B2");
byte[] updated = WorkbookEditor.SetCell(xlsx, "Sales", "B2", 1500);
string cellAfter = WorkbookEditor.ReadCell(updated, "Sales", "B2");
Console.WriteLine($"   B2 before: {cellBefore}, after SetCell: {cellAfter}");

// 6. Presentations: read the shared test fixture PPTX
Console.WriteLine("\n6. PPTX read/edit");
string pptxPath = Path.Combine(AppContext.BaseDirectory, "assets", "sample.pptx");
byte[] pptx = await File.ReadAllBytesAsync(pptxPath);
int slideCount = PresentationEditor.SlideCount(pptx);
IReadOnlyList<string> slideText = PresentationEditor.ExtractText(pptx);
string firstSlide = slideText.Count > 0 ? slideText[0] : "(empty)";
Console.WriteLine($"   {slideCount} slide(s); first slide text: \"{firstSlide}\"");

Console.WriteLine("\nDone.");
```

- [ ] **Step 3: Add the project to the solution**

Run: `dotnet sln DocToolkit.sln add samples/ConsoleSample/ConsoleSample.csproj -s samples`
Expected: `Project ... added to the solution.` No error.

- [ ] **Step 4: Restore and build**

Run: `dotnet restore DocToolkit.sln`
Expected: succeeds, resolves `Ank.DocToolkit` 0.2.1 from nuget.org.

Run: `dotnet build DocToolkit.sln -c Release -warnaserror`
Expected: `Build succeeded.`, `0 Warning(s)`, `0 Error(s)`.

- [ ] **Step 5: Run the sample and verify real output**

Run: `dotnet run --project samples/ConsoleSample -c Release`
Expected: output ending with `Done.`, no unhandled exception, and every numbered section prints a
non-zero byte count / non-empty extracted text (in particular, section 4's extracted text should
contain `Contoso Ltd`, and section 6 should report a slide count of at least 1). If any section
throws, do not proceed — the sample must actually run clean before commit.

- [ ] **Step 6: Commit**

```bash
git add samples/ConsoleSample DocToolkit.sln
git commit -m "samples: add ConsoleSample exercising all five core capabilities"
```

---

### Task 2: MinimalApiSample project

**Files:**
- Create: `samples/MinimalApiSample/MinimalApiSample.csproj`
- Create: `samples/MinimalApiSample/Program.cs`
- Modify: `DocToolkit.sln` (add project via `dotnet sln add`)

**Interfaces:**
- Consumes: `Ank.DocToolkit.Extensions.DependencyInjection` 0.2.1 from nuget.org. Exact interface
  signatures (verified against `src/DocToolkit.Extensions.DependencyInjection/*.cs` while writing
  this plan — use these exact types, do not guess a different shape):
  - `IHtmlToDocxConverter.ConvertAsync(string html, CancellationToken ct = default) : Task<byte[]>`
  - `IHtmlToPdfConverter.ConvertAsync(string html, CancellationToken ct = default) : Task<byte[]>`
  - `IDocxToPdfConverter.Convert(byte[] docx) : byte[]`
  - `IDocxEditor.ExtractText(byte[] docx) : string`
  - `IWorkbookEditor.ReadCell(byte[] xlsx, string sheetName, string cellRef) : string`
  - `IPresentationEditor.SlideCount(byte[] pptx) : int`
  - `ServiceCollectionExtensions.AddDocToolkit(this IServiceCollection, Action<DocToolkitOptions>? configure = null)`
- Produces: nothing consumed by later tasks — this is a leaf project.

- [ ] **Step 1: Create the project directory and csproj**

Create `samples/MinimalApiSample/MinimalApiSample.csproj`:

```xml
<Project Sdk="Microsoft.NET.Sdk.Web">

  <PropertyGroup>
    <TargetFramework>net8.0</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
    <IsPackable>false</IsPackable>
  </PropertyGroup>

  <ItemGroup>
    <PackageReference Include="Ank.DocToolkit.Extensions.DependencyInjection" Version="[0.2.1, )" />
  </ItemGroup>

</Project>
```

- [ ] **Step 2: Write Program.cs**

Create `samples/MinimalApiSample/Program.cs`. Six endpoints, one per injected interface, each a
thin wrapper — `byte[]` request/response bodies use ASP.NET Core's built-in base64-in-JSON
handling for `byte[]`, so no custom (de)serialization is needed:

```csharp
using DocToolkit.Extensions.DependencyInjection;

var builder = WebApplication.CreateBuilder(args);
builder.Services.AddDocToolkit();

var app = builder.Build();

app.MapPost("/html-to-docx", async (IHtmlToDocxConverter converter, HtmlRequest request) =>
{
    byte[] docx = await converter.ConvertAsync(request.Html);
    return Results.File(docx, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "output.docx");
});

app.MapPost("/html-to-pdf", async (IHtmlToPdfConverter converter, HtmlRequest request) =>
{
    byte[] pdf = await converter.ConvertAsync(request.Html);
    return Results.File(pdf, "application/pdf", "output.pdf");
});

app.MapPost("/docx-to-pdf", (IDocxToPdfConverter converter, FileRequest request) =>
{
    byte[] pdf = converter.Convert(request.Bytes);
    return Results.File(pdf, "application/pdf", "output.pdf");
});

app.MapPost("/docx/extract-text", (IDocxEditor editor, FileRequest request) =>
{
    string text = editor.ExtractText(request.Bytes);
    return Results.Text(text);
});

app.MapPost("/xlsx/read-cell", (IWorkbookEditor editor, CellRequest request) =>
{
    string value = editor.ReadCell(request.Bytes, request.Sheet, request.Cell);
    return Results.Text(value);
});

app.MapPost("/pptx/slide-count", (IPresentationEditor editor, FileRequest request) =>
{
    int count = editor.SlideCount(request.Bytes);
    return Results.Text(count.ToString());
});

app.Run();

internal sealed record HtmlRequest(string Html);
internal sealed record FileRequest(byte[] Bytes);
internal sealed record CellRequest(byte[] Bytes, string Sheet, string Cell);
```

- [ ] **Step 3: Add the project to the solution**

Run: `dotnet sln DocToolkit.sln add samples/MinimalApiSample/MinimalApiSample.csproj -s samples`
Expected: `Project ... added to the solution.` No error.

- [ ] **Step 4: Restore and build**

Run: `dotnet restore DocToolkit.sln`
Expected: succeeds, resolves `Ank.DocToolkit.Extensions.DependencyInjection` 0.2.1 (and its
`Ank.DocToolkit` dependency) from nuget.org.

Run: `dotnet build DocToolkit.sln -c Release -warnaserror`
Expected: `Build succeeded.`, `0 Warning(s)`, `0 Error(s)`.

- [ ] **Step 5: Run the sample and verify a real endpoint round-trip**

Run the app in the background on a fixed port, then exercise `/html-to-docx`:

```bash
dotnet run --project samples/MinimalApiSample -c Release --urls http://127.0.0.1:5299 &
APP_PID=$!
sleep 5
curl -s -X POST http://127.0.0.1:5299/html-to-docx \
  -H "Content-Type: application/json" \
  -d '{"html":"<p>hello from the sample</p>"}' \
  -o /tmp/sample-output.docx
ls -la /tmp/sample-output.docx
kill $APP_PID
```

Expected: `/tmp/sample-output.docx` exists and is non-trivially sized (a few KB, not 0 bytes — an
empty/near-empty file means the endpoint threw and ASP.NET Core returned a problem-details JSON
body instead of a real docx). If it's empty or tiny, read the app's stdout (don't redirect it away
in this step) before assuming the endpoint is fine.

- [ ] **Step 6: Commit**

```bash
git add samples/MinimalApiSample DocToolkit.sln
git commit -m "samples: add MinimalApiSample demonstrating AddDocToolkit() DI registration"
```

---

### Task 3: DocFX API-reference site

**Files:**
- Create: `docfx/docfx.json` (via `docfx init -y`, then edited)
- Create: `docfx/index.md` (replaces the scaffolded placeholder)
- Create: `docfx/toc.yml` (replaces the scaffolded placeholder — trims the "Docs" section, which
  is out of scope per the approved design)
- Delete: `docfx/docs/introduction.md`, `docfx/docs/getting-started.md`, `docfx/docs/toc.yml`
  (scaffolded by `docfx init` but out of scope — the approved design covers a landing page only,
  not a separate hand-written conceptual-docs section)
- Modify: `.gitignore` (ignore `docfx/_site/` and `docfx/api/` — both fully regenerated by every
  build, verified empirically while writing this plan)

**Interfaces:**
- Consumes: `src/DocToolkit/DocToolkit.csproj` and
  `src/DocToolkit.Extensions.DependencyInjection/DocToolkit.Extensions.DependencyInjection.csproj`
  (both already have `<GenerateDocumentationFile>true</GenerateDocumentationFile>` — no change
  needed there).
- Produces: `docfx/docfx.json`, consumed by Task 4's `docs.yml` (`docfx docfx/docfx.json`).

- [ ] **Step 1: Install DocFX and scaffold the folder**

Run: `dotnet tool update -g docfx`
Expected: installs or confirms DocFX 2.78+ (`docfx --version` should print `2.78.x` or newer
afterward — if `docfx` isn't found on PATH right after install, the global tools directory
(`~/.dotnet/tools` on Linux/macOS runners, already the case for this CI's Ubuntu runners) needs to
be on PATH; GitHub-hosted Ubuntu runners already have this configured by default).

Run, from the repo root:
```bash
mkdir docfx
cd docfx
docfx init -y
cd ..
```
Expected: creates `docfx/docfx.json`, `docfx/index.md`, `docfx/toc.yml`, `docfx/docs/introduction.md`,
`docfx/docs/getting-started.md`, `docfx/docs/toc.yml`. (Note: the flag is `-y`/`--yes`, not `-q` —
`docfx init -q` silently fails with no output and no scaffold in this DocFX version; confirmed
while writing this plan. If `-y` ever stops working, run `docfx init --help` to check current
flags before guessing.)

- [ ] **Step 2: Edit docfx.json — disable the browser-dependent features, set app metadata**

The default scaffold's `metadata[0].src[0]` is `{"src": "../src", "files": ["**/*.csproj"]}`,
which — because `docfx/` sits directly next to `src/` in this repo — already picks up both
`DocToolkit.csproj` and `DocToolkit.Extensions.DependencyInjection.csproj` with no path changes
needed. Only `globalMetadata` needs editing. Replace `docfx/docfx.json`'s `globalMetadata` block
(leave everything else from the scaffold as-is):

```json
    "globalMetadata": {
      "_appName": "DocToolkit",
      "_appTitle": "DocToolkit",
      "_enableSearch": false,
      "pdf": false
    }
```

- [ ] **Step 3: Replace the landing page**

Replace `docfx/index.md` with:

```markdown
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
  classes, no DI container required.
- **[Ank.DocToolkit.Extensions.DependencyInjection](https://www.nuget.org/packages/Ank.DocToolkit.Extensions.DependencyInjection/)**
  — `services.AddDocToolkit()`, for ASP.NET Core / worker-service consumers.

See the [API reference](api/) for both packages, or the
[GitHub repository](https://github.com/Ank-KhoaHo/DocToolkit) for source, runnable samples
(`samples/`), and the full README.
```

- [ ] **Step 4: Trim the out-of-scope conceptual-docs section**

```bash
rm -rf docfx/docs
```

Replace `docfx/toc.yml` (removing the now-deleted "Docs" entry) with:

```yaml
- name: API
  href: api/
```

- [ ] **Step 5: Gitignore the generated output**

Add to `.gitignore` (check the existing file first — append rather than duplicate any existing
`_site` or `artifacts`-style entries):

```
docfx/_site/
docfx/api/
```

- [ ] **Step 6: Build locally and verify real API content was generated**

Run: `docfx docfx/docfx.json`
Expected: `Build succeeded.`, `0 warning(s)`, `0 error(s)`. This should complete in well under a
minute — if it hangs and you see "Downloading Chrome Headless Shell" in the output, Step 2's
`globalMetadata` edit didn't take effect; stop and fix it before continuing (don't wait it out).

Run: `ls docfx/_site/api/*.html | wc -l`
Expected: at least 15 (one file per public type across both projects — verified at 18 while
writing this plan; a lower number, especially near 0, means the `metadata.src` glob isn't finding
the csproj files).

Run: `grep -o '<title>[^<]*</title>' docfx/_site/api/DocToolkit.HtmlToDocxConverter.html`
Expected: `<title>Class HtmlToDocxConverter </title>` — confirms real per-type content, not an
empty placeholder page.

- [ ] **Step 7: Clean up local build output before committing**

```bash
rm -rf docfx/_site docfx/api
```

(These are gitignored per Step 5, but `docfx.json`'s companion `obj/` or log files from the local
build shouldn't be staged either — run `git status` and confirm only the files this task
intentionally created/edited are shown.)

- [ ] **Step 8: Commit**

```bash
git add docfx .gitignore
git commit -m "docs: scaffold DocFX API-reference site for both packages"
```

---

### Task 4: Publish pipeline + doc updates

**Files:**
- Create: `.github/workflows/docs.yml`
- Modify: `README.md` (add a "Documentation" link)
- Modify: `CLAUDE.md` (document the `samples/` convention and the docs pipeline)

**Interfaces:**
- Consumes: `docfx/docfx.json` from Task 3; the `Release` workflow's name (`.github/workflows/release.yml`'s `name: Release` — confirm this still reads `Release` before wiring the `workflow_run` trigger, since that trigger matches on the *workflow's declared name*, not its filename).

- [ ] **Step 1: Confirm the Release workflow's name**

Run: `grep '^name:' .github/workflows/release.yml`
Expected: `name: Release`. This exact string is what Step 2's `workflow_run.workflows` list must
match — if it's ever renamed, `docs.yml` silently stops triggering (no error, it just never
fires), so treat this as a real dependency, not a coincidence.

- [ ] **Step 2: Create the docs workflow**

Create `.github/workflows/docs.yml`:

```yaml
name: Docs

# Publishes the DocFX API-reference site to GitHub Pages. Triggered by the Release
# workflow completing SUCCESSFULLY - not independently on the same tag push - so a
# version that fails release.yml's guards (missing CHANGELOG entry, banned dependency,
# broken tests) never gets a docs site describing it as shipped.
#
# Requires GitHub Pages enabled once, manually: Settings > Pages > Source > GitHub Actions.
#
# Known gotcha: workflow_run fires based on this file's state on the default branch at
# the time Release completes. The first release after this file merges to main is the
# first one that can trigger it - it cannot retroactively fire for an earlier release.

on:
  workflow_run:
    workflows: ["Release"]
    types: [completed]

permissions:
  contents: read
  pages: write
  id-token: write

concurrency:
  group: pages
  cancel-in-progress: false

jobs:
  build-and-deploy:
    if: github.event.workflow_run.conclusion == 'success'
    runs-on: ubuntu-latest
    environment:
      name: github-pages
      url: ${{ steps.deployment.outputs.page_url }}
    steps:
      - uses: actions/checkout@v4
        with:
          ref: ${{ github.event.workflow_run.head_sha }}

      - uses: actions/setup-dotnet@v4
        with:
          dotnet-version: |
            8.0.x
            10.0.x

      - name: Restore
        run: dotnet restore DocToolkit.sln

      - name: Install DocFX
        run: dotnet tool update -g docfx

      - name: Build docs site
        run: docfx docfx/docfx.json

      - uses: actions/configure-pages@v5

      - uses: actions/upload-pages-artifact@v3
        with:
          path: docfx/_site

      - name: Deploy to GitHub Pages
        id: deployment
        uses: actions/deploy-pages@v4
```

- [ ] **Step 3: Validate the workflow YAML**

Run: `python3 -c "import yaml; yaml.safe_load(open('.github/workflows/docs.yml'))" && echo "valid"`
Expected: `valid`

- [ ] **Step 4: Add a "Documentation" link to the README**

In `README.md`, near the top (next to the existing badges/intro), add:

```markdown
📖 [API documentation](https://ank-khoaho.github.io/DocToolkit/)
```

Read the current top of `README.md` first to place this naturally alongside the existing badge
line and intro paragraph — don't duplicate an existing "Documentation" mention if one already
exists from earlier work in this repo.

- [ ] **Step 5: Document the samples convention and docs pipeline in CLAUDE.md**

Read `CLAUDE.md`'s existing "Conventions" and "Releasing" sections first (they already document
the `-warnaserror` build gate and the tag-driven release process this task's new files plug into),
then add a new section, placed after "## The DI extensions package" and before "## Commands":

```markdown
## Samples and docs site

`samples/ConsoleSample` and `samples/MinimalApiSample` are runnable, added to `DocToolkit.sln`,
and reference the published packages via `PackageReference` (never `ProjectReference`) — same
reasoning as the extensions package itself: they prove the real published artifact works, not
whatever is currently on `main`. They're built by the existing CI `dotnet build` step with no
special handling; a breaking API change fails the next sample build.

`docfx/` holds a DocFX-generated API-reference site — separate from `docs/`, which holds this
project's planning/spec history, not site source. `.github/workflows/docs.yml` builds and deploys
it to GitHub Pages, triggered by `workflow_run` on `release.yml` completing **successfully** — not
independently on the same tag push, so a release that fails its guards never gets a docs site
describing it as shipped. Don't "simplify" this into a direct tag-push trigger; that would break
the guarantee.

**`docfx.json`'s `globalMetadata` must keep `_enableSearch: false` and `pdf: false`.** Without
them, DocFX's default template downloads a ~109 MB headless-Chromium binary via Playwright/Node.js
during the build — verified while adding this pipeline. Re-enabling either without re-verifying
the browser download doesn't come back will make `docs.yml` slow and pull in exactly the kind of
heavyweight native dependency this repo's premise guards otherwise keep out.
```

- [ ] **Step 6: Commit**

```bash
git add .github/workflows/docs.yml README.md CLAUDE.md
git commit -m "ci: publish the DocFX site to GitHub Pages on successful release"
```

---

## Manual step required (not automatable by any task above)

Before the next real `v*` tag is pushed, enable GitHub Pages on the repo: **Settings → Pages →
Source → GitHub Actions.** Until this is done, `docs.yml`'s `actions/deploy-pages` step will fail
with a clear error rather than silently doing nothing — that failure is expected and diagnosable,
not a sign the workflow itself is broken (same shape as the earlier nuget Trusted Publishing
policy and `CODECOV_TOKEN` setup steps in this project's history).
