# DI Extensions Stream/Async Parity Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add the `Stream`-based async overloads that already exist on every `Ank.DocToolkit` core static class to the six matching `Ank.DocToolkit.Extensions.DependencyInjection` interfaces, so DI consumers can read from and write to a request/response body or file without buffering a whole document into a `byte[]`.

**Architecture:** No new types, no new projects. Each of the six existing interfaces (`IDocxEditor`, `IPresentationEditor`, `IWorkbookEditor`, `IHtmlToDocxConverter`, `IHtmlToPdfConverter`, `IDocxToPdfConverter`) gains the `Stream`-based async members its core static counterpart already has; each matching `*Service` class gains a one-line delegating implementation, identical in spirit to every method already there.

**Tech Stack:** .NET 8 / .NET 10 · xUnit · the `Ank.DocToolkit` package at version floor `[0.2.0, )` — the published release that added the `Stream` overloads being wrapped (see Global Constraints)

**Source design:** [`docs/2026-08-03-di-extensions-parity-design.md`](2026-08-03-di-extensions-parity-design.md) — read it for the *why*; this plan is the *how*.

## Global Constraints

- **Additive only.** No existing public member's name, signature, or behavior changes.
- **No new NuGet package dependencies.** Every new method delegates to a core `DocToolkit` static method that already exists in a published `Ank.DocToolkit` release.
- **`src/DocToolkit.Extensions.DependencyInjection/DocToolkit.Extensions.DependencyInjection.csproj` references `Ank.DocToolkit` at `Version="[0.2.0, )"` (already fixed — do not lower it, and never switch it to a `ProjectReference`).** The Stream overloads this plan wraps first shipped in the published `0.2.0` release (see `CHANGELOG.md`); `Ank.DocToolkit 0.1.0` predates them. Because NuGet resolves a minimum-version range like `[X, )` to the *lowest* satisfying version, not the latest, the floor must name the version that actually introduced the API being wrapped, or restore silently picks an older version that won't compile. A `ProjectReference` would compile locally but ship a package with no declared `Ank.DocToolkit` dependency — never use one to work around a version-floor problem.
- **Fully qualify `DocToolkit` namespace types inside service implementation files** (e.g. `DocToolkit.DocxEditor.ExtractTextAsync(...)`), exactly as every existing method in those files already does — these files live in `namespace DocToolkit.Extensions.DependencyInjection`, which also declares the `I*` interface of the same short name, so an unqualified reference is needlessly easy to misread. Test files may use plain `using DocToolkit;`, matching the existing test files that already do.
- **Every new public interface member needs an XML doc comment.** `GenerateDocumentationFile` is `true` and CI builds with `-warnaserror`; a missing doc comment is CS1591, which fails the build.
- **No file-path convenience methods** (`ConvertToFileAsync`, `ConvertFile`) and **no per-call `allowRemoteImageDownload` override** — out of scope per the design's non-goals. The two HTML-converter `Stream` overloads thread `_options.AllowRemoteImageDownload`, exactly like the existing `byte[]` overloads already do.
- **Never assert byte equality between two separately *generated* packages — but do assert it for *edits*.** Building a `.docx`/`.xlsx` from scratch stamps it with fresh metadata, so two calls on identical input produce different bytes; that is true even of two calls to the same core static method, so it is never evidence of a wrapper defect. Editing an existing package preserves it and is byte-reproducible. Verified empirically for each: **non-deterministic** — `WorkbookEditor.Create`, `HtmlToDocxConverter.ConvertAsync`; **unconditionally deterministic** (OpenXML-SDK edit paths) — `DocxEditor.ReplaceTextAsync`, `PresentationEditor.ReplaceTextAsync`, plus `HtmlToPdfConverter`/`DocxToPdfConverter` output; **deterministic only between adjacent calls** — `WorkbookEditor.SetCell`/`SetCellAsync`, because ClosedXML writes through `ZipArchive`, which stamps each entry with a wall-clock last-modified timestamp at 2-second granularity (two calls 2.5s apart differ by 2 bytes; back-to-back calls are identical). For a generated package assert parity on readable content (`ExtractText`, `ReadCell`) plus the format's magic number; for an edit, assert byte equality against the corresponding static method — and for `WorkbookEditor` specifically, keep the two calls adjacent, with no I/O or delay between them.
- **Commit messages must not contain a `Co-Authored-By` trailer.**
- **Build must stay at 0 warnings** under `dotnet build DocToolkit.sln -c Release -warnaserror` (the same command CI runs).
- **Target frameworks stay `net8.0;net10.0`** for both the library and its test project — no `.csproj` changes needed in this plan.

---

## File Structure

No new files. Every task modifies one interface, its matching service, and its matching test file (plus README in the final task):

```
DocToolkit/
├── src/DocToolkit.Extensions.DependencyInjection/
│   ├── IDocxEditor.cs                        (Task 1)
│   ├── DocxEditorService.cs                  (Task 1)
│   ├── IPresentationEditor.cs                (Task 2)
│   ├── PresentationEditorService.cs          (Task 2)
│   ├── IWorkbookEditor.cs                    (Task 3)
│   ├── WorkbookEditorService.cs              (Task 3)
│   ├── IHtmlToDocxConverter.cs               (Task 4)
│   ├── HtmlToDocxConverterService.cs         (Task 4)
│   ├── IHtmlToPdfConverter.cs                (Task 5)
│   ├── HtmlToPdfConverterService.cs          (Task 5)
│   ├── IDocxToPdfConverter.cs                (Task 6)
│   ├── DocxToPdfConverterService.cs          (Task 6)
│   └── README.md                             (Task 7)
└── tests/DocToolkit.Extensions.DependencyInjection.Tests/
    ├── DocxEditorServiceTests.cs             (Task 1)
    ├── PresentationEditorServiceTests.cs     (Task 2)
    ├── WorkbookEditorServiceTests.cs         (Task 3)
    ├── HtmlToDocxConverterServiceTests.cs    (Task 4)
    ├── HtmlToPdfConverterServiceTests.cs     (Task 5)
    └── DocxToPdfConverterServiceTests.cs     (Task 6)
```

Tasks 1–6 are independent of each other (each touches a different interface/service pair and consumes only the already-published core static API) and may be done in any order. Task 7 (README + final verification) depends on all of them being done.

---

### Task 1: `IDocxEditor` / `DocxEditorService` — Stream/async parity

**Files:**
- Modify: `src/DocToolkit.Extensions.DependencyInjection/IDocxEditor.cs`
- Modify: `src/DocToolkit.Extensions.DependencyInjection/DocxEditorService.cs`
- Test: `tests/DocToolkit.Extensions.DependencyInjection.Tests/DocxEditorServiceTests.cs`

**Interfaces:**
- Consumes: `DocToolkit.DocxEditor.ReplaceTextAsync(Stream, IReadOnlyDictionary<string,string>, Stream, CancellationToken)`, `DocToolkit.DocxEditor.ExtractTextAsync(Stream, CancellationToken)`, `DocToolkit.DocxEditor.ExtractTextAsync(Stream, bool, CancellationToken)` — all already exist in `Ank.DocToolkit`.
- Produces: `IDocxEditor.ReplaceTextAsync`, `IDocxEditor.ExtractTextAsync(Stream, CancellationToken)`, `IDocxEditor.ExtractTextAsync(Stream, bool, CancellationToken)`.

- [ ] **Step 1: Write the failing tests**

Add to `tests/DocToolkit.Extensions.DependencyInjection.Tests/DocxEditorServiceTests.cs`, inside the `DocxEditorServiceTests` class (after the existing `ReplaceText_RejectsNullReplacements` test, before the closing `}`):

```csharp
    [Fact]
    public async Task ExtractTextAsync_MatchesTheStaticMethod()
    {
        var docx = await HtmlToDocxConverter.ConvertAsync("<p>Body copy.</p>");
        var sut = new DocxEditorService();

        var expected = await DocxEditor.ExtractTextAsync(new MemoryStream(docx));
        var actual = await sut.ExtractTextAsync(new MemoryStream(docx));

        Assert.Equal(expected, actual);
    }

    [Fact]
    public async Task ExtractTextAsync_WithHeadersAndFooters_MatchesTheStaticMethod()
    {
        var docx = DocxWithHeaderAndFooter("Body text.", "Page header", "Page footer");
        var sut = new DocxEditorService();

        var expected = await DocxEditor.ExtractTextAsync(new MemoryStream(docx), includeHeadersAndFooters: true);
        var actual = await sut.ExtractTextAsync(new MemoryStream(docx), includeHeadersAndFooters: true);

        Assert.Equal(expected, actual);
        Assert.Contains("Page header", actual);
    }

    [Fact]
    public async Task ReplaceTextAsync_MatchesTheStaticMethod()
    {
        var docx = await HtmlToDocxConverter.ConvertAsync("<p>Dear {{name}}.</p>");
        var sut = new DocxEditorService();
        var replacements = new Dictionary<string, string> { ["{{name}}"] = "Contoso Ltd" };

        using var expected = new MemoryStream();
        await DocxEditor.ReplaceTextAsync(new MemoryStream(docx), replacements, expected);

        using var actual = new MemoryStream();
        await sut.ReplaceTextAsync(new MemoryStream(docx), replacements, actual);

        Assert.Equal(expected.ToArray(), actual.ToArray());
    }
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `dotnet test DocToolkit.sln --filter "FullyQualifiedName~DocxEditorServiceTests"`
Expected: build FAILS — `DocxEditorService` does not contain a definition for `ExtractTextAsync` (CS1061), same for `ReplaceTextAsync`.

- [ ] **Step 3: Add the interface members**

In `src/DocToolkit.Extensions.DependencyInjection/IDocxEditor.cs`, insert before the closing `}` (after the existing `ExtractText(byte[], bool)` member):

```csharp

    /// <summary>
    /// Reads a .docx from <paramref name="source"/>, replaces every key with its value, and writes
    /// the result to <paramref name="destination"/>. See <see cref="ReplaceText"/> for exactly what
    /// counts as a match. <paramref name="source"/> is <b>read</b> to its end and
    /// <paramref name="destination"/> is <b>written</b>; neither is disposed, closed or sought.
    /// </summary>
    /// <exception cref="ArgumentNullException">Any argument is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="source"/> is not readable or held no bytes, or <paramref name="destination"/>
    /// is not writable.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The package could not be opened or edited.</exception>
    Task ReplaceTextAsync(
        Stream source, IReadOnlyDictionary<string, string> replacements, Stream destination,
        CancellationToken ct = default);

    /// <summary>
    /// Reads a .docx from <paramref name="source"/> and returns the plain text of its body. Headers,
    /// footers, footnotes and endnotes are not included. <paramref name="source"/> is <b>read</b> to
    /// its end and is neither disposed, closed nor sought.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="source"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="source"/> is not readable or held no bytes.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The package could not be opened or read.</exception>
    Task<string> ExtractTextAsync(Stream source, CancellationToken ct = default);

    /// <summary>
    /// Reads a .docx from <paramref name="source"/> and returns its plain text. See
    /// <see cref="ExtractText(byte[], bool)"/> for what <paramref name="includeHeadersAndFooters"/>
    /// controls. <paramref name="source"/> is <b>read</b> to its end and is neither disposed, closed
    /// nor sought.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="source"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="source"/> is not readable or held no bytes.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The package could not be opened or read.</exception>
    Task<string> ExtractTextAsync(Stream source, bool includeHeadersAndFooters, CancellationToken ct = default);
```

- [ ] **Step 4: Add the service implementation**

In `src/DocToolkit.Extensions.DependencyInjection/DocxEditorService.cs`, insert before the closing `}` (after the existing `ExtractText(byte[], bool)` method):

```csharp

    public Task ReplaceTextAsync(
        Stream source, IReadOnlyDictionary<string, string> replacements, Stream destination,
        CancellationToken ct = default)
        => DocToolkit.DocxEditor.ReplaceTextAsync(source, replacements, destination, ct);

    public Task<string> ExtractTextAsync(Stream source, CancellationToken ct = default)
        => DocToolkit.DocxEditor.ExtractTextAsync(source, ct);

    public Task<string> ExtractTextAsync(Stream source, bool includeHeadersAndFooters, CancellationToken ct = default)
        => DocToolkit.DocxEditor.ExtractTextAsync(source, includeHeadersAndFooters, ct);
```

- [ ] **Step 5: Run tests to verify they pass**

Run: `dotnet test DocToolkit.sln --filter "FullyQualifiedName~DocxEditorServiceTests"`
Expected: PASS (all `DocxEditorServiceTests`, both `net8.0` and `net10.0`).

- [ ] **Step 6: Commit**

```bash
git add src/DocToolkit.Extensions.DependencyInjection/IDocxEditor.cs src/DocToolkit.Extensions.DependencyInjection/DocxEditorService.cs tests/DocToolkit.Extensions.DependencyInjection.Tests/DocxEditorServiceTests.cs
git commit -m "feat(di-extensions): add Stream/async overloads to IDocxEditor"
```

---

### Task 2: `IPresentationEditor` / `PresentationEditorService` — Stream/async parity

**Files:**
- Modify: `src/DocToolkit.Extensions.DependencyInjection/IPresentationEditor.cs`
- Modify: `src/DocToolkit.Extensions.DependencyInjection/PresentationEditorService.cs`
- Test: `tests/DocToolkit.Extensions.DependencyInjection.Tests/PresentationEditorServiceTests.cs`

**Interfaces:**
- Consumes: `DocToolkit.PresentationEditor.SlideCountAsync(Stream, CancellationToken)`, `DocToolkit.PresentationEditor.ExtractTextAsync(Stream, CancellationToken)`, `DocToolkit.PresentationEditor.ReplaceTextAsync(Stream, IReadOnlyDictionary<string,string>, Stream, CancellationToken)` — all already exist in `Ank.DocToolkit`.
- Produces: `IPresentationEditor.SlideCountAsync`, `IPresentationEditor.ExtractTextAsync`, `IPresentationEditor.ReplaceTextAsync`.

- [ ] **Step 1: Write the failing tests**

Add to `tests/DocToolkit.Extensions.DependencyInjection.Tests/PresentationEditorServiceTests.cs`, inside the `PresentationEditorServiceTests` class (after the existing `ReplaceText_SubstitutesPlaceholders` test, before the closing `}`):

```csharp
    [Fact]
    public async Task SlideCountAsync_ExtractTextAsync_MatchTheStaticMethods()
    {
        var pptx = SamplePptx();
        var sut = new PresentationEditorService();

        Assert.Equal(
            await PresentationEditor.SlideCountAsync(new MemoryStream(pptx)),
            await sut.SlideCountAsync(new MemoryStream(pptx)));

        Assert.Equal(
            await PresentationEditor.ExtractTextAsync(new MemoryStream(pptx)),
            await sut.ExtractTextAsync(new MemoryStream(pptx)));
    }

    [Fact]
    public async Task ReplaceTextAsync_MatchesTheStaticMethod()
    {
        var pptx = SamplePptx();
        var sut = new PresentationEditorService();
        var replacements = new Dictionary<string, string> { ["{{who}}"] = "World" };

        using var expected = new MemoryStream();
        await PresentationEditor.ReplaceTextAsync(new MemoryStream(pptx), replacements, expected);

        using var actual = new MemoryStream();
        await sut.ReplaceTextAsync(new MemoryStream(pptx), replacements, actual);

        Assert.Equal(expected.ToArray(), actual.ToArray());
    }
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `dotnet test DocToolkit.sln --filter "FullyQualifiedName~PresentationEditorServiceTests"`
Expected: build FAILS — `PresentationEditorService` does not contain a definition for `SlideCountAsync` (CS1061), same for `ExtractTextAsync`/`ReplaceTextAsync`.

- [ ] **Step 3: Add the interface members**

In `src/DocToolkit.Extensions.DependencyInjection/IPresentationEditor.cs`, insert before the closing `}` (after the existing `ReplaceText` member):

```csharp

    /// <summary>
    /// Reads a .pptx from <paramref name="source"/> and returns its slide count, counted from the
    /// deck's slide list. <paramref name="source"/> is <b>read</b> to its end and is neither
    /// disposed, closed nor sought.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="source"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="source"/> is not readable or held no bytes.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The package could not be opened or read.</exception>
    Task<int> SlideCountAsync(Stream source, CancellationToken ct = default);

    /// <summary>
    /// Reads a .pptx from <paramref name="source"/> and returns all text found on every slide, in
    /// deck order. See <see cref="ExtractText(byte[])"/> for exactly what counts as a text-bearing
    /// body. <paramref name="source"/> is <b>read</b> to its end and is neither disposed, closed nor
    /// sought.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="source"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="source"/> is not readable or held no bytes.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The package could not be opened or read.</exception>
    Task<IReadOnlyList<string>> ExtractTextAsync(Stream source, CancellationToken ct = default);

    /// <summary>
    /// Reads a .pptx from <paramref name="source"/>, replaces every key with its value across all
    /// slide text, and writes the result to <paramref name="destination"/>. See
    /// <see cref="ReplaceText"/> for exactly what counts as a match. <paramref name="source"/> is
    /// <b>read</b> to its end and <paramref name="destination"/> is <b>written</b>; neither is
    /// disposed, closed or sought.
    /// </summary>
    /// <exception cref="ArgumentNullException">Any argument is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="source"/> is not readable or held no bytes, or <paramref name="destination"/>
    /// is not writable.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The package could not be opened or edited.</exception>
    Task ReplaceTextAsync(
        Stream source, IReadOnlyDictionary<string, string> replacements, Stream destination,
        CancellationToken ct = default);
```

- [ ] **Step 4: Add the service implementation**

In `src/DocToolkit.Extensions.DependencyInjection/PresentationEditorService.cs`, insert before the closing `}` (after the existing `ReplaceText` method):

```csharp

    public Task<int> SlideCountAsync(Stream source, CancellationToken ct = default)
        => DocToolkit.PresentationEditor.SlideCountAsync(source, ct);

    public Task<IReadOnlyList<string>> ExtractTextAsync(Stream source, CancellationToken ct = default)
        => DocToolkit.PresentationEditor.ExtractTextAsync(source, ct);

    public Task ReplaceTextAsync(
        Stream source, IReadOnlyDictionary<string, string> replacements, Stream destination,
        CancellationToken ct = default)
        => DocToolkit.PresentationEditor.ReplaceTextAsync(source, replacements, destination, ct);
```

- [ ] **Step 5: Run tests to verify they pass**

Run: `dotnet test DocToolkit.sln --filter "FullyQualifiedName~PresentationEditorServiceTests"`
Expected: PASS (all `PresentationEditorServiceTests`, both `net8.0` and `net10.0`).

- [ ] **Step 6: Commit**

```bash
git add src/DocToolkit.Extensions.DependencyInjection/IPresentationEditor.cs src/DocToolkit.Extensions.DependencyInjection/PresentationEditorService.cs tests/DocToolkit.Extensions.DependencyInjection.Tests/PresentationEditorServiceTests.cs
git commit -m "feat(di-extensions): add Stream/async overloads to IPresentationEditor"
```

---

### Task 3: `IWorkbookEditor` / `WorkbookEditorService` — Stream/async parity

**Files:**
- Modify: `src/DocToolkit.Extensions.DependencyInjection/IWorkbookEditor.cs`
- Modify: `src/DocToolkit.Extensions.DependencyInjection/WorkbookEditorService.cs`
- Test: `tests/DocToolkit.Extensions.DependencyInjection.Tests/WorkbookEditorServiceTests.cs`

**Interfaces:**
- Consumes: `DocToolkit.WorkbookEditor.CreateAsync(string, IEnumerable<IEnumerable<object?>>, Stream, CancellationToken)`, `DocToolkit.WorkbookEditor.ReadCellAsync(Stream, string, string, CancellationToken)`, `DocToolkit.WorkbookEditor.SetCellAsync(Stream, string, string, object?, Stream, CancellationToken)` — all already exist in `Ank.DocToolkit`.
- Produces: `IWorkbookEditor.CreateAsync`, `IWorkbookEditor.ReadCellAsync`, `IWorkbookEditor.SetCellAsync`.

- [ ] **Step 1: Write the failing test**

`WorkbookEditorServiceTests.cs` currently has no `using DocToolkit;` — add it (the new test calls the core static `WorkbookEditor` directly for parity, same as `DocxEditorServiceTests.cs` and `PresentationEditorServiceTests.cs` already do). At the top of the file, change:

```csharp
using System.Linq;
using DocToolkit.Extensions.DependencyInjection;
using Xunit;
```

to:

```csharp
using System.Linq;
using DocToolkit;
using DocToolkit.Extensions.DependencyInjection;
using Xunit;
```

Then add to the `WorkbookEditorServiceTests` class (after the existing `Create_RejectsABlankSheetName` test, before the closing `}`):

```csharp
    [Fact]
    public async Task CreateAsync_ReadCellAsync_SetCellAsync_MatchTheStaticMethods()
    {
        var sut = new WorkbookEditorService();
        var rows = new object?[][]
        {
            new object?[] { "Region", "Total" },
            new object?[] { "North", 1200 },
        };

        using var created = new MemoryStream();
        await sut.CreateAsync("Sales", rows, created);
        var xlsx = created.ToArray();

        // Parity is asserted on readable content rather than on bytes: ClosedXML stamps every
        // package it builds with fresh metadata, so two Create calls on identical input never
        // produce identical bytes - not even two calls to the same static method.
        Assert.Equal(new byte[] { 0x50, 0x4B, 0x03, 0x04 }, xlsx.Take(4).ToArray());
        Assert.Equal("Region", WorkbookEditor.ReadCell(xlsx, "Sales", "A1"));

        var cell = await sut.ReadCellAsync(new MemoryStream(xlsx), "Sales", "B2");
        Assert.Equal(await WorkbookEditor.ReadCellAsync(new MemoryStream(xlsx), "Sales", "B2"), cell);
        Assert.Equal("1200", cell);

        // Editing an existing package is deterministic - only building one from scratch stamps
        // fresh metadata - so this half can hold the wrapper to byte-exact parity.
        using var updated = new MemoryStream();
        await sut.SetCellAsync(new MemoryStream(xlsx), "Sales", "B2", 1500, updated);

        using var expectedUpdated = new MemoryStream();
        await WorkbookEditor.SetCellAsync(new MemoryStream(xlsx), "Sales", "B2", 1500, expectedUpdated);

        Assert.Equal(expectedUpdated.ToArray(), updated.ToArray());
        Assert.Equal("1500", await sut.ReadCellAsync(new MemoryStream(updated.ToArray()), "Sales", "B2"));
    }
```

- [ ] **Step 2: Run test to verify it fails**

Run: `dotnet test DocToolkit.sln --filter "FullyQualifiedName~WorkbookEditorServiceTests"`
Expected: build FAILS — `WorkbookEditorService` does not contain a definition for `CreateAsync` (CS1061), same for `ReadCellAsync`/`SetCellAsync`.

- [ ] **Step 3: Add the interface members**

In `src/DocToolkit.Extensions.DependencyInjection/IWorkbookEditor.cs`, insert before the closing `}` (after the existing `SetCell` member):

```csharp

    /// <summary>
    /// Builds a workbook with one sheet populated from <paramref name="rows"/> and writes it to
    /// <paramref name="destination"/>. See <see cref="Create"/> for the exact typing and culture
    /// rules applied to each cell. <paramref name="destination"/> is <b>written</b> and is not
    /// disposed, closed or sought.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="rows"/> or <paramref name="destination"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="sheetName"/> is blank, a row is null, or <paramref name="destination"/> is
    /// not writable.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The workbook could not be built or written.</exception>
    Task CreateAsync(
        string sheetName, IEnumerable<IEnumerable<object?>> rows, Stream destination,
        CancellationToken ct = default);

    /// <summary>
    /// Reads a workbook from <paramref name="source"/> and returns a cell as a string.
    /// <paramref name="cellRef"/> is an A1-style reference. <paramref name="source"/> is
    /// <b>read</b> to its end and is neither disposed, closed nor sought.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="source"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="source"/> is not readable or held no bytes, or a name is blank.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">
    /// The workbook could not be opened, the sheet does not exist, or the reference is not valid.
    /// </exception>
    Task<string> ReadCellAsync(Stream source, string sheetName, string cellRef, CancellationToken ct = default);

    /// <summary>
    /// Reads a workbook from <paramref name="source"/>, sets one cell, and writes the result to
    /// <paramref name="destination"/>. <paramref name="cellRef"/> is an A1-style reference.
    /// <paramref name="source"/> is <b>read</b> to its end and <paramref name="destination"/> is
    /// <b>written</b>; neither is disposed, closed or sought.
    /// </summary>
    /// <exception cref="ArgumentNullException">
    /// <paramref name="source"/> or <paramref name="destination"/> is null.
    /// </exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="source"/> is not readable or held no bytes, a name is blank, or
    /// <paramref name="destination"/> is not writable.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">
    /// The workbook could not be opened, the sheet does not exist, or the reference is not valid.
    /// </exception>
    Task SetCellAsync(
        Stream source, string sheetName, string cellRef, object? value, Stream destination,
        CancellationToken ct = default);
```

- [ ] **Step 4: Add the service implementation**

In `src/DocToolkit.Extensions.DependencyInjection/WorkbookEditorService.cs`, insert before the closing `}` (after the existing `SetCell` method):

```csharp

    public Task CreateAsync(
        string sheetName, IEnumerable<IEnumerable<object?>> rows, Stream destination,
        CancellationToken ct = default)
        => DocToolkit.WorkbookEditor.CreateAsync(sheetName, rows, destination, ct);

    public Task<string> ReadCellAsync(Stream source, string sheetName, string cellRef, CancellationToken ct = default)
        => DocToolkit.WorkbookEditor.ReadCellAsync(source, sheetName, cellRef, ct);

    public Task SetCellAsync(
        Stream source, string sheetName, string cellRef, object? value, Stream destination,
        CancellationToken ct = default)
        => DocToolkit.WorkbookEditor.SetCellAsync(source, sheetName, cellRef, value, destination, ct);
```

- [ ] **Step 5: Run test to verify it passes**

Run: `dotnet test DocToolkit.sln --filter "FullyQualifiedName~WorkbookEditorServiceTests"`
Expected: PASS (all `WorkbookEditorServiceTests`, both `net8.0` and `net10.0`).

- [ ] **Step 6: Commit**

```bash
git add src/DocToolkit.Extensions.DependencyInjection/IWorkbookEditor.cs src/DocToolkit.Extensions.DependencyInjection/WorkbookEditorService.cs tests/DocToolkit.Extensions.DependencyInjection.Tests/WorkbookEditorServiceTests.cs
git commit -m "feat(di-extensions): add Stream/async overloads to IWorkbookEditor"
```

---

### Task 4: `IHtmlToDocxConverter` / `HtmlToDocxConverterService` — Stream overload

**Files:**
- Modify: `src/DocToolkit.Extensions.DependencyInjection/IHtmlToDocxConverter.cs`
- Modify: `src/DocToolkit.Extensions.DependencyInjection/HtmlToDocxConverterService.cs`
- Test: `tests/DocToolkit.Extensions.DependencyInjection.Tests/HtmlToDocxConverterServiceTests.cs`

**Interfaces:**
- Consumes: `DocToolkit.HtmlToDocxConverter.ConvertAsync(string, bool, Stream, CancellationToken)` — already exists in `Ank.DocToolkit`.
- Produces: `IHtmlToDocxConverter.ConvertAsync(string, Stream, CancellationToken)`.

- [ ] **Step 1: Write the failing test**

`HtmlToDocxConverterServiceTests.cs` currently has no `using System.IO;` — add it. At the top of the file, change:

```csharp
using System.Linq;
using DocToolkit;
using DocToolkit.Extensions.DependencyInjection;
using Microsoft.Extensions.Options;
using Xunit;
```

to:

```csharp
using System.IO;
using System.Linq;
using DocToolkit;
using DocToolkit.Extensions.DependencyInjection;
using Microsoft.Extensions.Options;
using Xunit;
```

Then add to the `HtmlToDocxConverterServiceTests` class (after the existing `ConvertAsync_RejectsNullHtml` test, before the closing `}`):

```csharp
    [Fact]
    public async Task ConvertAsync_ToStream_MatchesTheByteArrayOverload()
    {
        var sut = new HtmlToDocxConverterService(Options.Create(new DocToolkitOptions()));

        var expected = await sut.ConvertAsync("<h1>Title</h1><p>Body copy.</p>");

        using var destination = new MemoryStream();
        await sut.ConvertAsync("<h1>Title</h1><p>Body copy.</p>", destination);
        var actual = destination.ToArray();

        // Parity is asserted on readable content rather than on bytes: building a .docx stamps
        // the package with fresh metadata, so two conversions of identical markup never produce
        // identical bytes - not even two calls to the same static method.
        Assert.Equal(new byte[] { 0x50, 0x4B, 0x03, 0x04 }, actual.Take(4).ToArray());
        Assert.Equal(DocxEditor.ExtractText(expected), DocxEditor.ExtractText(actual));
        Assert.Contains("Body copy.", DocxEditor.ExtractText(actual));
    }
```

- [ ] **Step 2: Run test to verify it fails**

Run: `dotnet test DocToolkit.sln --filter "FullyQualifiedName~HtmlToDocxConverterServiceTests"`
Expected: build FAILS — no overload for method `ConvertAsync` takes 2 arguments (CS1501).

- [ ] **Step 3: Add the interface member**

In `src/DocToolkit.Extensions.DependencyInjection/IHtmlToDocxConverter.cs`, insert before the closing `}` (after the existing `ConvertAsync(string, CancellationToken)` member):

```csharp

    /// <summary>
    /// Converts <paramref name="html"/> and writes the .docx to <paramref name="destination"/>.
    /// <paramref name="destination"/> is <b>written</b>, from its current position, and is
    /// <b>not</b> disposed, closed or sought.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="html"/> or <paramref name="destination"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="destination"/> is not writable.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The HTML could not be converted or written.</exception>
    Task ConvertAsync(string html, Stream destination, CancellationToken ct = default);
```

- [ ] **Step 4: Add the service implementation**

In `src/DocToolkit.Extensions.DependencyInjection/HtmlToDocxConverterService.cs`, insert before the closing `}` (after the existing `ConvertAsync` method):

```csharp

    public Task ConvertAsync(string html, Stream destination, CancellationToken ct = default)
        => DocToolkit.HtmlToDocxConverter.ConvertAsync(html, _options.AllowRemoteImageDownload, destination, ct);
```

- [ ] **Step 5: Run test to verify it passes**

Run: `dotnet test DocToolkit.sln --filter "FullyQualifiedName~HtmlToDocxConverterServiceTests"`
Expected: PASS (all `HtmlToDocxConverterServiceTests`, both `net8.0` and `net10.0`).

- [ ] **Step 6: Commit**

```bash
git add src/DocToolkit.Extensions.DependencyInjection/IHtmlToDocxConverter.cs src/DocToolkit.Extensions.DependencyInjection/HtmlToDocxConverterService.cs tests/DocToolkit.Extensions.DependencyInjection.Tests/HtmlToDocxConverterServiceTests.cs
git commit -m "feat(di-extensions): add Stream overload to IHtmlToDocxConverter"
```

---

### Task 5: `IHtmlToPdfConverter` / `HtmlToPdfConverterService` — Stream overload

**Files:**
- Modify: `src/DocToolkit.Extensions.DependencyInjection/IHtmlToPdfConverter.cs`
- Modify: `src/DocToolkit.Extensions.DependencyInjection/HtmlToPdfConverterService.cs`
- Test: `tests/DocToolkit.Extensions.DependencyInjection.Tests/HtmlToPdfConverterServiceTests.cs`

**Interfaces:**
- Consumes: `DocToolkit.HtmlToPdfConverter.ConvertAsync(string, bool, Stream, CancellationToken)` — already exists in `Ank.DocToolkit`.
- Produces: `IHtmlToPdfConverter.ConvertAsync(string, Stream, CancellationToken)`.

- [ ] **Step 1: Write the failing test**

`HtmlToPdfConverterServiceTests.cs` currently has no `using System.IO;` — add it. At the top of the file, change:

```csharp
using System.Linq;
using DocToolkit.Extensions.DependencyInjection;
using Microsoft.Extensions.Options;
using Xunit;
```

to:

```csharp
using System.IO;
using System.Linq;
using DocToolkit.Extensions.DependencyInjection;
using Microsoft.Extensions.Options;
using Xunit;
```

Then add to the `HtmlToPdfConverterServiceTests` class (after the existing `ConvertAsync_RejectsNullHtml` test, before the closing `}`):

```csharp
    [Fact]
    public async Task ConvertAsync_ToStream_MatchesTheByteArrayOverload()
    {
        var sut = new HtmlToPdfConverterService(Options.Create(new DocToolkitOptions()));

        var expected = await sut.ConvertAsync("<h1>Invoice</h1><p>Total: 18,100.00</p>");

        using var destination = new MemoryStream();
        await sut.ConvertAsync("<h1>Invoice</h1><p>Total: 18,100.00</p>", destination);

        Assert.Equal(expected, destination.ToArray());
    }
```

- [ ] **Step 2: Run test to verify it fails**

Run: `dotnet test DocToolkit.sln --filter "FullyQualifiedName~HtmlToPdfConverterServiceTests"`
Expected: build FAILS — no overload for method `ConvertAsync` takes 2 arguments (CS1501).

- [ ] **Step 3: Add the interface member**

In `src/DocToolkit.Extensions.DependencyInjection/IHtmlToPdfConverter.cs`, insert before the closing `}` (after the existing `ConvertAsync(string, CancellationToken)` member):

```csharp

    /// <summary>
    /// Converts <paramref name="html"/> and writes the PDF to <paramref name="destination"/>.
    /// <paramref name="destination"/> is <b>written</b>, from its current position, and is
    /// <b>not</b> disposed, closed or sought.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="html"/> or <paramref name="destination"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="destination"/> is not writable.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The HTML could not be converted or written.</exception>
    Task ConvertAsync(string html, Stream destination, CancellationToken ct = default);
```

- [ ] **Step 4: Add the service implementation**

In `src/DocToolkit.Extensions.DependencyInjection/HtmlToPdfConverterService.cs`, insert before the closing `}` (after the existing `ConvertAsync` method):

```csharp

    public Task ConvertAsync(string html, Stream destination, CancellationToken ct = default)
        => DocToolkit.HtmlToPdfConverter.ConvertAsync(html, _options.AllowRemoteImageDownload, destination, ct);
```

- [ ] **Step 5: Run test to verify it passes**

Run: `dotnet test DocToolkit.sln --filter "FullyQualifiedName~HtmlToPdfConverterServiceTests"`
Expected: PASS (all `HtmlToPdfConverterServiceTests`, both `net8.0` and `net10.0`).

- [ ] **Step 6: Commit**

```bash
git add src/DocToolkit.Extensions.DependencyInjection/IHtmlToPdfConverter.cs src/DocToolkit.Extensions.DependencyInjection/HtmlToPdfConverterService.cs tests/DocToolkit.Extensions.DependencyInjection.Tests/HtmlToPdfConverterServiceTests.cs
git commit -m "feat(di-extensions): add Stream overload to IHtmlToPdfConverter"
```

---

### Task 6: `IDocxToPdfConverter` / `DocxToPdfConverterService` — Stream overload

**Files:**
- Modify: `src/DocToolkit.Extensions.DependencyInjection/IDocxToPdfConverter.cs`
- Modify: `src/DocToolkit.Extensions.DependencyInjection/DocxToPdfConverterService.cs`
- Test: `tests/DocToolkit.Extensions.DependencyInjection.Tests/DocxToPdfConverterServiceTests.cs`

**Interfaces:**
- Consumes: `DocToolkit.DocxToPdfConverter.ConvertAsync(Stream, Stream, CancellationToken)` — already exists in `Ank.DocToolkit`.
- Produces: `IDocxToPdfConverter.ConvertAsync(Stream, Stream, CancellationToken)`.

- [ ] **Step 1: Write the failing test**

`DocxToPdfConverterServiceTests.cs` currently has no `using System.IO;` — add it. At the top of the file, change:

```csharp
using System.Linq;
using DocToolkit;
using DocToolkit.Extensions.DependencyInjection;
using Xunit;
```

to:

```csharp
using System.IO;
using System.Linq;
using DocToolkit;
using DocToolkit.Extensions.DependencyInjection;
using Xunit;
```

Then add to the `DocxToPdfConverterServiceTests` class (after the existing `Convert_RejectsEmptyInput` test, before the closing `}`):

```csharp
    [Fact]
    public async Task ConvertAsync_Stream_MatchesTheByteArrayOverload()
    {
        var sut = new DocxToPdfConverterService();
        var docx = await HtmlToDocxConverter.ConvertAsync("<h1>Invoice</h1><p>Total: 18,100.00</p>");

        var expected = sut.Convert(docx);

        using var destination = new MemoryStream();
        await sut.ConvertAsync(new MemoryStream(docx), destination);

        Assert.Equal(expected, destination.ToArray());
    }
```

- [ ] **Step 2: Run test to verify it fails**

Run: `dotnet test DocToolkit.sln --filter "FullyQualifiedName~DocxToPdfConverterServiceTests"`
Expected: build FAILS — `DocxToPdfConverterService` does not contain a definition for `ConvertAsync` (CS1061).

- [ ] **Step 3: Add the interface member**

In `src/DocToolkit.Extensions.DependencyInjection/IDocxToPdfConverter.cs`, insert before the closing `}` (after the existing `Convert(byte[])` member):

```csharp

    /// <summary>
    /// Reads a .docx from <paramref name="source"/> and writes the rendered PDF to
    /// <paramref name="destination"/>. Neither stream is disposed, closed, sought or read back; the
    /// PDF is written straight through as the renderer produces it, so nothing here ever holds the
    /// whole rendered document in memory.
    /// </summary>
    /// <exception cref="ArgumentNullException">Either stream is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="source"/> is not readable or held no bytes, or <paramref name="destination"/>
    /// is not writable.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The DOCX could not be rendered.</exception>
    Task ConvertAsync(Stream source, Stream destination, CancellationToken ct = default);
```

- [ ] **Step 4: Add the service implementation**

In `src/DocToolkit.Extensions.DependencyInjection/DocxToPdfConverterService.cs`, insert before the closing `}` (after the existing `Convert` method):

```csharp

    public Task ConvertAsync(Stream source, Stream destination, CancellationToken ct = default)
        => DocToolkit.DocxToPdfConverter.ConvertAsync(source, destination, ct);
```

- [ ] **Step 5: Run test to verify it passes**

Run: `dotnet test DocToolkit.sln --filter "FullyQualifiedName~DocxToPdfConverterServiceTests"`
Expected: PASS (all `DocxToPdfConverterServiceTests`, both `net8.0` and `net10.0`).

- [ ] **Step 6: Commit**

```bash
git add src/DocToolkit.Extensions.DependencyInjection/IDocxToPdfConverter.cs src/DocToolkit.Extensions.DependencyInjection/DocxToPdfConverterService.cs tests/DocToolkit.Extensions.DependencyInjection.Tests/DocxToPdfConverterServiceTests.cs
git commit -m "feat(di-extensions): add Stream overload to IDocxToPdfConverter"
```

---

### Task 7: README update + full solution verification

**Files:**
- Modify: `src/DocToolkit.Extensions.DependencyInjection/README.md`

**Interfaces:**
- Consumes: nothing new — this task only documents the surface Tasks 1–6 already produced and verifies the whole solution.
- Produces: nothing new.

- [ ] **Step 1: Add a Stream-overload example to the README**

In `src/DocToolkit.Extensions.DependencyInjection/README.md`, insert a new code block directly after the existing usage example (the one ending with `public Task<byte[]> RenderAsync(string html) => _toPdf.ConvertAsync(html);` followed by its closing `}`):

```csharp
// Every interface also has a Stream overload, so a large document never has to be buffered
// into a byte[] — write straight to an HTTP response body instead:
app.MapPost("/invoices/pdf", async (string html, IHtmlToPdfConverter toPdf, HttpResponse response) =>
{
    response.ContentType = "application/pdf";
    await toPdf.ConvertAsync(html, response.Body);
});
```

Then update the paragraph below the usage examples — change:

```markdown
All six interfaces — `IHtmlToDocxConverter`, `IDocxToPdfConverter`, `IHtmlToPdfConverter`,
`IDocxEditor`, `IWorkbookEditor`, `IPresentationEditor` — mirror
[`Ank.DocToolkit`](https://www.nuget.org/packages/Ank.DocToolkit)'s static API one-for-one, are
registered as singletons (each wraps stateless logic), and are safe to inject and call
concurrently. See the core package's README for what each one does and the offline/licensing
guarantees behind them.
```

to:

```markdown
All six interfaces — `IHtmlToDocxConverter`, `IDocxToPdfConverter`, `IHtmlToPdfConverter`,
`IDocxEditor`, `IWorkbookEditor`, `IPresentationEditor` — mirror
[`Ank.DocToolkit`](https://www.nuget.org/packages/Ank.DocToolkit)'s static API one-for-one,
including its `Stream`-based async overloads, are registered as singletons (each wraps stateless
logic), and are safe to inject and call concurrently. See the core package's README for what each
one does and the offline/licensing guarantees behind them.
```

- [ ] **Step 2: Commit the README change**

```bash
git add src/DocToolkit.Extensions.DependencyInjection/README.md
git commit -m "docs(di-extensions): document the Stream overloads"
```

- [ ] **Step 3: Full solution build, 0 warnings**

Run: `dotnet build DocToolkit.sln -c Release -warnaserror`
Expected: `Build succeeded.` with `0 Warning(s)` and `0 Error(s)`.

- [ ] **Step 4: Full solution test run**

Run: `dotnet test DocToolkit.sln -c Release --no-build`
Expected: all tests pass, across both `net8.0` and `net10.0`, in both `DocToolkit.Tests` and `DocToolkit.Extensions.DependencyInjection.Tests`.

If Step 3 or Step 4 fails, fix the issue in the task that introduced it and re-run both before considering this plan complete — do not skip ahead.
