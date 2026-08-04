# XLSX Bulk Reading Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development
> (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use
> checkbox (`- [ ]`) syntax for tracking.

**Goal:** Let a caller who is handed an arbitrary `.xlsx` discover its sheets and read one whole
sheet, without knowing anything about the file in advance.

**Architecture:** Four new public methods on the existing `WorkbookEditor` static class, following
the house pattern exactly — a `byte[]` overload and a `Stream` overload per capability, both calling
one private `*Core` so they cannot drift. No new source files; no new dependencies. The extent of a
sheet comes from ClosedXML's `LastCellUsed()`, and the result is anchored at A1 and padded
rectangular.

**Tech Stack:** .NET 8 / .NET 10, ClosedXML 0.105.1, xUnit.

**Spec:** `docs/2026-08-04-xlsx-bulk-reading-design.md`

## Global Constraints

- **Target frameworks are `net8.0;net10.0`.** Every test runs once per framework, so *N* tests
  report *2N* results.
- **The build has 0 warnings and must stay there.** Verify with
  `dotnet build DocToolkit.sln -c Release -warnaserror --no-incremental`. **`--no-incremental` is
  mandatory** — MSBuild skips unchanged projects and a skipped project emits no diagnostics, so
  without it `-warnaserror` reports `0 Warning(s)` on a tree that has warnings.
- **Every public member needs XML doc comments** on all parameters and exceptions, or `CS1573`
  fires and the build fails.
- **No new package references.** ClosedXML is already referenced.
- **Commit messages follow Conventional Commits** (`type(scope)?: description`), scope `core`.
  **Never add a `Co-Authored-By` trailer.**
- **This work must branch from a `main` that already contains the public-API approval test**
  (`tests/DocToolkit.Tests/PublicApiApprovalTests.cs`, PR #66). Task 5 regenerates its approved
  files. If that file does not exist, stop and get #66 merged first.
- **`main` cannot be pushed directly.** Work on `feat/xlsx-bulk-reading` and open a PR.

---

### Task 1: `SheetNames` — list the sheets in a workbook

**Files:**
- Modify: `src/DocToolkit/WorkbookEditor.cs`
- Test: `tests/DocToolkit.Tests/WorkbookEditorTests.cs`

**Interfaces:**
- Consumes: existing private helpers `Open(byte[])` and `ValidateArguments` in `WorkbookEditor.cs`.
- Produces:
  - `public static IReadOnlyList<string> WorkbookEditor.SheetNames(byte[] xlsx)`
  - `public static Task<IReadOnlyList<string>> WorkbookEditor.SheetNamesAsync(Stream source, CancellationToken ct = default)`
  - `private static List<string> WorkbookEditor.SheetNamesCore(XLWorkbook workbook)`

- [ ] **Step 1: Write the failing tests**

Add to `tests/DocToolkit.Tests/WorkbookEditorTests.cs`, inside the existing
`WorkbookEditorTests` class:

```csharp
    /// <summary>
    /// A workbook with three sheets, the middle one hidden, built directly with ClosedXML because
    /// WorkbookEditor.Create only makes single-sheet workbooks.
    /// </summary>
    private static byte[] ThreeSheetWorkbook()
    {
        using var workbook = new XLWorkbook();
        workbook.Worksheets.Add("First").Cell("A1").Value = "a";
        var hidden = workbook.Worksheets.Add("Hidden");
        hidden.Cell("A1").Value = "b";
        hidden.Hide();
        workbook.Worksheets.Add("Third").Cell("A1").Value = "c";

        using var ms = new MemoryStream();
        workbook.SaveAs(ms);
        return ms.ToArray();
    }

    [Fact]
    public void SheetNames_ReturnsEverySheetInTabOrder_IncludingHiddenOnes()
    {
        Assert.Equal(
            new[] { "First", "Hidden", "Third" },
            WorkbookEditor.SheetNames(ThreeSheetWorkbook()));
    }

    [Fact]
    public void SheetNames_RejectsMissingContent()
    {
        Assert.Throws<ArgumentNullException>(() => WorkbookEditor.SheetNames(null!));
        Assert.Throws<ArgumentException>(() => WorkbookEditor.SheetNames(Array.Empty<byte>()));
    }

    [Fact]
    public void SheetNames_WrapsAFileThatIsNotAWorkbook()
    {
        Assert.Throws<DocumentConversionException>(
            () => WorkbookEditor.SheetNames(new byte[] { 1, 2, 3, 4 }));
    }

    [Fact]
    public async Task SheetNamesAsync_AgreesWithTheByteArrayOverload()
    {
        var xlsx = ThreeSheetWorkbook();
        using var source = new MemoryStream(xlsx);

        Assert.Equal(WorkbookEditor.SheetNames(xlsx), await WorkbookEditor.SheetNamesAsync(source));
    }
```

- [ ] **Step 2: Run the tests to verify they fail**

```bash
dotnet test DocToolkit.sln -c Release --filter "FullyQualifiedName~WorkbookEditorTests.SheetNames"
```

Expected: **build failure**, `CS0117: 'WorkbookEditor' does not contain a definition for 'SheetNames'`.
A compile error is the correct "fails first" signal here — the method does not exist yet.

- [ ] **Step 3: Implement**

In `src/DocToolkit/WorkbookEditor.cs`, insert immediately **before** the existing `ReadCell`
method (that is, after `CreateCore`):

```csharp
    /// <summary>
    /// Lists every sheet in the workbook, in tab order, including hidden sheets — hiding a sheet
    /// is a presentation choice, not a privacy boundary, and a caller who cannot see a hidden sheet
    /// listed has no way to discover it exists.
    /// </summary>
    /// <param name="xlsx">The workbook bytes.</param>
    /// <returns>The sheet names, in tab order.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="xlsx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="xlsx"/> is empty.</exception>
    /// <exception cref="DocumentConversionException">The workbook could not be opened.</exception>
    public static IReadOnlyList<string> SheetNames(byte[] xlsx)
    {
        ValidateWorkbook(xlsx);

        try
        {
            using var workbook = Open(xlsx);
            return SheetNamesCore(workbook);
        }
        catch (Exception ex) when (ex is not DocumentConversionException)
        {
            throw new DocumentConversionException("Failed to read XLSX.", ex);
        }
    }

    /// <summary>
    /// Reads a workbook from <paramref name="source"/> and lists every sheet in tab order,
    /// including hidden sheets. <paramref name="source"/> is <b>read</b> to its end and is neither
    /// disposed, closed nor sought.
    /// </summary>
    /// <param name="source">The stream the workbook is read from.</param>
    /// <param name="ct">Cancels the read.</param>
    /// <returns>The sheet names, in tab order.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="source"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="source"/> is not readable or held no bytes.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The workbook could not be opened.</exception>
    public static async Task<IReadOnlyList<string>> SheetNamesAsync(
        Stream source, CancellationToken ct = default)
    {
        StreamPipeline.RequireReadable(source, nameof(source));
        ct.ThrowIfCancellationRequested();

        using var xlsx = await StreamPipeline
            .DrainAsync(source, "Workbook content was empty.", nameof(source), "Failed to read XLSX.", ct)
            .ConfigureAwait(false);

        try
        {
            using var workbook = new XLWorkbook(xlsx);
            return SheetNamesCore(workbook);
        }
        catch (Exception ex) when (ex is not DocumentConversionException)
        {
            throw new DocumentConversionException("Failed to read XLSX.", ex);
        }
    }

    // Ordered by Position explicitly: Position is the tab order, whereas the enumeration order of
    // Worksheets is not documented to be. Sorting makes the guarantee true by construction.
    private static List<string> SheetNamesCore(XLWorkbook workbook)
        => workbook.Worksheets.OrderBy(sheet => sheet.Position).Select(sheet => sheet.Name).ToList();
```

Then add this private helper next to the existing `ValidateArguments` (which stays as it is —
`ReadCell` and `SetCell` still use it):

```csharp
    private static void ValidateWorkbook(byte[] xlsx)
    {
        ArgumentNullException.ThrowIfNull(xlsx);
        if (xlsx.Length == 0)
            throw new ArgumentException("Workbook content was empty.", nameof(xlsx));
    }
```

- [ ] **Step 4: Run the tests to verify they pass**

```bash
dotnet test DocToolkit.sln -c Release --filter "FullyQualifiedName~WorkbookEditorTests.SheetNames"
```

Expected: PASS, 8 results (4 tests × 2 target frameworks).

- [ ] **Step 5: Commit**

```bash
git add src/DocToolkit/WorkbookEditor.cs tests/DocToolkit.Tests/WorkbookEditorTests.cs
git commit -m "feat(core): list a workbook's sheets with SheetNames"
```

---

### Task 2: `ReadSheet` — read one whole sheet

**Files:**
- Modify: `src/DocToolkit/WorkbookEditor.cs`
- Test: `tests/DocToolkit.Tests/WorkbookEditorTests.cs`

**Interfaces:**
- Consumes: `ValidateWorkbook(byte[])` from Task 1; existing `Open(byte[])` and
  `Sheet(XLWorkbook, string)`.
- Produces:
  - `public static IReadOnlyList<IReadOnlyList<string>> WorkbookEditor.ReadSheet(byte[] xlsx, string sheetName)`
  - `public static Task<IReadOnlyList<IReadOnlyList<string>>> WorkbookEditor.ReadSheetAsync(Stream source, string sheetName, CancellationToken ct = default)`
  - `private static List<IReadOnlyList<string>> WorkbookEditor.ReadSheetCore(XLWorkbook workbook, string sheetName)`

- [ ] **Step 1: Write the failing tests**

Add to `tests/DocToolkit.Tests/WorkbookEditorTests.cs`:

```csharp
    /// <summary>A workbook whose data deliberately does not start at A1.</summary>
    private static byte[] OffsetWorkbook()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("Data");
        sheet.Cell("C3").Value = "Region";
        sheet.Cell("D3").Value = "Total";
        sheet.Cell("C4").Value = "North";
        sheet.Cell("D4").Value = 1200;

        using var ms = new MemoryStream();
        workbook.SaveAs(ms);
        return ms.ToArray();
    }

    [Fact]
    public void ReadSheet_ReturnsTheRowsItWasCreatedWith()
    {
        var rows = WorkbookEditor.ReadSheet(SampleWorkbook(), "Sales");

        Assert.Equal(3, rows.Count);
        Assert.Equal(new[] { "Region", "Total" }, rows[0]);
        Assert.Equal(new[] { "North", "1200" }, rows[1]);
        Assert.Equal(new[] { "South", "950" }, rows[2]);
    }

    /// <summary>
    /// The result is anchored at A1, not at the first used cell, so rows[r][c] means what a caller
    /// reading a spreadsheet would expect. Implementing this by iterating the used range instead
    /// puts "Region" at rows[0][0] and looks entirely plausible.
    /// </summary>
    [Fact]
    public void ReadSheet_AnchorsTheResultAtA1_WhenTheDataStartsElsewhere()
    {
        var rows = WorkbookEditor.ReadSheet(OffsetWorkbook(), "Data");

        Assert.Equal(4, rows.Count);
        Assert.Equal("Region", rows[2][2]);
        Assert.Equal("Total", rows[2][3]);
        Assert.Equal("North", rows[3][2]);
        Assert.Equal("1200", rows[3][3]);

        Assert.Equal(new[] { "", "", "", "" }, rows[0]);
        Assert.Equal(new[] { "", "", "", "" }, rows[1]);
    }

    [Fact]
    public void ReadSheet_PadsEveryRowToTheSameWidth()
    {
        var xlsx = WorkbookEditor.Create("S", new[]
        {
            new object?[] { "a", "b", "c" },
            new object?[] { "d" },
        });

        var rows = WorkbookEditor.ReadSheet(xlsx, "S");

        Assert.All(rows, row => Assert.Equal(3, row.Count));
        Assert.Equal(new[] { "d", "", "" }, rows[1]);
    }

    /// <summary>
    /// Blank rows inside the range are kept. Dropping them would shift every later index and
    /// silently break the positional guarantee the whole design rests on.
    /// </summary>
    [Fact]
    public void ReadSheet_KeepsBlankRowsInsideTheRange()
    {
        var xlsx = WorkbookEditor.Create("S", new[]
        {
            new object?[] { "a" },
            new object?[] { null },
            new object?[] { "c" },
        });

        var rows = WorkbookEditor.ReadSheet(xlsx, "S");

        Assert.Equal(3, rows.Count);
        Assert.Equal(new[] { "" }, rows[1]);
        Assert.Equal(new[] { "c" }, rows[2]);
    }

    [Fact]
    public void ReadSheet_ReturnsAnEmptyListForAnEmptySheet()
    {
        using var workbook = new XLWorkbook();
        workbook.Worksheets.Add("Blank");
        using var ms = new MemoryStream();
        workbook.SaveAs(ms);

        Assert.Empty(WorkbookEditor.ReadSheet(ms.ToArray(), "Blank"));
    }

    /// <summary>
    /// "Used" must mean "has a value", never "has formatting" — otherwise one bolded empty cell
    /// pads every row out to it for no reason a caller could see.
    /// </summary>
    [Fact]
    public void ReadSheet_IgnoresCellsThatAreOnlyFormatted()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("S");
        sheet.Cell("A1").Value = "a";
        sheet.Cell("Z1").Style.Font.Bold = true;      // formatting only, no value
        using var ms = new MemoryStream();
        workbook.SaveAs(ms);

        var rows = WorkbookEditor.ReadSheet(ms.ToArray(), "S");

        Assert.Single(rows);
        Assert.Equal(new[] { "a" }, rows[0]);
    }

    /// <summary>Nothing in this library evaluates formulas; a formula cell reads back its cached value.</summary>
    [Fact]
    public void ReadSheet_ReturnsTheCachedValueOfAFormulaCell()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("S");
        sheet.Cell("A1").Value = 2;
        sheet.Cell("B1").Value = 3;
        sheet.Cell("C1").FormulaA1 = "A1+B1";
        using var ms = new MemoryStream();
        workbook.SaveAs(ms);

        Assert.Equal("5", WorkbookEditor.ReadSheet(ms.ToArray(), "S")[0][2]);
    }

    [Fact]
    public void ReadSheet_ThrowsWhenTheSheetDoesNotExist()
    {
        Assert.Throws<DocumentConversionException>(
            () => WorkbookEditor.ReadSheet(SampleWorkbook(), "Nope"));
    }

    [Fact]
    public void ReadSheet_RejectsMissingArguments()
    {
        Assert.Throws<ArgumentNullException>(() => WorkbookEditor.ReadSheet(null!, "Sales"));
        Assert.Throws<ArgumentNullException>(() => WorkbookEditor.ReadSheet(SampleWorkbook(), null!));
        Assert.Throws<ArgumentException>(() => WorkbookEditor.ReadSheet(Array.Empty<byte>(), "Sales"));
        Assert.Throws<ArgumentException>(() => WorkbookEditor.ReadSheet(SampleWorkbook(), " "));
    }

    [Fact]
    public async Task ReadSheetAsync_AgreesWithTheByteArrayOverload()
    {
        var xlsx = OffsetWorkbook();
        using var source = new MemoryStream(xlsx);

        var expected = WorkbookEditor.ReadSheet(xlsx, "Data");
        var actual = await WorkbookEditor.ReadSheetAsync(source, "Data");

        Assert.Equal(expected, actual);
    }
```

- [ ] **Step 2: Run the tests to verify they fail**

```bash
dotnet test DocToolkit.sln -c Release --filter "FullyQualifiedName~WorkbookEditorTests.ReadSheet"
```

Expected: **build failure**, `CS0117: 'WorkbookEditor' does not contain a definition for 'ReadSheet'`.

- [ ] **Step 3: Implement**

In `src/DocToolkit/WorkbookEditor.cs`, insert immediately after `SheetNamesCore`:

```csharp
    /// <summary>
    /// Reads a whole sheet as strings, anchored at A1: if the data starts at C3, its first value
    /// is at <c>rows[2][2]</c>. Every row is padded to the last used column, so all rows have the
    /// same length; blank cells — and entirely blank rows inside the range — come back as empty
    /// strings rather than being dropped, which keeps <c>rows[r][c]</c> positionally meaningful.
    ///
    /// Values are produced exactly as <see cref="ReadCell"/> produces them, so the two can never
    /// disagree about what a cell says. A formula cell yields its cached value: nothing in this
    /// library evaluates formulas.
    /// </summary>
    /// <param name="xlsx">The workbook bytes.</param>
    /// <param name="sheetName">The sheet to read.</param>
    /// <returns>
    /// The sheet's used range, anchored at A1 and padded rectangular; empty if the sheet holds no
    /// values.
    /// </returns>
    /// <exception cref="ArgumentNullException">Any argument is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="xlsx"/> is empty, or <paramref name="sheetName"/> is blank.
    /// </exception>
    /// <exception cref="DocumentConversionException">
    /// The workbook could not be opened, or the sheet does not exist.
    /// </exception>
    public static IReadOnlyList<IReadOnlyList<string>> ReadSheet(byte[] xlsx, string sheetName)
    {
        ValidateWorkbook(xlsx);
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);

        try
        {
            using var workbook = Open(xlsx);
            return ReadSheetCore(workbook, sheetName);
        }
        catch (Exception ex) when (ex is not DocumentConversionException)
        {
            throw new DocumentConversionException("Failed to read XLSX.", ex);
        }
    }

    /// <summary>
    /// Reads a workbook from <paramref name="source"/> and returns a whole sheet as strings. See
    /// <see cref="ReadSheet"/> for the anchoring, padding and formula rules — this overload applies
    /// the identical logic. <paramref name="source"/> is <b>read</b> to its end and is neither
    /// disposed, closed nor sought.
    /// </summary>
    /// <param name="source">The stream the workbook is read from.</param>
    /// <param name="sheetName">The sheet to read.</param>
    /// <param name="ct">Cancels the read.</param>
    /// <returns>
    /// The sheet's used range, anchored at A1 and padded rectangular; empty if the sheet holds no
    /// values.
    /// </returns>
    /// <exception cref="ArgumentNullException"><paramref name="source"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="source"/> is not readable or held no bytes, or <paramref name="sheetName"/>
    /// is blank.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">
    /// The workbook could not be opened, or the sheet does not exist.
    /// </exception>
    public static async Task<IReadOnlyList<IReadOnlyList<string>>> ReadSheetAsync(
        Stream source, string sheetName, CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        StreamPipeline.RequireReadable(source, nameof(source));
        ct.ThrowIfCancellationRequested();

        using var xlsx = await StreamPipeline
            .DrainAsync(source, "Workbook content was empty.", nameof(source), "Failed to read XLSX.", ct)
            .ConfigureAwait(false);

        try
        {
            using var workbook = new XLWorkbook(xlsx);
            return ReadSheetCore(workbook, sheetName);
        }
        catch (Exception ex) when (ex is not DocumentConversionException)
        {
            throw new DocumentConversionException("Failed to read XLSX.", ex);
        }
    }

    private static List<IReadOnlyList<string>> ReadSheetCore(XLWorkbook workbook, string sheetName)
    {
        var sheet = Sheet(workbook, sheetName);

        // The extent comes from LastCellUsed() rather than LastRowUsed()/LastColumnUsed(): those
        // return range rows/columns whose RowNumber()/ColumnNumber() are documented as positions
        // *within the range*, which is an off-by-origin waiting to happen. A cell's Address is
        // absolute. LastCellUsed() also ignores formatting, so one bolded empty cell out at Z1
        // cannot pad every row out to it. Null means the sheet holds no values at all.
        var last = sheet.LastCellUsed();
        if (last is null)
            return new List<IReadOnlyList<string>>();

        var lastRow = last.Address.RowNumber;
        var lastColumn = last.Address.ColumnNumber;

        // From row 1 and column 1, not from the first used cell: the result is anchored at A1 so
        // rows[r][c] addresses the sheet the way the caller sees it in Excel.
        var rows = new List<IReadOnlyList<string>>(lastRow);
        for (var r = 1; r <= lastRow; r++)
        {
            var row = new string[lastColumn];
            for (var c = 1; c <= lastColumn; c++)
                row[c - 1] = sheet.Cell(r, c).GetString();
            rows.Add(row);
        }

        return rows;
    }
```

- [ ] **Step 4: Run the tests to verify they pass**

```bash
dotnet test DocToolkit.sln -c Release --filter "FullyQualifiedName~WorkbookEditorTests.ReadSheet"
```

Expected: PASS, 20 results (10 tests × 2 target frameworks).

- [ ] **Step 5: Commit**

```bash
git add src/DocToolkit/WorkbookEditor.cs tests/DocToolkit.Tests/WorkbookEditorTests.cs
git commit -m "feat(core): read a whole sheet with ReadSheet"
```

---

### Task 3: Register both async overloads with `StreamOverloadTests`

**Files:**
- Modify: `tests/DocToolkit.Tests/StreamOverloadTests.cs`

**Interfaces:**
- Consumes: `WorkbookEditor.SheetNamesAsync` and `WorkbookEditor.ReadSheetAsync` from Tasks 1–2.
- Produces: nothing new; extends existing test data.

`CLAUDE.md` is explicit that an overload missing from these lists is the only way to escape the
whole `Stream`-overload suite. Both new methods take a `Stream source` and no destination, so they
belong in `SourceReaderNames` only — exactly like `WorkbookEditor.ReadCellAsync`.

- [ ] **Step 1: Record the current result count**

```bash
dotnet test DocToolkit.sln -c Release --filter "FullyQualifiedName~StreamOverloadTests"
```

Write down the number of results. The next steps must increase it — assuming the registration took
is precisely the failure mode this task exists to prevent.

- [ ] **Step 2: Add both names to `SourceReaderNames`**

In `tests/DocToolkit.Tests/StreamOverloadTests.cs`, in the `SourceReaderNames` array, replace:

```csharp
        "WorkbookEditor.ReadCellAsync",
        "WorkbookEditor.SetCellAsync",
```

with:

```csharp
        "WorkbookEditor.ReadCellAsync",
        "WorkbookEditor.SheetNamesAsync",
        "WorkbookEditor.ReadSheetAsync",
        "WorkbookEditor.SetCellAsync",
```

- [ ] **Step 3: Add both to the `InvokeAsync` dispatch**

In the same file, in the `InvokeAsync` switch, replace:

```csharp
            "WorkbookEditor.ReadCellAsync" =>
                WorkbookEditor.ReadCellAsync(source!, "Sales", "A1", ct),
```

with:

```csharp
            "WorkbookEditor.ReadCellAsync" =>
                WorkbookEditor.ReadCellAsync(source!, "Sales", "A1", ct),
            "WorkbookEditor.SheetNamesAsync" =>
                WorkbookEditor.SheetNamesAsync(source!, ct),
            "WorkbookEditor.ReadSheetAsync" =>
                WorkbookEditor.ReadSheetAsync(source!, "Sales", ct),
```

No change is needed to `SourceBytesFor` or `NewSource`: both already route anything starting with
`"WorkbookEditor"` to the `Xlsx` fixture, which has a sheet named `Sales`.

- [ ] **Step 4: Run the suite and confirm the count rose**

```bash
dotnet test DocToolkit.sln -c Release --filter "FullyQualifiedName~StreamOverloadTests"
```

Expected: PASS, and **more results than Step 1 recorded**. If the count is unchanged, the names did
not reach the theory data — fix that before continuing, because every assertion below would then be
passing vacuously.

- [ ] **Step 5: Commit**

```bash
git add tests/DocToolkit.Tests/StreamOverloadTests.cs
git commit -m "test(core): cover the new XLSX readers in the Stream overload suite"
```

---

### Task 4: Extend the air-gap guard

**Files:**
- Modify: `tests/DocToolkit.Tests/AirGapGuardTests.cs`

**Interfaces:**
- Consumes: `WorkbookEditor.SheetNames` and `WorkbookEditor.ReadSheet` from Tasks 1–2.
- Produces: nothing new; extends an existing test.

The guard asserts **zero** socket connections across the whole public API. A new public method that
is never exercised there is a hole in that claim.

- [ ] **Step 1: Exercise both new methods inside the existing test**

In `tests/DocToolkit.Tests/AirGapGuardTests.cs`, in `WorkbookEditor_ContactsNothing`, replace:

```csharp
        try
        {
            WorkbookEditor.ReadCell(withLinks, "Sales", "A1");
            WorkbookEditor.ReadCell(withLinks, "Sales", "C1");   // the external-workbook formula
            var updated = WorkbookEditor.SetCell(withLinks, "Sales", "B1", 1500);
            Assert.Equal("1500", WorkbookEditor.ReadCell(updated, "Sales", "B1"));
        }
```

with:

```csharp
        try
        {
            WorkbookEditor.ReadCell(withLinks, "Sales", "A1");
            WorkbookEditor.ReadCell(withLinks, "Sales", "C1");   // the external-workbook formula
            var updated = WorkbookEditor.SetCell(withLinks, "Sales", "B1", 1500);
            Assert.Equal("1500", WorkbookEditor.ReadCell(updated, "Sales", "B1"));

            // Bulk reads walk every cell, so they touch the external-workbook formula and the
            // hyperlink relationship together — the two things in a workbook that ask to be told
            // what is on another machine.
            WorkbookEditor.SheetNames(withLinks);
            WorkbookEditor.ReadSheet(withLinks, "Sales");
        }
```

- [ ] **Step 2: Update the assertion message to name what actually ran**

In the same method, replace:

```csharp
        await probe.AssertSilentAsync("WorkbookEditor.Create / ReadCell / SetCell");
```

with:

```csharp
        await probe.AssertSilentAsync(
            "WorkbookEditor.Create / ReadCell / SetCell / SheetNames / ReadSheet");
```

- [ ] **Step 3: Run the guard**

```bash
dotnet test DocToolkit.sln -c Release --filter "FullyQualifiedName~AirGapGuardTests"
```

Expected: PASS.

- [ ] **Step 4: Commit**

```bash
git add tests/DocToolkit.Tests/AirGapGuardTests.cs
git commit -m "test(core): prove the new XLSX readers open no sockets"
```

---

### Task 5: Re-approve the public API surface

**Files:**
- Modify: `tests/DocToolkit.Tests/PublicApi/DocToolkit.approved.txt`

**Interfaces:**
- Consumes: the four public methods from Tasks 1–2.
- Produces: nothing new.

Four public methods were added, so the approval test is red until the surface is re-approved. That
diff is the reviewable record of what the PR adds to the public API — the guard working, not
friction.

- [ ] **Step 1: Run the approval test and confirm it fails**

```bash
dotnet test DocToolkit.sln -c Release --filter "FullyQualifiedName~PublicApiApprovalTests"
```

Expected: FAIL, with a diff naming `SheetNames`, `SheetNamesAsync`, `ReadSheet` and `ReadSheetAsync`.
If it passes, the approval file is not being compared against the real surface — investigate before
going further.

- [ ] **Step 2: Copy the received file over the approved file**

The failing test writes a `*.received.txt` next to the approved file in the test output directory.
Copy it over the checked-in approved file:

```bash
cp tests/DocToolkit.Tests/bin/Release/net8.0/PublicApi/DocToolkit.received.txt \
   tests/DocToolkit.Tests/PublicApi/DocToolkit.approved.txt
```

- [ ] **Step 3: Read the diff before accepting it**

```bash
git diff tests/DocToolkit.Tests/PublicApi/DocToolkit.approved.txt
```

Expected: exactly four added lines, one per new method, with the signatures from the spec. **Anything
else in that diff is an unintended public API change** — stop and investigate rather than committing it.

- [ ] **Step 4: Re-run and confirm it passes**

```bash
dotnet test DocToolkit.sln -c Release --filter "FullyQualifiedName~PublicApiApprovalTests"
```

Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add tests/DocToolkit.Tests/PublicApi/DocToolkit.approved.txt
git commit -m "test(core): approve the new XLSX reading API"
```

---

### Task 6: Document the new surface and ship

**Files:**
- Modify: `src/DocToolkit/README.md`
- Modify: `README.md`
- Modify: `docs/2026-08-03-enhancement-backlog.md`
- Modify: `samples/ConsoleSample/Program.cs`

**Interfaces:**
- Consumes: everything from Tasks 1–5.
- Produces: nothing new.

- [ ] **Step 1: Add the methods to the package README's usage block**

In `src/DocToolkit/README.md`, in the `## Usage` fenced `csharp` block, replace:

```csharp
string cell = WorkbookEditor.ReadCell(xlsx, "Sales", "B2");
byte[] updated = WorkbookEditor.SetCell(xlsx, "Sales", "B2", 1500);
```

with:

```csharp
string cell = WorkbookEditor.ReadCell(xlsx, "Sales", "B2");
byte[] updated = WorkbookEditor.SetCell(xlsx, "Sales", "B2", 1500);

// Read a workbook you were handed, without knowing its shape in advance
IReadOnlyList<string> sheets = WorkbookEditor.SheetNames(xlsx);              // tab order, hidden included
IReadOnlyList<IReadOnlyList<string>> grid = WorkbookEditor.ReadSheet(xlsx, "Sales");
string topLeft = grid[0][0];    // anchored at A1, padded rectangular, blanks are ""
```

- [ ] **Step 2: Show a bulk read in the root README's usage block**

In `README.md`, in the fenced `csharp` usage block, replace:

```csharp
byte[] xlsx = WorkbookEditor.Create("Sales", new[] { new object?[] { "Region", "Total" } });
```

with:

```csharp
byte[] xlsx = WorkbookEditor.Create("Sales", new[] { new object?[] { "Region", "Total" } });

// Read one back without knowing its shape in advance
IReadOnlyList<string> sheets = WorkbookEditor.SheetNames(xlsx);
IReadOnlyList<IReadOnlyList<string>> grid = WorkbookEditor.ReadSheet(xlsx, sheets[0]);
```

- [ ] **Step 3: Demonstrate it in the console sample**

In `samples/ConsoleSample/Program.cs`, replace:

```csharp
string cellBefore = WorkbookEditor.ReadCell(xlsx, "Sales", "B2");
byte[] updated = WorkbookEditor.SetCell(xlsx, "Sales", "B2", 1500);
string cellAfter = WorkbookEditor.ReadCell(updated, "Sales", "B2");
Console.WriteLine($"   B2 before: {cellBefore}, after SetCell: {cellAfter}");
```

with:

```csharp
string cellBefore = WorkbookEditor.ReadCell(xlsx, "Sales", "B2");
byte[] updated = WorkbookEditor.SetCell(xlsx, "Sales", "B2", 1500);
string cellAfter = WorkbookEditor.ReadCell(updated, "Sales", "B2");
Console.WriteLine($"   B2 before: {cellBefore}, after SetCell: {cellAfter}");

// Bulk read: discover the sheets, then read one whole sheet without knowing its shape
IReadOnlyList<string> sheetNames = WorkbookEditor.SheetNames(updated);
IReadOnlyList<IReadOnlyList<string>> grid = WorkbookEditor.ReadSheet(updated, sheetNames[0]);
Console.WriteLine($"   sheets: {string.Join(", ", sheetNames)}; "
                  + $"{grid.Count} row(s) x {grid[0].Count} column(s)");
Console.WriteLine($"   row 2: {string.Join(" | ", grid[1])}");
```

**Note:** the sample uses top-level statements, so `sheetNames` and `grid` must not collide with a
name already declared earlier in the file. `slideText` is declared *below* this point in section 7,
so do not reuse that name.

- [ ] **Step 4: Mark the reading slice of A3 done in the backlog**

In `docs/2026-08-03-enhancement-backlog.md`, replace the **A3** table row:

```
| A3 | **XLSX surface is thin.** `Create` (single sheet), `ReadCell`, `SetCell` only. No multi-sheet, read-range / read-sheet, append-rows, list-sheet-names, formulas, or CSV import/export. | `src/DocToolkit/WorkbookEditor.cs` |
```

with:

```
| A3 | **XLSX surface is thin.** Reading is now covered — `SheetNames` and `ReadSheet` shipped 2026-08-04 (`docs/2026-08-04-xlsx-bulk-reading-design.md`). Still open, each its own decision: multi-sheet create, append-rows, formulas, CSV import/export. The last two may honestly be "no". | `src/DocToolkit/WorkbookEditor.cs` |
```

A3 stays open — one of its four parts shipped, three did not.

- [ ] **Step 5: Update the test counts the new tests invalidate**

Adding roughly fifteen tests makes every hard-coded count stale. Get the real number first:

```bash
dotnet test DocToolkit.sln -c Release 2>&1 | tail -5
```

Then update, using the actual figures rather than arithmetic on the old ones:

- `README.md`: the `dotnet test` comment (`# 288 tests x 2 target frameworks = 576 results`) and
  the `tests/` line in the *Repository layout* block.
- `CLAUDE.md`: `288 tests (246 core + 42 extensions) → 576 results` under *Conventions*, and the
  `# 288 tests x 2 TFMs = 576 results` comment under *Commands*.

**Leave the `182 tests` figure in `CLAUDE.md`'s Layout block alone** unless you are fixing it
deliberately — it already disagrees with the 246 stated two sections earlier, so it is pre-existing
drift, not something this change caused. If you do fix it, say so in the PR body rather than
folding it in silently.

- [ ] **Step 6: Run the full suite and the zero-warning build**

```bash
dotnet build DocToolkit.sln -c Release -warnaserror --no-incremental
dotnet test  DocToolkit.sln -c Release
```

Expected: `0 Warning(s), 0 Error(s)`, and every test passing on both target frameworks. The sample
is built by the solution build, so a broken sample fails here.

- [ ] **Step 7: Commit and open the pull request**

```bash
git add README.md CLAUDE.md src/DocToolkit/README.md samples/ConsoleSample/Program.cs \
        docs/2026-08-03-enhancement-backlog.md
git commit -m "docs(core): document XLSX bulk reading"
git push -u origin feat/xlsx-bulk-reading
gh pr create --base main --title "feat(core): read XLSX sheets in bulk"
```

Write the PR body to say what the four methods do, that values are strings matching `ReadCell`, that
results are anchored at A1, and that only the reading slice of A3 is covered.

---

## Notes for the implementer

**Do not `dotnet add package` anything.** ClosedXML is already referenced, and the four premise
guards in CI fail the build for a dependency that breaks licensing, adds native binaries, breaks
Linux, or opens a socket.

**`ValidateArguments` stays.** `ReadCell` and `SetCell` still use it; Task 1 adds `ValidateWorkbook`
alongside it for the two new methods that take no `cellRef`. Do not merge them.

**If a test passes the first time you run it, be suspicious.** Every test in this plan has a step
that runs it *before* the implementation exists, precisely so a vacuous pass is visible.

### ClosedXML behaviour these tests assume — measured, not inferred

Run against ClosedXML 0.105.1 on 2026-08-04 with a throwaway probe test, because the expected values
below are what several assertions in this plan hinge on, and a plan built on guesses about a
third-party library produces tests that get "fixed" into meaninglessness when they fail:

| Probe | Result |
|---|---|
| Three sheets, middle one `Hide()`n, ordered by `Position` | `First:Visible, Hidden:Hidden, Third:Visible` — hidden sheets survive and tab order holds |
| `LastCellUsed()` where data occupies C3 and D4 | `D4`, `RowNumber=4`, `ColumnNumber=4` — **absolute**, not range-relative |
| `LastCellUsed()` on a sheet with no values | `null` |
| `A1` has a value, `Z1` is bold with no value | `LastCellUsed() == A1` — formatting does **not** widen the range |
| `C1.FormulaA1 = "A1+B1"` over `2` and `3`, saved and reopened | `GetString() == "5"`, `DataType == Number` — the cached value is written |
| `Create` with rows `["a"], [null], ["c"]` | `LastCellUsed() == A3` — the blank middle row is inside the range |

If any of these changes under a future ClosedXML bump, the corresponding test fails loudly, which is
the point of having them.
