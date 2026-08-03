# Repeating table rows — implementation plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development
> (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use
> checkbox (`- [ ]`) syntax for tracking.

**Goal:** `DocxEditor.FillRows` expands a table row once per record, preserving the template row's
formatting, so a Word template can render invoice line items.

**Architecture:** Discovery finds template rows by walking each table's *direct child* rows — never
`Descendants`, which reaches into nested tables. Each record deep-clones the template row, and the
existing `RunTextSplicer` performs substitution on the clone, so formatting and hyperlink survival
come from already-proven code rather than a second implementation.

**Tech Stack:** `DocumentFormat.OpenXml`, xunit.

Design doc: `docs/2026-08-03-repeating-table-rows-design.md`. Read it first — the reasons for
composing with `ReplaceText` rather than overloading it, and for reusing the splicer, are there.

## Global Constraints

- **Branch from `main`, PR back into it.** `main` cannot be pushed directly.
- **Every merge to `main` publishes a release.** Keep the branch coherent; don't merge half a feature.
- **Conventional Commits** — `ci.yml`'s `commit-format` checks every commit in the PR.
- **Never add a `Co-Authored-By` trailer.**
- **Never rewrite `src/DocToolkit/DocToolkit.csproj` wholesale** — it carries the package metadata.
- **If you touch a dependency, regenerate the lockfile** with `dotnet restore <project>
  --force-evaluate` or the locked-mode guard fails. This task should need no new dependency.
- **Never use `Descendants<TableRow>()` or `Descendants<Paragraph>()` for discovery.** That is the
  documented cause of a silent, schema-valid data-loss bug in this codebase.
- Build runs at **0 warnings** under `-warnaserror`. Targets `net8.0;net10.0`; 224 tests × 2 = 448
  results today, and this plan adds to that number.
- **`RunTextSplicer` is `internal`** and the core test project already sees it. Do not make it public.

---

### Task 1: Find template rows without reaching into nested tables

**Files:**
- Create: `src/DocToolkit/TableRowFinder.cs`
- Test: `tests/DocToolkit.Tests/TableRowFinderTests.cs`
- Test helper: `tests/DocToolkit.Tests/DocxFixtures.cs` (extend)

**Interfaces:**
- Consumes: `DocumentFormat.OpenXml.Wordprocessing`.
- Produces: `internal static class TableRowFinder` with
  `IReadOnlyList<TableRow> Find(OpenXmlElement scope, string marker)`, returning innermost rows
  first. `marker` is the literal `"{{" + collection + "."`.

Discovery and detection **both** avoid `Descendants`. A row is a template row when one of its own
cells' **direct child paragraphs** contains the marker — a nested table's text must not make its
container look like a template row.

- [ ] **Step 1: Write the failing test**

Add to `tests/DocToolkit.Tests/DocxFixtures.cs`:

```csharp
    /// <summary>A .docx whose single table has a header row and one row per marker text supplied.</summary>
    public static byte[] TableDocx(params string[] rowCellTexts)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var body = doc.AddMainDocumentPart().Document = new Document(new Body());
            var table = new Table();
            foreach (var text in rowCellTexts)
            {
                table.Append(new TableRow(
                    new TableCell(new Paragraph(new Run(new Text(text))))));
            }
            doc.MainDocumentPart!.Document.Body!.Append(table);
            doc.MainDocumentPart.Document.Save();
        }
        return ms.ToArray();
    }
```

Create `tests/DocToolkit.Tests/TableRowFinderTests.cs`:

```csharp
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocToolkit.Tests;

public class TableRowFinderTests
{
    [Fact]
    public void Find_ReturnsOnlyRowsHoldingTheMarker()
    {
        var docx = DocxFixtures.TableDocx("Description", "{{item.Desc}}", "Total");

        using var ms = new MemoryStream(docx);
        using var doc = WordprocessingDocument.Open(ms, false);
        var found = TableRowFinder.Find(doc.MainDocumentPart!.Document.Body!, "{{item.");

        Assert.Single(found);
        Assert.Contains("{{item.Desc}}", found[0].InnerText);
    }
}
```

- [ ] **Step 2: Run it and watch it fail**

Run: `dotnet test tests/DocToolkit.Tests -c Release --filter FullyQualifiedName~TableRowFinderTests`

Expected: **compile failure** — `TableRowFinder` does not exist.

- [ ] **Step 3: Write the minimal implementation**

Create `src/DocToolkit/TableRowFinder.cs`:

```csharp
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocToolkit;

/// <summary>
/// Locates the table rows a repeating-row fill should expand.
///
/// Deliberately does NOT use <c>Descendants&lt;TableRow&gt;()</c>. That also yields the rows of
/// tables nested inside cells, which would both mis-identify a container row as a template row and
/// sweep an inner row into the outer row's expansion. The same shape of bug, via
/// <c>Descendants&lt;Paragraph&gt;()</c> reaching into text boxes, once caused schema-valid silent
/// data loss in this codebase - see CLAUDE.md.
///
/// Rows are returned innermost-first, so a nested template row is expanded before any row that
/// contains it.
/// </summary>
internal static class TableRowFinder
{
    public static IReadOnlyList<TableRow> Find(OpenXmlElement scope, string marker)
    {
        var found = new List<TableRow>();
        Collect(scope, marker, found);
        return found;
    }

    private static void Collect(OpenXmlElement scope, string marker, List<TableRow> found)
    {
        foreach (var table in scope.ChildElements.OfType<Table>())
        {
            foreach (var row in table.ChildElements.OfType<TableRow>().ToList())
            {
                // Innermost first: a nested template row is expanded before its container.
                foreach (var cell in row.ChildElements.OfType<TableCell>())
                    Collect(cell, marker, found);

                if (OwnsMarker(row, marker)) found.Add(row);
            }
        }
    }

    /// <summary>
    /// True when one of the row's own cells' direct child paragraphs holds the marker. Text inside
    /// a nested table belongs to that table's rows, not to this one.
    /// </summary>
    private static bool OwnsMarker(TableRow row, string marker)
    {
        foreach (var cell in row.ChildElements.OfType<TableCell>())
            foreach (var paragraph in cell.ChildElements.OfType<Paragraph>())
                if (paragraph.InnerText.Contains(marker, StringComparison.Ordinal))
                    return true;

        return false;
    }
}
```

- [ ] **Step 4: Run the test and watch it pass**

Run: `dotnet test tests/DocToolkit.Tests -c Release --filter FullyQualifiedName~TableRowFinderTests`
Expected: PASS, 2 results (one per TFM).

- [ ] **Step 5: Write the nested-table test — the trap this class exists for**

Append to `TableRowFinderTests.cs`:

```csharp
    [Fact]
    public void Find_DoesNotTreatAContainerRowAsATemplateRow()
    {
        // An outer table whose single row holds a nested table; only the INNER row has the marker.
        var inner = new Table(new TableRow(
            new TableCell(new Paragraph(new Run(new Text("{{item.Desc}}"))))));
        var outer = new Table(new TableRow(
            new TableCell(new Paragraph(new Run(new Text("no marker here"))), inner)));

        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            doc.AddMainDocumentPart().Document = new Document(new Body(outer));
            doc.MainDocumentPart!.Document.Save();
        }

        using var read = new MemoryStream(ms.ToArray());
        using var opened = WordprocessingDocument.Open(read, false);
        var found = TableRowFinder.Find(opened.MainDocumentPart!.Document.Body!, "{{item.");

        // Exactly the inner row. Descendants-based discovery would return both, and the outer row
        // would then be cloned per record with the inner table inside it.
        Assert.Single(found);
        Assert.DoesNotContain("no marker here", found[0].InnerText);
    }
```

- [ ] **Step 6: Run it**

Run: `dotnet test tests/DocToolkit.Tests -c Release --filter FullyQualifiedName~TableRowFinderTests`
Expected: PASS, 4 results.

**If this fails with two rows found**, discovery or detection is using `Descendants` somewhere.
Do not "fix" it by filtering the results afterwards — fix the walk.

- [ ] **Step 7: Commit**

```bash
git add src/DocToolkit/TableRowFinder.cs tests/DocToolkit.Tests/TableRowFinderTests.cs tests/DocToolkit.Tests/DocxFixtures.cs
git commit -m "feat(core): find template table rows without descending into nested tables

Discovery walks each table's direct child rows and detects the marker in a
row's own cells' direct paragraphs. Descendants<TableRow>() would yield the
rows of tables nested inside cells, which would both mis-identify a
container as a template row and sweep an inner row into the outer
expansion - the same shape as the documented Descendants<Paragraph>()
text-box bug.

Rows come back innermost-first, so a nested template row is expanded
before any row containing it."
```

---

### Task 2: Expand one template row per record

**Files:**
- Modify: `src/DocToolkit/DocxEditor.cs`
- Test: `tests/DocToolkit.Tests/DocxEditorFillRowsTests.cs` (create)

**Interfaces:**
- Consumes: `TableRowFinder.Find` from Task 1; `RunTextSplicer.Apply` (already `internal`).
- Produces: `public static byte[] FillRows(byte[] docx, string collection,
  IEnumerable<IReadOnlyDictionary<string, string>> rows)` and the private `FillRowsCore` both
  overloads will share.

- [ ] **Step 1: Write the failing test**

Create `tests/DocToolkit.Tests/DocxEditorFillRowsTests.cs`:

```csharp
namespace DocToolkit.Tests;

public class DocxEditorFillRowsTests
{
    private static IReadOnlyDictionary<string, string> Row(string desc, string qty) =>
        new Dictionary<string, string> { ["Desc"] = desc, ["Qty"] = qty };

    [Fact]
    public void FillRows_ProducesOneRowPerRecordInOrder()
    {
        var docx = DocxFixtures.TableDocx("Description", "{{item.Desc}} x{{item.Qty}}");

        var filled = DocxEditor.FillRows(docx, "item", new[]
        {
            Row("Widget", "2"),
            Row("Gadget", "5"),
            Row("Doohickey", "1"),
        });

        var text = DocxEditor.ExtractText(filled);
        Assert.Contains("Widget x2", text);
        Assert.Contains("Gadget x5", text);
        Assert.Contains("Doohickey x1", text);
        Assert.DoesNotContain("{{item.", text);
        Assert.True(text.IndexOf("Widget", StringComparison.Ordinal)
                  < text.IndexOf("Gadget", StringComparison.Ordinal), "records kept their order");
    }
}
```

- [ ] **Step 2: Run it and watch it fail**

Run: `dotnet test tests/DocToolkit.Tests -c Release --filter FullyQualifiedName~DocxEditorFillRowsTests`
Expected: **compile failure** — `FillRows` does not exist.

- [ ] **Step 3: Implement `FillRows` and `FillRowsCore`**

Add to `src/DocToolkit/DocxEditor.cs` (in-place edit; do not replace the file):

```csharp
    /// <summary>
    /// Expands a table row once per record, so a template can render a variable-length list such as
    /// invoice line items.
    ///
    /// A row is a <b>template row</b> when one of its cells contains a placeholder prefixed with
    /// <paramref name="collection"/> — for example <c>{{item.Desc}}</c> when
    /// <paramref name="collection"/> is <c>item</c>. Each record deep-clones that row, so every
    /// clone keeps the template's run formatting, cell shading and borders, and substitution runs
    /// through the same splicer <see cref="ReplaceText"/> uses — a placeholder split across runs is
    /// still replaced, and a hyperlink in a cell is left intact.
    ///
    /// <b>Dictionary keys are bare field names</b> (<c>Desc</c>), not full placeholders — unlike
    /// <see cref="ReplaceText"/>, whose keys are the placeholder text including braces. The
    /// collection name is already an argument, so repeating it in every key of every record would
    /// duplicate it many times over.
    ///
    /// A placeholder with no matching key resolves to empty rather than being left visible.
    /// Placeholders for other prefixes are untouched, so a second call fills a second table. An
    /// empty <paramref name="rows"/> removes the template row, and removes the whole table when
    /// that row was its only one.
    ///
    /// Compose with <see cref="ReplaceText"/> for document-level scalars, expanding rows first.
    /// </summary>
    /// <exception cref="ArgumentNullException">Any argument is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty, or <paramref name="collection"/> is blank.</exception>
    /// <exception cref="DocumentConversionException">
    /// The package could not be opened or edited, or no template row was found for
    /// <paramref name="collection"/> — a mismatch between the call and the template is a bug, not a
    /// no-op.
    /// </exception>
    public static byte[] FillRows(
        byte[] docx, string collection, IEnumerable<IReadOnlyDictionary<string, string>> rows)
    {
        ArgumentNullException.ThrowIfNull(docx);
        ArgumentNullException.ThrowIfNull(collection);
        ArgumentNullException.ThrowIfNull(rows);
        if (docx.Length == 0)
            throw new ArgumentException("DOCX content was empty.", nameof(docx));
        if (string.IsNullOrWhiteSpace(collection))
            throw new ArgumentException("Collection name was blank.", nameof(collection));

        using var ms = new MemoryStream();
        ms.Write(docx, 0, docx.Length);
        ms.Position = 0;

        FillRowsCore(ms, collection, rows);
        return ms.ToArray();
    }

    /// <summary>The one real implementation; both overloads call it so they cannot drift apart.</summary>
    private static void FillRowsCore(
        MemoryStream package, string collection, IEnumerable<IReadOnlyDictionary<string, string>> rows)
    {
        var records = rows.ToList();
        var marker = "{{" + collection + ".";

        try
        {
            using var doc = WordprocessingDocument.Open(package, true);
            var body = doc.MainDocumentPart?.Document?.Body
                ?? throw new DocumentConversionException("The DOCX package had no document body.");

            var templates = TableRowFinder.Find(body, marker);
            if (templates.Count == 0)
            {
                throw new DocumentConversionException(
                    $"No table row containing '{marker}' was found, so there was nothing to fill.");
            }

            foreach (var template in templates)
                ExpandRow(template, collection, records);

            doc.MainDocumentPart!.Document.Save();
        }
        catch (DocumentConversionException)
        {
            throw;
        }
        catch (Exception ex)
        {
            throw new DocumentConversionException("Failed to fill table rows in the DOCX package.", ex);
        }
    }

    private static void ExpandRow(
        TableRow template, string collection, List<IReadOnlyDictionary<string, string>> records)
    {
        var parent = template.Parent
            ?? throw new DocumentConversionException("A template row had no parent table.");

        foreach (var record in records)
        {
            var clone = (TableRow)template.CloneNode(deep: true);
            Substitute(clone, collection, record);
            parent.InsertBefore(clone, template);
        }

        template.Remove();

        // A w:tbl with no w:tr is rejected by Word. Removing the table is better than writing a
        // package that saves without error and fails to open.
        if (parent is Table table && !table.ChildElements.OfType<TableRow>().Any())
            table.Remove();
    }

    /// <summary>
    /// Substitutes into the clone's own cells' direct paragraphs, scoped exactly as detection is,
    /// so a nested table's content is left to its own expansion.
    /// </summary>
    private static void Substitute(
        TableRow clone, string collection, IReadOnlyDictionary<string, string> record)
    {
        var replacements = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var pair in record)
            replacements["{{" + collection + "." + pair.Key + "}}"] = pair.Value ?? string.Empty;

        foreach (var cell in clone.ChildElements.OfType<TableCell>())
        {
            foreach (var paragraph in cell.ChildElements.OfType<Paragraph>())
            {
                var texts = paragraph.Descendants<Text>().ToList();
                RunTextSplicer.Apply(texts, t => t.Text, (t, v) => t.Text = v, replacements);
            }
        }

        // Any placeholder for this collection with no matching key resolves to empty rather than
        // staying visible to an end user. Done as a second pass because the keys are unknown until
        // the document is read.
        ClearUnmatched(clone, collection);
    }

    private static void ClearUnmatched(TableRow clone, string collection)
    {
        var marker = "{{" + collection + ".";
        foreach (var cell in clone.ChildElements.OfType<TableCell>())
        {
            foreach (var paragraph in cell.ChildElements.OfType<Paragraph>())
            {
                var texts = paragraph.Descendants<Text>().ToList();
                var merged = string.Concat(texts.Select(t => t.Text));
                var leftovers = new Dictionary<string, string>(StringComparer.Ordinal);

                var at = merged.IndexOf(marker, StringComparison.Ordinal);
                while (at >= 0)
                {
                    var close = merged.IndexOf("}}", at, StringComparison.Ordinal);
                    if (close < 0) break;
                    leftovers[merged[at..(close + 2)]] = string.Empty;
                    at = merged.IndexOf(marker, close, StringComparison.Ordinal);
                }

                if (leftovers.Count > 0)
                    RunTextSplicer.Apply(texts, t => t.Text, (t, v) => t.Text = v, leftovers);
            }
        }
    }
```

- [ ] **Step 4: Check how the existing code collects a paragraph's text nodes, and match it**

The sketch above uses `paragraph.Descendants<Text>()`. **Verify that against what `ReplaceTextCore`
already does before keeping it.** A `w:p` cannot contain a table, but it *can* contain a text box
(`w:txbxContent`), which nests whole paragraphs — and `DocxEditor`'s doc comment says text boxes are
deliberately "treated as the separate paragraphs they are." `Descendants<Text>()` would instead fold
a text box's content into the enclosing paragraph's node list, giving `FillRows` different
text-box semantics from `ReplaceText` for no stated reason.

Read `ReplaceTextCore` in the same file and reuse whatever paragraph/text collection it already
uses. If that logic is currently inline, extract it to a private helper and call it from both — one
behaviour, one implementation, which is the same reasoning that makes `*Core` methods exist here.

- [ ] **Step 5: Run the test and watch it pass**

Run: `dotnet test tests/DocToolkit.Tests -c Release --filter FullyQualifiedName~DocxEditorFillRowsTests`
Expected: PASS, 2 results.

- [ ] **Step 5: Verify the whole suite is still green**

```bash
dotnet build DocToolkit.sln -c Release -warnaserror
dotnet test  DocToolkit.sln -c Release --no-build
```

Expected: 0 warnings; all results pass.

- [ ] **Step 6: Commit**

```bash
git add src/DocToolkit/DocxEditor.cs tests/DocToolkit.Tests/DocxEditorFillRowsTests.cs
git commit -m "feat(core): add DocxEditor.FillRows for repeating table rows

Expands a template row once per record. Each record deep-clones the row, so
formatting, cell shading and borders survive, and substitution goes through
the same RunTextSplicer that ReplaceText uses rather than a second copy of
the offset-to-run mapping.

Keys are bare field names, unlike ReplaceText's full placeholders - the
collection name is already an argument. Unmatched placeholders resolve to
empty rather than staying visible.

Additive: no existing signature changes."
```

---

### Task 3: The behaviours that make it safe

**Files:**
- Modify: `tests/DocToolkit.Tests/DocxEditorFillRowsTests.cs`
- Modify: `src/DocToolkit/DocxEditor.cs` (only if a test exposes a defect)

**Interfaces:**
- Consumes: `FillRows` from Task 2.
- Produces: no new API — proof that the documented semantics hold.

- [ ] **Step 1: Verify the empty-table hazard empirically, before coding for it**

The design assumes Word rejects a `w:tbl` with no `w:tr`. **Check rather than assume.** Add a
temporary probe, run it, record the answer, then delete it:

```csharp
    [Fact]
    public void Probe_ZeroRowTableValidity()
    {
        var docx = DocxFixtures.TableDocx("{{item.Desc}}");
        var filled = DocxEditor.FillRows(docx, "item", Array.Empty<IReadOnlyDictionary<string,string>>());

        using var ms = new MemoryStream(filled);
        using var doc = WordprocessingDocument.Open(ms, false);
        var validator = new DocumentFormat.OpenXml.Validation.OpenXmlValidator();
        var errors = validator.Validate(doc).ToList();
        Assert.Fail($"validation errors: {errors.Count}\n" +
                    string.Join("\n", errors.Take(3).Select(e => e.Description)));
    }
```

Run it, read the output, then **delete this test**. If validation is clean the table-removal
behaviour is optional rather than required — keep it anyway for tidiness, and note the finding in
the design doc.

- [ ] **Step 2: Write the formatting-survival test — the point of the feature**

```csharp
    [Fact]
    public void FillRows_ClonesKeepTheTemplateRowFormatting()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var bold = new Run(new RunProperties(new Bold()), new Text("{{item.Desc}}"));
            var table = new Table(
                new TableRow(new TableCell(new Paragraph(bold))));
            doc.AddMainDocumentPart().Document = new Document(new Body(table));
            doc.MainDocumentPart!.Document.Save();
        }

        var filled = DocxEditor.FillRows(ms.ToArray(), "item", new[]
        {
            (IReadOnlyDictionary<string,string>)new Dictionary<string,string> { ["Desc"] = "Widget" },
            (IReadOnlyDictionary<string,string>)new Dictionary<string,string> { ["Desc"] = "Gadget" },
        });

        using var read = new MemoryStream(filled);
        using var opened = WordprocessingDocument.Open(read, false);
        var runs = opened.MainDocumentPart!.Document.Body!
            .Descendants<Run>().Where(r => r.InnerText.Length > 0).ToList();

        Assert.Equal(2, runs.Count);
        Assert.All(runs, r => Assert.NotNull(r.RunProperties?.Bold));
    }
```

- [ ] **Step 3: Write the remaining semantic tests**

```csharp
    [Fact]
    public void FillRows_SubstitutesAPlaceholderSplitAcrossRuns()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            // "{{item." + "Desc}}" - one visible placeholder, two runs.
            var paragraph = new Paragraph(
                new Run(new Text("{{item.") { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new Text("Desc}}") { Space = SpaceProcessingModeValues.Preserve }));
            doc.AddMainDocumentPart().Document =
                new Document(new Body(new Table(new TableRow(new TableCell(paragraph)))));
            doc.MainDocumentPart!.Document.Save();
        }

        var filled = DocxEditor.FillRows(ms.ToArray(), "item", new[]
        {
            (IReadOnlyDictionary<string,string>)new Dictionary<string,string> { ["Desc"] = "Widget" },
        });

        var text = DocxEditor.ExtractText(filled);
        Assert.Contains("Widget", text);
        Assert.DoesNotContain("{{item.", text);
    }

    [Fact]
    public void FillRows_LeavesOtherPrefixesAlone()
    {
        var docx = DocxFixtures.TableDocx("{{item.Desc}}", "{{payment.Total}}");

        var filled = DocxEditor.FillRows(docx, "item", new[]
        {
            (IReadOnlyDictionary<string,string>)new Dictionary<string,string> { ["Desc"] = "Widget" },
        });

        var text = DocxEditor.ExtractText(filled);
        Assert.Contains("Widget", text);
        Assert.Contains("{{payment.Total}}", text);   // a second FillRows call fills this
    }

    [Fact]
    public void FillRows_ResolvesAnUnmatchedPlaceholderToEmpty()
    {
        var docx = DocxFixtures.TableDocx("{{item.Desc}}|{{item.Missing}}");

        var filled = DocxEditor.FillRows(docx, "item", new[]
        {
            (IReadOnlyDictionary<string,string>)new Dictionary<string,string> { ["Desc"] = "Widget" },
        });

        var text = DocxEditor.ExtractText(filled);
        Assert.Contains("Widget|", text);
        Assert.DoesNotContain("{{item.Missing}}", text);
    }

    [Fact]
    public void FillRows_WithNoRecordsRemovesTheTemplateRow()
    {
        var docx = DocxFixtures.TableDocx("Header", "{{item.Desc}}");

        var filled = DocxEditor.FillRows(
            docx, "item", Array.Empty<IReadOnlyDictionary<string, string>>());

        var text = DocxEditor.ExtractText(filled);
        Assert.Contains("Header", text);
        Assert.DoesNotContain("{{item.", text);
    }

    [Fact]
    public void FillRows_KeepsAHyperlinkInACellIntact()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(new Body());
            var rel = main.AddHyperlinkRelationship(new Uri("https://example.com/"), true);

            var link = new Hyperlink(new Run(new Text("terms"))) { Id = rel.Id };
            var paragraph = new Paragraph(new Run(new Text("{{item.Desc}} ")), link);
            main.Document.Body!.Append(new Table(new TableRow(new TableCell(paragraph))));
            main.Document.Save();
        }

        var filled = DocxEditor.FillRows(ms.ToArray(), "item", new[]
        {
            (IReadOnlyDictionary<string,string>)new Dictionary<string,string> { ["Desc"] = "Widget" },
            (IReadOnlyDictionary<string,string>)new Dictionary<string,string> { ["Desc"] = "Gadget" },
        });

        using var read = new MemoryStream(filled);
        using var opened = WordprocessingDocument.Open(read, false);
        var links = opened.MainDocumentPart!.Document.Body!.Descendants<Hyperlink>().ToList();

        Assert.Equal(2, links.Count);                       // one per clone
        Assert.All(links, l => Assert.Equal("terms", l.InnerText));
        Assert.All(links, l => Assert.False(string.IsNullOrEmpty(l.Id?.Value)));
    }

    [Fact]
    public void FillRows_ExpandsTwoTemplateRowsIndependently()
    {
        // Both rows carry the same prefix. Design: each expands in its own right - clones of the
        // first, then clones of the second - NOT a repeating two-row block.
        var docx = DocxFixtures.TableDocx("A:{{item.Desc}}", "B:{{item.Desc}}");

        var filled = DocxEditor.FillRows(docx, "item", new[]
        {
            (IReadOnlyDictionary<string,string>)new Dictionary<string,string> { ["Desc"] = "one" },
            (IReadOnlyDictionary<string,string>)new Dictionary<string,string> { ["Desc"] = "two" },
        });

        var text = DocxEditor.ExtractText(filled);
        Assert.Contains("A:one", text);
        Assert.Contains("A:two", text);
        Assert.Contains("B:one", text);
        Assert.Contains("B:two", text);
        Assert.True(text.IndexOf("A:two", StringComparison.Ordinal)
                  < text.IndexOf("B:one", StringComparison.Ordinal),
                  "the first row's clones all precede the second row's");
    }

    [Fact]
    public void FillRows_ExpandsATemplateRowInsideANestedTable()
    {
        var inner = new Table(new TableRow(
            new TableCell(new Paragraph(new Run(new Text("{{item.Desc}}"))))));
        var outer = new Table(new TableRow(
            new TableCell(new Paragraph(new Run(new Text("container"))), inner)));

        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            doc.AddMainDocumentPart().Document = new Document(new Body(outer));
            doc.MainDocumentPart!.Document.Save();
        }

        var filled = DocxEditor.FillRows(ms.ToArray(), "item", new[]
        {
            (IReadOnlyDictionary<string,string>)new Dictionary<string,string> { ["Desc"] = "one" },
            (IReadOnlyDictionary<string,string>)new Dictionary<string,string> { ["Desc"] = "two" },
        });

        var text = DocxEditor.ExtractText(filled);
        Assert.Contains("one", text);
        Assert.Contains("two", text);
        // The container row was not itself cloned - "container" appears exactly once.
        Assert.Equal(1, text.Split("container").Length - 1);
    }

    [Fact]
    public void FillRows_ThrowsWhenNoTemplateRowMatches()
    {
        var docx = DocxFixtures.TableDocx("Header", "no placeholders here");

        var ex = Assert.Throws<DocumentConversionException>(() => DocxEditor.FillRows(
            docx, "item", new[]
            {
                (IReadOnlyDictionary<string,string>)new Dictionary<string,string> { ["Desc"] = "Widget" },
            }));

        Assert.Contains("{{item.", ex.Message);
    }

    [Theory]
    [InlineData(null, "item")]
    [InlineData(new byte[0], "item")]
    public void FillRows_RejectsBadArguments(byte[]? docx, string collection)
    {
        var records = Array.Empty<IReadOnlyDictionary<string, string>>();
        if (docx is null)
            Assert.Throws<ArgumentNullException>(() => DocxEditor.FillRows(null!, collection, records));
        else
            Assert.Throws<ArgumentException>(() => DocxEditor.FillRows(docx, collection, records));
    }

    [Fact]
    public void FillRows_RejectsABlankCollectionName()
    {
        var docx = DocxFixtures.TableDocx("{{item.Desc}}");
        Assert.Throws<ArgumentException>(() => DocxEditor.FillRows(
            docx, " ", Array.Empty<IReadOnlyDictionary<string, string>>()));
    }
```

- [ ] **Step 4: Run them all**

Run: `dotnet test tests/DocToolkit.Tests -c Release --filter FullyQualifiedName~DocxEditorFillRowsTests`

Expected: all PASS. **A failure here is a real defect in Task 2's implementation** — fix the
implementation, not the test, unless the test itself encodes something the design doc contradicts.

- [ ] **Step 5: Commit**

```bash
git add tests/DocToolkit.Tests/DocxEditorFillRowsTests.cs src/DocToolkit/DocxEditor.cs
git commit -m "test(core): prove the FillRows semantics hold

Formatting survives cloning, a placeholder split across runs is still
substituted, other prefixes are untouched so a second call fills a second
table, an unmatched placeholder resolves to empty rather than staying
visible, an empty collection removes the template row, and a call that
matches no template row throws rather than silently succeeding."
```

---

### Task 4: The `Stream` overload, and registering it where it cannot escape the suite

**Files:**
- Modify: `src/DocToolkit/DocxEditor.cs`
- Modify: `tests/DocToolkit.Tests/StreamOverloadTests.cs` — **the name lists at the top**
- Modify: `tests/DocToolkit.Tests/DocxEditorFillRowsTests.cs`

**Interfaces:**
- Consumes: `FillRowsCore` from Task 2, `StreamPipeline`.
- Produces: `public static Task FillRowsAsync(Stream source, string collection,
  IEnumerable<IReadOnlyDictionary<string, string>> rows, Stream destination,
  CancellationToken ct = default)`.

`CLAUDE.md`: *"If you add a new `Stream` overload, add it to the name lists at the top of that file
— an overload missing from those lists is the only way to escape the whole suite."* Step 3 is that
registration and is **not optional**.

- [ ] **Step 1: Write the failing parity test**

```csharp
    [Fact]
    public async Task FillRowsAsync_MatchesTheByteArrayOverload()
    {
        var docx = DocxFixtures.TableDocx("Header", "{{item.Desc}}");
        var records = new[]
        {
            (IReadOnlyDictionary<string,string>)new Dictionary<string,string> { ["Desc"] = "Widget" },
            (IReadOnlyDictionary<string,string>)new Dictionary<string,string> { ["Desc"] = "Gadget" },
        };

        var expected = DocxEditor.FillRows(docx, "item", records);

        using var destination = new MemoryStream();
        await DocxEditor.FillRowsAsync(new MemoryStream(docx), "item", records, destination);

        Assert.Equal(
            DocxEditor.ExtractText(expected),
            DocxEditor.ExtractText(destination.ToArray()));
    }
```

Parity is asserted on extracted text, not bytes. Two OpenXML saves are byte-deterministic (measured
2026-08-03), but text parity is what the method actually promises and does not depend on that
remaining true.

- [ ] **Step 2: Implement the overload**

```csharp
    /// <summary>
    /// Reads a .docx from <paramref name="source"/>, expands the template row once per record, and
    /// writes the result to <paramref name="destination"/>. See <see cref="FillRows"/> for what
    /// counts as a template row and how formatting survives — this overload applies identical logic.
    ///
    /// <paramref name="source"/> is <b>read</b> to its end and <paramref name="destination"/> is
    /// <b>written</b>; neither is disposed, closed or sought, and neither has to be seekable.
    /// </summary>
    /// <exception cref="ArgumentNullException">Any argument is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="source"/> is not readable or held no bytes, <paramref name="destination"/> is
    /// not writable, or <paramref name="collection"/> is blank.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The package could not be opened or edited.</exception>
    public static async Task FillRowsAsync(
        Stream source, string collection, IEnumerable<IReadOnlyDictionary<string, string>> rows,
        Stream destination, CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(collection);
        ArgumentNullException.ThrowIfNull(rows);
        StreamPipeline.RequireReadable(source, nameof(source));
        StreamPipeline.RequireWritable(destination, nameof(destination));
        if (string.IsNullOrWhiteSpace(collection))
            throw new ArgumentException("Collection name was blank.", nameof(collection));

        using var buffer = await StreamPipeline.DrainAsync(source, nameof(source), ct)
            .ConfigureAwait(false);

        FillRowsCore(buffer, collection, rows);

        await StreamPipeline.EmitAsync(buffer, destination, ct).ConfigureAwait(false);
    }
```

- [ ] **Step 3: Register the overload in `StreamOverloadTests` — mandatory**

Open `tests/DocToolkit.Tests/StreamOverloadTests.cs` and add `FillRowsAsync` to the method-name
lists at the top of the file, alongside `ReplaceTextAsync` and `ExtractTextAsync`. Match whatever
shape those entries already use.

- [ ] **Step 4: Run the stream suite and confirm the new overload is actually covered**

```bash
dotnet test tests/DocToolkit.Tests -c Release --filter FullyQualifiedName~StreamOverloadTests
```

Expected: PASS, **and the result count rises** versus before the registration. If the count is
unchanged, the name was not picked up — the overload is escaping the suite, which is the exact
failure `CLAUDE.md` warns about. Fix the registration before continuing.

- [ ] **Step 5: Full suite**

```bash
dotnet build DocToolkit.sln -c Release -warnaserror
dotnet test  DocToolkit.sln -c Release --no-build
```

- [ ] **Step 6: Commit**

```bash
git add src/DocToolkit/DocxEditor.cs tests/DocToolkit.Tests/StreamOverloadTests.cs tests/DocToolkit.Tests/DocxEditorFillRowsTests.cs
git commit -m "feat(core): add DocxEditor.FillRowsAsync

Follows the house Stream shape - source, then destination, then a
CancellationToken - and calls the same FillRowsCore as the byte[] overload
so the two cannot drift.

Registered in StreamOverloadTests' name lists: an overload missing from
those is the only way to escape the caller-owned-stream, forward-only and
genuinely-async guarantees."
```

---

### Task 5: Air-gap registration and documentation

**Files:**
- Modify: `tests/DocToolkit.Tests/AirGapGuardTests.cs`
- Modify: `src/DocToolkit/README.md`
- Modify: `README.md`
- Modify: `CLAUDE.md`

**Interfaces:**
- Consumes: everything above.
- Produces: the public-facing description of the feature.

- [ ] **Step 1: Add `FillRows` to the air-gap guard**

`AirGapGuardTests` asserts zero socket connections across the *whole* public API, and `README.md`
quantifies it. Add a case that runs `FillRows` against markup naming the loopback listener, in the
style the file already uses, and assert the accepted-connection count is still exactly zero.

- [ ] **Step 2: Run the guard**

```bash
dotnet test tests/DocToolkit.Tests -c Release --filter FullyQualifiedName~AirGapGuardTests
```

Expected: PASS, with a higher result count than before.

- [ ] **Step 3: Document it in the package README**

In `src/DocToolkit/README.md`, under the DOCX section, add:

````markdown
### Repeating table rows

A row whose cells contain `{{item.Field}}` placeholders repeats once per record:

```csharp
byte[] filled = DocxEditor.FillRows(docx, "item", new[]
{
    new Dictionary<string, string> { ["Desc"] = "Widget", ["Qty"] = "2" },
    new Dictionary<string, string> { ["Desc"] = "Gadget", ["Qty"] = "5" },
});

// then fill the document-level scalars
filled = DocxEditor.ReplaceText(filled, new() { ["{{customer}}"] = "Contoso Ltd" });
```

Every clone keeps the template row's formatting, shading and borders. Note that row keys are **bare
field names** (`Desc`), while `ReplaceText` keys are full placeholders (`{{customer}}`) — the
collection name is already an argument here.
````

- [ ] **Step 4: Mention it in the root README's usage block**

Add one line to the existing `## Usage` sample so the capability is visible on the landing page.

- [ ] **Step 5: Add the trap to `CLAUDE.md`**

Under *Traps in this codebase*, add a paragraph recording that `TableRowFinder` deliberately avoids
`Descendants<TableRow>()`, and why — matching the tone of the existing `DocxEditor`/text-box entry.

- [ ] **Step 6: Full verification**

```bash
dotnet build DocToolkit.sln -c Release -warnaserror
dotnet test  DocToolkit.sln -c Release --no-build
```

Expected: 0 warnings; every result passes.

- [ ] **Step 7: Commit and open the PR**

```bash
git add tests/DocToolkit.Tests/AirGapGuardTests.cs src/DocToolkit/README.md README.md CLAUDE.md
git commit -m "docs(core): document repeating table rows, and guard them offline

FillRows joins AirGapGuardTests, which asserts zero connections across the
whole public API - the guard's value is being exhaustive, and README
quantifies it.

Records in CLAUDE.md why TableRowFinder avoids Descendants<TableRow>(),
alongside the existing text-box entry it mirrors."
git push -u origin feat/repeating-table-rows
gh pr create --base main \
  --title "feat(core): repeating table rows for DOCX templates" \
  --body "Implements backlog item A4, per docs/2026-08-03-repeating-table-rows-design.md.

A table row whose cells contain \`{{item.Field}}\` placeholders now repeats once per record:

\`\`\`csharp
byte[] filled = DocxEditor.FillRows(docx, \"item\", lineItems);
filled = DocxEditor.ReplaceText(filled, new() { [\"{{customer}}\"] = \"Contoso Ltd\" });
\`\`\`

Additive - no existing signature changes. Each record deep-clones the template row and substitution
runs through the existing RunTextSplicer, so formatting, cell shading and hyperlink survival come
from already-proven code rather than a second copy of the offset-to-run mapping.

Discovery lives in its own TableRowFinder and deliberately avoids Descendants<TableRow>(), which
would reach into nested tables - the same shape as the documented Descendants<Paragraph>()
text-box bug. That behaviour is tested in isolation rather than only end-to-end.

FillRowsAsync is registered in StreamOverloadTests' name lists and FillRows in AirGapGuardTests,
both of which are exhaustive-by-design and silently weakened by an omission."
```

---

## Notes for the reviewer

- **Task 1 exists to make the nested-table trap testable in isolation.** Folding discovery into
  `DocxEditor` would leave the most dangerous behaviour in the feature provable only end-to-end.
- **Task 3 Step 1 is a probe, not a test** — it is written to fail so its output is visible, read,
  then deleted. Do not commit it.
- **Task 4 Step 3 is the step most likely to be skipped** and the one with the quietest
  consequence: the overload silently escapes every stream guarantee. Step 4's result-count check is
  how you know it worked.
- **This ships a release on merge.** The whole feature should land in one PR.
