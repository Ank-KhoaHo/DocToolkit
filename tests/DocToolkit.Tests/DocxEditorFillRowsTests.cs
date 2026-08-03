using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocToolkit.Tests;

public class DocxEditorFillRowsTests
{
    /// <summary>Row data keyed by bare field name, as <c>FillRows</c> expects.</summary>
    private static IReadOnlyDictionary<string, string> Rec(params (string Key, string Value)[] fields)
        => fields.ToDictionary(f => f.Key, f => f.Value, StringComparer.Ordinal);

    /// <summary>
    /// Asserts the package is schema-valid, not merely readable.
    ///
    /// Worth its own helper because a fixture in this very file once built tables without a
    /// <c>w:tblGrid</c>: every assertion passed while the documents were invalid. Extracted text
    /// tells you what a document says, never whether Word will open it.
    /// </summary>
    private static void AssertValid(byte[] docx)
    {
        using var ms = new MemoryStream(docx);
        using var doc = WordprocessingDocument.Open(ms, false);
        var errors = new OpenXmlValidator().Validate(doc).ToList();
        Assert.True(errors.Count == 0,
            "expected a schema-valid package, got:\n" +
            string.Join("\n", errors.Take(3).Select(e => "  " + e.Description)));
    }

    [Fact]
    public void FillRows_ProducesOneRowPerRecordInOrder()
    {
        var docx = DocxFixtures.Build(DocxFixtures.Tbl(
            DocxFixtures.Row(DocxFixtures.R("Description")),
            DocxFixtures.Row(DocxFixtures.R("{{item.Desc}} x{{item.Qty}}"))));

        var filled = DocxEditor.FillRows(docx, "item", new[]
        {
            Rec(("Desc", "Widget"), ("Qty", "2")),
            Rec(("Desc", "Gadget"), ("Qty", "5")),
            Rec(("Desc", "Doohickey"), ("Qty", "1")),
        });

        var text = DocxEditor.ExtractText(filled);

        Assert.Contains("Widget x2", text);
        Assert.Contains("Gadget x5", text);
        Assert.Contains("Doohickey x1", text);
        Assert.DoesNotContain("{{item.", text);
        Assert.Contains("Description", text);   // the header row is untouched
        Assert.True(
            text.IndexOf("Widget", StringComparison.Ordinal) < text.IndexOf("Gadget", StringComparison.Ordinal),
            "records keep the order they were given in");
    }

    [Fact]
    public void FillRows_ProducesExactlyOneTableRowPerRecord()
    {
        var docx = DocxFixtures.Build(DocxFixtures.Tbl(
            DocxFixtures.Row(DocxFixtures.R("Description")),
            DocxFixtures.Row(DocxFixtures.R("{{item.Desc}}"))));

        var filled = DocxEditor.FillRows(docx, "item", new[]
        {
            Rec(("Desc", "Widget")),
            Rec(("Desc", "Gadget")),
        });

        using var ms = new MemoryStream(filled);
        using var doc = WordprocessingDocument.Open(ms, false);
        var rows = doc.MainDocumentPart!.Document!.Body!
            .ChildElements.OfType<Table>().Single()
            .ChildElements.OfType<TableRow>().ToList();

        Assert.Equal(3, rows.Count);          // header + two records; the template row is gone
        Assert.Contains("Description", rows[0].InnerText);
        Assert.Contains("Widget", rows[1].InnerText);
        Assert.Contains("Gadget", rows[2].InnerText);
        AssertValid(filled);
    }

    [Fact]
    public void FillRows_ClonesKeepTheTemplateRowFormatting()
    {
        var docx = DocxFixtures.Build(DocxFixtures.Tbl(
            DocxFixtures.Row(DocxFixtures.R("{{item.Desc}}", bold: true))));

        var filled = DocxEditor.FillRows(docx, "item", new[]
        {
            Rec(("Desc", "Widget")),
            Rec(("Desc", "Gadget")),
        });

        using var ms = new MemoryStream(filled);
        using var doc = WordprocessingDocument.Open(ms, false);
        var runs = doc.MainDocumentPart!.Document!.Body!
            .Descendants<Run>().Where(r => r.InnerText.Length > 0).ToList();

        Assert.Equal(2, runs.Count);
        Assert.All(runs, r => Assert.NotNull(r.RunProperties?.Bold));
        AssertValid(filled);
    }

    [Fact]
    public void FillRows_SubstitutesAPlaceholderSplitAcrossRuns()
    {
        // One visible placeholder, two runs - the case a per-run string.Replace misses entirely.
        var docx = DocxFixtures.Build(DocxFixtures.Tbl(
            DocxFixtures.Row(DocxFixtures.R("{{item."), DocxFixtures.R("Desc}}"))));

        var filled = DocxEditor.FillRows(docx, "item", new[] { Rec(("Desc", "Widget")) });

        var text = DocxEditor.ExtractText(filled);
        Assert.Contains("Widget", text);
        Assert.DoesNotContain("{{item.", text);
    }

    [Fact]
    public void FillRows_KeepsAHyperlinkInACellIntact()
    {
        using var built = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(built, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(new Body());
            var rel = main.AddHyperlinkRelationship(new Uri("https://example.com/"), true);

            var row = DocxFixtures.RowOf(new Paragraph(
                DocxFixtures.R("{{item.Desc}} "),
                new Hyperlink(DocxFixtures.R("terms")) { Id = rel.Id }));

            main.Document!.Body!.Append(DocxFixtures.Tbl(row));
            main.Document.Save();
        }

        var filled = DocxEditor.FillRows(built.ToArray(), "item", new[]
        {
            Rec(("Desc", "Widget")),
            Rec(("Desc", "Gadget")),
        });

        using var ms = new MemoryStream(filled);
        using var opened = WordprocessingDocument.Open(ms, false);
        var links = opened.MainDocumentPart!.Document!.Body!.Descendants<Hyperlink>().ToList();

        Assert.Equal(2, links.Count);                                     // one per clone
        Assert.All(links, l => Assert.Equal("terms", l.InnerText));       // text intact
        Assert.All(links, l => Assert.False(string.IsNullOrEmpty(l.Id?.Value)));  // relationship intact
    }

    [Fact]
    public void FillRows_LeavesOtherPrefixesAlone()
    {
        var docx = DocxFixtures.Build(DocxFixtures.Tbl(
            DocxFixtures.Row(DocxFixtures.R("{{item.Desc}}")),
            DocxFixtures.Row(DocxFixtures.R("{{payment.Total}}"))));

        var filled = DocxEditor.FillRows(docx, "item", new[] { Rec(("Desc", "Widget")) });

        var text = DocxEditor.ExtractText(filled);
        Assert.Contains("Widget", text);
        Assert.Contains("{{payment.Total}}", text);   // a second FillRows call fills this
    }

    [Fact]
    public void FillRows_ResolvesAnUnmatchedPlaceholderToEmpty()
    {
        var docx = DocxFixtures.Build(DocxFixtures.Tbl(
            DocxFixtures.Row(DocxFixtures.R("[{{item.Desc}}|{{item.Missing}}]"))));

        var filled = DocxEditor.FillRows(docx, "item", new[] { Rec(("Desc", "Widget")) });

        var text = DocxEditor.ExtractText(filled);
        Assert.Contains("[Widget|]", text);
        Assert.DoesNotContain("{{item.Missing}}", text);
    }

    [Fact]
    public void FillRows_WithNoRecordsRemovesTheTemplateRowButKeepsTheRest()
    {
        var docx = DocxFixtures.Build(DocxFixtures.Tbl(
            DocxFixtures.Row(DocxFixtures.R("Header")),
            DocxFixtures.Row(DocxFixtures.R("{{item.Desc}}"))));

        var filled = DocxEditor.FillRows(
            docx, "item", Array.Empty<IReadOnlyDictionary<string, string>>());

        var text = DocxEditor.ExtractText(filled);
        Assert.Contains("Header", text);
        Assert.DoesNotContain("{{item.", text);
        AssertValid(filled);
    }

    [Fact]
    public void FillRows_WithNoRecordsRemovesATableThatHeldOnlyTheTemplateRow()
    {
        var docx = DocxFixtures.Build(
            DocxFixtures.P(DocxFixtures.R("before")),
            DocxFixtures.Tbl(DocxFixtures.Row(DocxFixtures.R("{{item.Desc}}"))));

        var filled = DocxEditor.FillRows(
            docx, "item", Array.Empty<IReadOnlyDictionary<string, string>>());

        using var ms = new MemoryStream(filled);
        using var doc = WordprocessingDocument.Open(ms, false);
        Assert.Empty(doc.MainDocumentPart!.Document!.Body!.ChildElements.OfType<Table>());
        Assert.Contains("before", DocxEditor.ExtractText(filled));
        AssertValid(filled);
    }

    [Fact]
    public void FillRows_ExpandsTwoTemplateRowsIndependently()
    {
        // Both rows carry the prefix. Each expands in its own right - clones of the first, then
        // clones of the second - rather than the pair repeating as a block.
        var docx = DocxFixtures.Build(DocxFixtures.Tbl(
            DocxFixtures.Row(DocxFixtures.R("A:{{item.Desc}}")),
            DocxFixtures.Row(DocxFixtures.R("B:{{item.Desc}}"))));

        var filled = DocxEditor.FillRows(docx, "item", new[]
        {
            Rec(("Desc", "one")),
            Rec(("Desc", "two")),
        });

        var text = DocxEditor.ExtractText(filled);
        Assert.Contains("A:one", text);
        Assert.Contains("A:two", text);
        Assert.Contains("B:one", text);
        Assert.Contains("B:two", text);
        Assert.True(
            text.IndexOf("A:two", StringComparison.Ordinal) < text.IndexOf("B:one", StringComparison.Ordinal),
            "the first template row's clones all precede the second's");
    }

    [Fact]
    public void FillRows_ExpandsATemplateRowInsideANestedTableWithoutCloningItsContainer()
    {
        var inner = DocxFixtures.Tbl(DocxFixtures.Row(DocxFixtures.R("{{item.Desc}}")));
        var docx = DocxFixtures.Build(DocxFixtures.Tbl(
            DocxFixtures.RowOf(DocxFixtures.P(DocxFixtures.R("container")), inner)));

        var filled = DocxEditor.FillRows(docx, "item", new[]
        {
            Rec(("Desc", "one")),
            Rec(("Desc", "two")),
        });

        var text = DocxEditor.ExtractText(filled);
        Assert.Contains("one", text);
        Assert.Contains("two", text);
        // The container row was not itself a template row, so it appears exactly once.
        Assert.Equal(1, text.Split("container").Length - 1);
    }

    [Fact]
    public void FillRows_ThrowsWhenNoTemplateRowMatches()
    {
        var docx = DocxFixtures.Build(DocxFixtures.Tbl(
            DocxFixtures.Row(DocxFixtures.R("no placeholders here"))));

        var ex = Assert.Throws<DocumentConversionException>(
            () => DocxEditor.FillRows(docx, "item", new[] { Rec(("Desc", "Widget")) }));

        Assert.Contains("{{item.", ex.Message);
    }

    [Fact]
    public async Task FillRowsAsync_MatchesTheByteArrayOverload()
    {
        var docx = DocxFixtures.Build(DocxFixtures.Tbl(
            DocxFixtures.Row(DocxFixtures.R("Description")),
            DocxFixtures.Row(DocxFixtures.R("{{item.Desc}} x{{item.Qty}}"))));

        var records = new[]
        {
            Rec(("Desc", "Widget"), ("Qty", "2")),
            Rec(("Desc", "Gadget"), ("Qty", "5")),
        };

        var expected = DocxEditor.FillRows(docx, "item", records);

        using var destination = new MemoryStream();
        await DocxEditor.FillRowsAsync(new MemoryStream(docx), "item", records, destination);
        var actual = destination.ToArray();

        // Parity on readable content, not bytes. Two OpenXML saves happen to be byte-deterministic
        // (measured 2026-08-03), but text is what the method promises and does not depend on that
        // staying true.
        Assert.Equal(DocxEditor.ExtractText(expected), DocxEditor.ExtractText(actual));
        AssertValid(actual);
    }

    [Fact]
    public async Task FillRowsAsync_RejectsBadArguments()
    {
        var docx = DocxFixtures.Build(DocxFixtures.Tbl(
            DocxFixtures.Row(DocxFixtures.R("{{item.Desc}}"))));
        var none = Array.Empty<IReadOnlyDictionary<string, string>>();

        await Assert.ThrowsAsync<ArgumentNullException>(() =>
            DocxEditor.FillRowsAsync(new MemoryStream(docx), null!, none, new MemoryStream()));
        await Assert.ThrowsAsync<ArgumentNullException>(() =>
            DocxEditor.FillRowsAsync(new MemoryStream(docx), "item", null!, new MemoryStream()));
        await Assert.ThrowsAsync<ArgumentException>(() =>
            DocxEditor.FillRowsAsync(new MemoryStream(docx), " ", none, new MemoryStream()));
    }

    [Fact]
    public void FillRows_RejectsBadArguments()
    {
        var docx = DocxFixtures.Build(DocxFixtures.Tbl(
            DocxFixtures.Row(DocxFixtures.R("{{item.Desc}}"))));
        var none = Array.Empty<IReadOnlyDictionary<string, string>>();

        Assert.Throws<ArgumentNullException>(() => DocxEditor.FillRows(null!, "item", none));
        Assert.Throws<ArgumentNullException>(() => DocxEditor.FillRows(docx, null!, none));
        Assert.Throws<ArgumentNullException>(() => DocxEditor.FillRows(docx, "item", null!));
        Assert.Throws<ArgumentException>(() => DocxEditor.FillRows(Array.Empty<byte>(), "item", none));
        Assert.Throws<ArgumentException>(() => DocxEditor.FillRows(docx, " ", none));
    }
}
