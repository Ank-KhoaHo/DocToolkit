using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocToolkit.Tests;

public class DocxEditorFootnoteEndnoteTocTests
{
    /// <summary>
    /// Asserts the package is schema-valid, not merely readable — the same discipline
    /// <see cref="DocxEditorReplaceImageTests"/> uses for the same reason: extracted text says
    /// nothing about whether Word will open the file.
    /// </summary>
    private static void AssertValid(byte[] docx)
    {
        var errors = DocxFixtures.Validate(docx);
        Assert.True(errors.Count == 0,
            "expected a schema-valid package, got:\n" +
            string.Join("\n", errors.Take(3).Select(e => "  " + e.Description)));
    }

    // -----------------------------------------------------------------------------------------
    // A71: AddFootnote
    // -----------------------------------------------------------------------------------------

    [Fact]
    public void AddFootnote_InsertsAReferenceAndCreatesTheFootnotesPart()
    {
        var docx = DocxFixtures.Build(DocxFixtures.P(DocxFixtures.R("See the note{{note}} here.")));

        var filled = DocxEditor.AddFootnote(docx, "{{note}}", "This is the footnote text.");

        using var ms = new MemoryStream(filled);
        using var doc = WordprocessingDocument.Open(ms, false);

        Assert.NotNull(doc.MainDocumentPart!.FootnotesPart);
        var footnote = doc.MainDocumentPart.FootnotesPart!.Footnotes!.Elements<Footnote>().Single();
        Assert.Equal("This is the footnote text.", footnote.InnerText);

        var reference = doc.MainDocumentPart.Document!.Body!.Descendants<FootnoteReference>().Single();
        Assert.Equal(footnote.Id!.Value, reference.Id!.Value);

        var text = DocxEditor.ExtractText(filled);
        Assert.Contains("See the note", text);
        Assert.Contains(" here.", text);
        Assert.DoesNotContain("{{note}}", text);

        AssertValid(filled);
    }

    [Fact]
    public void AddFootnote_MatchesAPlaceholderSplitAcrossRuns()
    {
        var docx = DocxFixtures.Build(DocxFixtures.P(
            DocxFixtures.R("start {{no"),
            DocxFixtures.R("te}} end")));

        var filled = DocxEditor.AddFootnote(docx, "{{note}}", "Split-run footnote.");

        var text = DocxEditor.ExtractText(filled);
        Assert.Contains("start ", text);
        Assert.Contains(" end", text);
        Assert.DoesNotContain("{{no", text);

        using var ms = new MemoryStream(filled);
        using var doc = WordprocessingDocument.Open(ms, false);
        Assert.Single(doc.MainDocumentPart!.Document!.Body!.Descendants<FootnoteReference>());

        AssertValid(filled);
    }

    [Fact]
    public void AddFootnote_OnADocumentWithAnExistingFootnote_PicksTheNextId()
    {
        // First call creates footnote id 1; the second must not collide with it.
        var docx = DocxFixtures.Build(DocxFixtures.P(
            DocxFixtures.R("first{{a}} and second{{b}}")));

        var oneAdded = DocxEditor.AddFootnote(docx, "{{a}}", "First footnote.");
        var bothAdded = DocxEditor.AddFootnote(oneAdded, "{{b}}", "Second footnote.");

        using var ms = new MemoryStream(bothAdded);
        using var doc = WordprocessingDocument.Open(ms, false);
        var ids = doc.MainDocumentPart!.FootnotesPart!.Footnotes!.Elements<Footnote>()
            .Select(f => (int)f.Id!.Value).OrderBy(id => id).ToList();

        Assert.Equal(new[] { 1, 2 }, ids);
        AssertValid(bothAdded);
    }

    [Fact]
    public void AddFootnote_EmbedsOneEntryPerOccurrence()
    {
        var docx = DocxFixtures.Build(DocxFixtures.P(
            DocxFixtures.R("one{{note}} two{{note}} three")));

        var filled = DocxEditor.AddFootnote(docx, "{{note}}", "Repeated footnote text.");

        using var ms = new MemoryStream(filled);
        using var doc = WordprocessingDocument.Open(ms, false);
        Assert.Equal(2, doc.MainDocumentPart!.FootnotesPart!.Footnotes!.Elements<Footnote>().Count());
        Assert.Equal(2, doc.MainDocumentPart.Document!.Body!.Descendants<FootnoteReference>().Count());

        AssertValid(filled);
    }

    [Fact]
    public void AddFootnote_ThrowsWhenThePlaceholderIsAbsent()
    {
        var docx = DocxFixtures.Build(DocxFixtures.P(DocxFixtures.R("nothing to add a note to")));

        var ex = Assert.Throws<DocumentConversionException>(
            () => DocxEditor.AddFootnote(docx, "{{note}}", "text"));

        Assert.Contains("{{note}}", ex.Message);
    }

    [Fact]
    public void AddFootnote_RejectsNullArguments()
    {
        var docx = DocxFixtures.Build(DocxFixtures.P(DocxFixtures.R("{{note}}")));

        Assert.Throws<ArgumentNullException>(() => DocxEditor.AddFootnote(null!, "{{note}}", "text"));
        Assert.Throws<ArgumentNullException>(() => DocxEditor.AddFootnote(docx, null!, "text"));
        Assert.Throws<ArgumentNullException>(() => DocxEditor.AddFootnote(docx, "{{note}}", null!));
    }

    [Fact]
    public async Task AddFootnoteAsync_FromFile_MatchesTheByteArrayOverload()
    {
        var docx = DocxFixtures.Build(DocxFixtures.P(DocxFixtures.R("See{{note}}here.")));

        using var input = new TempFile();
        using var output = new TempFile();
        await File.WriteAllBytesAsync(input.Path, docx);

        await DocxEditor.AddFootnoteAsync(input.Path, output.Path, "{{note}}", "Async footnote.");

        var expected = DocxEditor.AddFootnote(docx, "{{note}}", "Async footnote.");
        var actual = await File.ReadAllBytesAsync(output.Path);

        Assert.Equal(DocxEditor.ExtractText(expected), DocxEditor.ExtractText(actual));
        AssertValid(actual);
    }

    // -----------------------------------------------------------------------------------------
    // A71: AddEndnote
    // -----------------------------------------------------------------------------------------

    [Fact]
    public void AddEndnote_InsertsAReferenceAndCreatesTheEndnotesPart()
    {
        var docx = DocxFixtures.Build(DocxFixtures.P(DocxFixtures.R("See the note{{note}} here.")));

        var filled = DocxEditor.AddEndnote(docx, "{{note}}", "This is the endnote text.");

        using var ms = new MemoryStream(filled);
        using var doc = WordprocessingDocument.Open(ms, false);

        Assert.NotNull(doc.MainDocumentPart!.EndnotesPart);
        var endnote = doc.MainDocumentPart.EndnotesPart!.Endnotes!.Elements<Endnote>().Single();
        Assert.Equal("This is the endnote text.", endnote.InnerText);

        var reference = doc.MainDocumentPart.Document!.Body!.Descendants<EndnoteReference>().Single();
        Assert.Equal(endnote.Id!.Value, reference.Id!.Value);

        var text = DocxEditor.ExtractText(filled);
        Assert.Contains("See the note", text);
        Assert.Contains(" here.", text);
        Assert.DoesNotContain("{{note}}", text);

        AssertValid(filled);
    }

    [Fact]
    public void AddEndnote_MatchesAPlaceholderSplitAcrossRuns()
    {
        var docx = DocxFixtures.Build(DocxFixtures.P(
            DocxFixtures.R("start {{no"),
            DocxFixtures.R("te}} end")));

        var filled = DocxEditor.AddEndnote(docx, "{{note}}", "Split-run endnote.");

        var text = DocxEditor.ExtractText(filled);
        Assert.Contains("start ", text);
        Assert.Contains(" end", text);
        Assert.DoesNotContain("{{no", text);

        using var ms = new MemoryStream(filled);
        using var doc = WordprocessingDocument.Open(ms, false);
        Assert.Single(doc.MainDocumentPart!.Document!.Body!.Descendants<EndnoteReference>());

        AssertValid(filled);
    }

    [Fact]
    public void AddEndnote_EmbedsOneEntryPerOccurrence()
    {
        var docx = DocxFixtures.Build(DocxFixtures.P(
            DocxFixtures.R("one{{note}} two{{note}} three")));

        var filled = DocxEditor.AddEndnote(docx, "{{note}}", "Repeated endnote text.");

        using var ms = new MemoryStream(filled);
        using var doc = WordprocessingDocument.Open(ms, false);
        Assert.Equal(2, doc.MainDocumentPart!.EndnotesPart!.Endnotes!.Elements<Endnote>().Count());
        Assert.Equal(2, doc.MainDocumentPart.Document!.Body!.Descendants<EndnoteReference>().Count());

        AssertValid(filled);
    }

    [Fact]
    public void AddEndnote_OnADocumentWithAnExistingEndnote_PicksTheNextId()
    {
        // First call creates endnote id 1; the second must not collide with it.
        var docx = DocxFixtures.Build(DocxFixtures.P(
            DocxFixtures.R("first{{a}} and second{{b}}")));

        var oneAdded = DocxEditor.AddEndnote(docx, "{{a}}", "First endnote.");
        var bothAdded = DocxEditor.AddEndnote(oneAdded, "{{b}}", "Second endnote.");

        using var ms = new MemoryStream(bothAdded);
        using var doc = WordprocessingDocument.Open(ms, false);
        var ids = doc.MainDocumentPart!.EndnotesPart!.Endnotes!.Elements<Endnote>()
            .Select(e => (int)e.Id!.Value).OrderBy(id => id).ToList();

        Assert.Equal(new[] { 1, 2 }, ids);
        AssertValid(bothAdded);
    }

    [Fact]
    public void AddEndnote_AndAddFootnote_UseIndependentIdSpaces()
    {
        // Measured: footnote and endnote ids are independent numbering spaces. Adding a footnote
        // then an endnote to the same document must not produce a collision, and neither should
        // be forced to skip past the other's ids.
        var docx = DocxFixtures.Build(DocxFixtures.P(
            DocxFixtures.R("first{{a}} and second{{b}}")));

        var withFootnote = DocxEditor.AddFootnote(docx, "{{a}}", "A footnote.");
        var withBoth = DocxEditor.AddEndnote(withFootnote, "{{b}}", "An endnote.");

        using var ms = new MemoryStream(withBoth);
        using var doc = WordprocessingDocument.Open(ms, false);

        Assert.Equal(1, doc.MainDocumentPart!.FootnotesPart!.Footnotes!.Elements<Footnote>().Single().Id!.Value);
        Assert.Equal(1, doc.MainDocumentPart.EndnotesPart!.Endnotes!.Elements<Endnote>().Single().Id!.Value);

        AssertValid(withBoth);
    }

    [Fact]
    public void AddEndnote_ThrowsWhenThePlaceholderIsAbsent()
    {
        var docx = DocxFixtures.Build(DocxFixtures.P(DocxFixtures.R("nothing to add a note to")));

        var ex = Assert.Throws<DocumentConversionException>(
            () => DocxEditor.AddEndnote(docx, "{{note}}", "text"));

        Assert.Contains("{{note}}", ex.Message);
    }

    [Fact]
    public void AddEndnote_RejectsNullArguments()
    {
        var docx = DocxFixtures.Build(DocxFixtures.P(DocxFixtures.R("{{note}}")));

        Assert.Throws<ArgumentNullException>(() => DocxEditor.AddEndnote(null!, "{{note}}", "text"));
        Assert.Throws<ArgumentNullException>(() => DocxEditor.AddEndnote(docx, null!, "text"));
        Assert.Throws<ArgumentNullException>(() => DocxEditor.AddEndnote(docx, "{{note}}", null!));
    }

    [Fact]
    public async Task AddEndnoteAsync_FromFile_MatchesTheByteArrayOverload()
    {
        var docx = DocxFixtures.Build(DocxFixtures.P(DocxFixtures.R("See{{note}}here.")));

        using var input = new TempFile();
        using var output = new TempFile();
        await File.WriteAllBytesAsync(input.Path, docx);

        await DocxEditor.AddEndnoteAsync(input.Path, output.Path, "{{note}}", "Async endnote.");

        var expected = DocxEditor.AddEndnote(docx, "{{note}}", "Async endnote.");
        var actual = await File.ReadAllBytesAsync(output.Path);

        Assert.Equal(DocxEditor.ExtractText(expected), DocxEditor.ExtractText(actual));
        AssertValid(actual);
    }

    // -----------------------------------------------------------------------------------------
    // A71: AddTableOfContents
    // -----------------------------------------------------------------------------------------

    private static byte[] BuildDocumentWithHeadingsAndTocPlaceholder()
    {
        using var word = OfficeIMO.Word.WordDocument.Create();
        word.AddParagraph("{{toc}}");
        word.AddParagraph("Overview").Style = OfficeIMO.Word.WordParagraphStyles.Heading1;
        word.AddParagraph("Some intro text.");
        word.AddParagraph("Details").Style = OfficeIMO.Word.WordParagraphStyles.Heading2;
        word.AddParagraph("Some detail text.");

        using var ms = new MemoryStream();
        word.Save(ms);
        return ms.ToArray();
    }

    [Fact]
    public void AddTableOfContents_ReplacesThePlaceholderParagraphWithADirtyToc()
    {
        var docx = BuildDocumentWithHeadingsAndTocPlaceholder();

        var withToc = DocxEditor.AddTableOfContents(docx, "{{toc}}");

        using var ms = new MemoryStream(withToc);
        using var doc = WordprocessingDocument.Open(ms, false);
        var body = doc.MainDocumentPart!.Document!.Body!;

        // The placeholder is gone; the TOC's own field is present and marked dirty, matching the
        // exact mechanism this ticket's design measured: w:dirty is what keeps Word (and any
        // consumer honouring it) from trusting the cached "No table of contents entries found."
        Assert.DoesNotContain("{{toc}}", DocxEditor.ExtractText(withToc));

        var field = body.Descendants<SimpleField>().Single();
        Assert.True(field.Dirty?.Value);

        var updateFields = doc.MainDocumentPart.DocumentSettingsPart!.Settings!
            .Elements<UpdateFieldsOnOpen>().SingleOrDefault();
        Assert.NotNull(updateFields);
        Assert.True(updateFields!.Val?.Value);

        AssertValid(withToc);
    }

    [Fact]
    public void AddTableOfContents_RendersRealHeadingsWhenConvertedToPdf()
    {
        // The measurement this whole task is built on, pinned rather than only observed once:
        // DocToolkit's OWN DocxToPdfConverter recomputes the field live, so the rendered PDF shows
        // real heading text and page numbers, not the field's own cached placeholder string.
        var docx = BuildDocumentWithHeadingsAndTocPlaceholder();

        var withToc = DocxEditor.AddTableOfContents(docx, "{{toc}}");
        var pdf = DocxToPdfConverter.Convert(withToc);
        var pageText = string.Join("\n", PdfEditor.ExtractText(pdf));

        Assert.DoesNotContain("No table of contents entries found.", pageText);
        Assert.Contains("Overview", pageText);
        Assert.Contains("Details", pageText);
    }

    [Fact]
    public void AddTableOfContents_RespectsMinAndMaxLevel()
    {
        using var word = OfficeIMO.Word.WordDocument.Create();
        word.AddParagraph("{{toc}}");
        word.AddParagraph("Top").Style = OfficeIMO.Word.WordParagraphStyles.Heading1;
        word.AddParagraph("Deep").Style = OfficeIMO.Word.WordParagraphStyles.Heading3;
        using var ms = new MemoryStream();
        word.Save(ms);
        var docx = ms.ToArray();

        var withToc = DocxEditor.AddTableOfContents(docx, "{{toc}}", minLevel: 1, maxLevel: 1);

        using var check = new MemoryStream(withToc);
        using var doc = WordprocessingDocument.Open(check, false);
        var instr = doc.MainDocumentPart!.Document!.Body!.Descendants<SimpleField>().Single()
            .Instruction!.Value!;
        Assert.Contains("\"1-1\"", instr);
    }

    [Fact]
    public void AddTableOfContents_RejectsAnOutOfRangeLevel()
    {
        var docx = BuildDocumentWithHeadingsAndTocPlaceholder();

        Assert.Throws<ArgumentOutOfRangeException>(
            () => DocxEditor.AddTableOfContents(docx, "{{toc}}", minLevel: 0, maxLevel: 3));
        Assert.Throws<ArgumentOutOfRangeException>(
            () => DocxEditor.AddTableOfContents(docx, "{{toc}}", minLevel: 1, maxLevel: 10));
        Assert.Throws<ArgumentException>(
            () => DocxEditor.AddTableOfContents(docx, "{{toc}}", minLevel: 3, maxLevel: 1));
    }

    [Fact]
    public void AddTableOfContents_ThrowsWhenThePlaceholderIsAbsent()
    {
        var docx = DocxFixtures.Build(DocxFixtures.P(DocxFixtures.R("nothing here")));

        var ex = Assert.Throws<DocumentConversionException>(
            () => DocxEditor.AddTableOfContents(docx, "{{toc}}"));

        Assert.Contains("{{toc}}", ex.Message);
    }

    [Fact]
    public void AddTableOfContents_RefusesAPlaceholderParagraphWithOtherContent()
    {
        // Refuses rather than silently discarding the paragraph's other content -- replacing a
        // whole paragraph has no way to preserve a neighbour the way an inline splice can.
        var docx = DocxFixtures.Build(DocxFixtures.P(DocxFixtures.R("Table of contents: {{toc}}")));

        var ex = Assert.Throws<DocumentConversionException>(
            () => DocxEditor.AddTableOfContents(docx, "{{toc}}"));

        Assert.Contains("{{toc}}", ex.Message);
        Assert.Contains("other text", ex.Message);
    }

    [Fact]
    public void AddTableOfContents_RejectsNullArguments()
    {
        var docx = BuildDocumentWithHeadingsAndTocPlaceholder();

        Assert.Throws<ArgumentNullException>(() => DocxEditor.AddTableOfContents(null!, "{{toc}}"));
        Assert.Throws<ArgumentNullException>(() => DocxEditor.AddTableOfContents(docx, null!));
    }

    [Fact]
    public async Task AddTableOfContentsAsync_FromFile_MatchesTheByteArrayOverload()
    {
        var docx = BuildDocumentWithHeadingsAndTocPlaceholder();

        using var input = new TempFile();
        using var output = new TempFile();
        await File.WriteAllBytesAsync(input.Path, docx);

        await DocxEditor.AddTableOfContentsAsync(input.Path, output.Path, "{{toc}}");

        var expectedText = DocxEditor.ExtractText(DocxEditor.AddTableOfContents(docx, "{{toc}}"));
        var actual = await File.ReadAllBytesAsync(output.Path);

        Assert.Equal(expectedText, DocxEditor.ExtractText(actual));
        AssertValid(actual);
    }
}
