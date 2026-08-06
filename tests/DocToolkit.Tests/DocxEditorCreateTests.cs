using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocToolkit.Tests;

public class DocxEditorCreateTests
{
    /// <summary>
    /// Asserts the package is schema-valid, not merely readable. Extracted text tells you what a
    /// document says, never whether Word will open it.
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
    public void Create_WithNoBlocks_ProducesAValidEmptyDocument()
    {
        // Deliberately valid: this is the "blank DOCX to then edit" case, and a body with no
        // content is schema-valid.
        var docx = DocxEditor.Create(Array.Empty<DocxBlock>());

        AssertValid(docx);
        Assert.Equal("", DocxEditor.ExtractText(docx).Trim());
    }

    [Fact]
    public void Create_WritesParagraphsInOrder()
    {
        var docx = DocxEditor.Create(new[]
        {
            DocxBlock.Paragraph("First."),
            DocxBlock.Paragraph("Second."),
        });

        AssertValid(docx);
        var text = DocxEditor.ExtractText(docx);
        Assert.Contains("First.", text);
        Assert.Contains("Second.", text);
        Assert.True(text.IndexOf("First.", StringComparison.Ordinal)
                  < text.IndexOf("Second.", StringComparison.Ordinal),
            "blocks must appear in the order they were given");
    }

    [Fact]
    public void Create_RejectsNullBlocks()
        => Assert.Throws<ArgumentNullException>(() => DocxEditor.Create(null!));

    [Fact]
    public void Create_RejectsANullBlockInTheSequence()
        => Assert.Throws<ArgumentException>(
            () => DocxEditor.Create(new DocxBlock?[] { DocxBlock.Paragraph("ok"), null }!));

    /// <summary>
    /// The output must be a real document, not merely one that opens: the existing editing API has
    /// to work against it.
    /// </summary>
    [Fact]
    public void Create_ProducesADocumentTheEditingApiCanEdit()
    {
        var docx = DocxEditor.Create(new[] { DocxBlock.Paragraph("Dear {{customer}}, hello.") });

        var edited = DocxEditor.ReplaceText(docx, new Dictionary<string, string>
        {
            ["{{customer}}"] = "Acme",
        });

        AssertValid(edited);
        Assert.Contains("Dear Acme, hello.", DocxEditor.ExtractText(edited));
    }

    [Fact]
    public void Create_WritesHeadingsAsRealWordHeadingStyles()
    {
        var docx = DocxEditor.Create(new[]
        {
            DocxBlock.Heading("Title", 1),
            DocxBlock.Heading("Section", 2),
            DocxBlock.Paragraph("Body."),
        });

        AssertValid(docx);

        using var ms = new MemoryStream(docx);
        using var doc = WordprocessingDocument.Open(ms, false);

        var styleIds = doc.MainDocumentPart!.Document!.Body!
            .Descendants<ParagraphStyleId>().Select(s => s.Val?.Value).ToList();
        Assert.Contains("Heading1", styleIds);
        Assert.Contains("Heading2", styleIds);

        // The reference alone is not enough. A document that references Heading1 without DEFINING
        // it renders as plain text - no navigation pane entry, no table of contents - and no schema
        // check would ever flag it. This is the assertion that makes the feature real.
        var definitions = doc.MainDocumentPart.StyleDefinitionsPart;
        Assert.NotNull(definitions);
        var defined = definitions!.Styles!.Descendants<Style>()
            .Select(s => s.StyleId?.Value).ToList();
        Assert.Contains("Heading1", defined);
        Assert.Contains("Heading2", defined);

        // Normal is asserted because every heading style is basedOn it, and a DANGLING basedOn
        // degrades exactly as silently as a missing style: measured, a heading style whose
        // basedOn names an undefined Normal produces ZERO OpenXmlValidator errors. Without this
        // line, dropping the Normal definition while leaving basedOn in place passes every test.
        Assert.Contains("Normal", defined);

        // outlineLvl is what actually drives the navigation pane and TOC depth.
        var heading1 = definitions.Styles.Descendants<Style>()
            .Single(s => s.StyleId?.Value == "Heading1");
        Assert.Equal(0, heading1.StyleParagraphProperties!.OutlineLevel!.Val!.Value);

        // w:pStyle references a PARAGRAPH style. A style with the right id and outline level but
        // w:type="character" satisfies every other assertion here and validates clean - measured,
        // zero OpenXmlValidator errors - while being the wrong kind of style entirely.
        //
        // These four assertions - the style is defined, basedOn resolves, outlineLvl is right, and
        // the type is Paragraph - are deliberately the whole set. They are what make a heading
        // FUNCTION as one. Bold and font size are cosmetic: getting them wrong produces an ugly
        // heading, not a non-heading, so they are not pinned here and a future restyle should not
        // have to edit this test.
        Assert.Equal(StyleValues.Paragraph, heading1.Type!.Value);
    }

    [Fact]
    public void Create_DoesNotDefineHeadingStylesItDoesNotUse()
    {
        var docx = DocxEditor.Create(new[] { DocxBlock.Heading("Only one", 1) });

        AssertValid(docx);

        using var ms = new MemoryStream(docx);
        using var doc = WordprocessingDocument.Open(ms, false);
        var defined = doc.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Descendants<Style>()
            .Select(s => s.StyleId?.Value).ToList();

        Assert.Contains("Heading1", defined);
        Assert.DoesNotContain("Heading4", defined);
    }

    /// <summary>
    /// The other half of the "only what is used" rule: a document with no headings gets no styles
    /// part at all. Asserted rather than assumed because dropping the early return in
    /// AddHeadingStyles would give every document an unused styles part, and nothing else in this
    /// file would notice — the documents would stay schema-valid and their text unchanged.
    /// </summary>
    [Fact]
    public void Create_WithNoHeadings_AddsNoStylesPart()
    {
        var docx = DocxEditor.Create(new[] { DocxBlock.Paragraph("Body only.") });

        AssertValid(docx);

        using var ms = new MemoryStream(docx);
        using var doc = WordprocessingDocument.Open(ms, false);
        Assert.Null(doc.MainDocumentPart!.StyleDefinitionsPart);
    }
}
