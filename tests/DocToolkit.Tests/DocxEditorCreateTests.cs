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
}
