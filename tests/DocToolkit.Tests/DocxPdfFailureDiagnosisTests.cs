using System.IO.Compression;
using System.Text;

namespace DocToolkit.Tests;

/// <summary>
/// The two commonest reasons a real Word document will not render to PDF.
///
/// <b>Measured 2026-08-18 over 99 documents carrying real content</b> — the corpus's 111 genuine
/// <c>.doc</c> files converted first, since govdocs1 predates <c>.docx</c>. <b>DOCX → PDF succeeded
/// on 71.7%</b>, and 15 of the 28 failures are these two: 8 for a negative paragraph indent, 7 for
/// header or footer content wider than the page.
///
/// <b>Both are legal in Word, which is the thing worth saying.</b> The renderer's own messages are
/// accurate and leave a reader hunting for a mistake in a document that does not contain one.
///
/// <b>Neither is repaired.</b> Clamping a negative indent pulls content back inside a margin the
/// author put it outside of; shrinking a header changes a layout somebody chose. Unlike the HTML
/// repairs, no browser behaviour says what the right answer is — these are decisions, and they are
/// filed rather than taken.
///
/// <b>A third message was added later, and it is the one with a remedy.</b> A character no font on
/// the machine can encode now says so, says that this DEPENDS ON THE MACHINE — the same document
/// renders on a host whose fonts cover the script, so it can pass in development and fail in
/// production — and points at <see cref="PdfFontOptions"/>, which did not exist when the other two
/// were written.
/// </summary>
public class DocxPdfFailureDiagnosisTests
{
    /// <summary>Builds a one-paragraph document carrying the given <c>w:ind</c> attributes.</summary>
    private static byte[] WithIndent(string indentAttributes)
    {
        var docx = DocxEditor.Create([DocxBlock.Paragraph("Some text long enough to wrap onto a second line.")]);

        using var ms = new MemoryStream();
        ms.Write(docx, 0, docx.Length);
        ms.Position = 0;

        using (var zip = new ZipArchive(ms, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = zip.GetEntry("word/document.xml")!;
            string xml;
            using (var reader = new StreamReader(entry.Open(), Encoding.UTF8)) xml = reader.ReadToEnd();

            var paragraph = xml.IndexOf("<w:p ", StringComparison.Ordinal);
            if (paragraph < 0) paragraph = xml.IndexOf("<w:p>", StringComparison.Ordinal);
            var afterOpen = xml.IndexOf('>', paragraph) + 1;
            xml = xml[..afterOpen] + $"<w:pPr><w:ind {indentAttributes}/></w:pPr>" + xml[afterOpen..];

            entry.Delete();
            var fresh = zip.CreateEntry("word/document.xml", CompressionLevel.Optimal);
            using var writer = new StreamWriter(fresh.Open(), new UTF8Encoding(false));
            writer.Write(xml);
        }

        return ms.ToArray();
    }

    // ---- the negative indent: now CLAMPED, so these documents render ------------------------------

    [Theory]
    [InlineData("w:left=\"-360\"")]
    [InlineData("w:right=\"-360\"")]
    [InlineData("w:right=\"-7\"")]                       // tiny values were refused too
    [InlineData("w:left=\"-360\" w:right=\"-360\"")]
    public void ANegativeIndentNowRenders(string attributes)
    {
        // These asserted the DIAGNOSIS until the clamp shipped. The message was right and the
        // document still did not render; now it does, so the assertion moves from "says why it
        // failed" to "did not fail". Measured across the corpus: 71/99 to 75/99.
        Assert.True(PdfProbe.IsPdf(DocxToPdfConverter.Convert(WithIndent(attributes))));
    }

    [Fact]
    public void TheClampKeepsTheDocumentsText()
    {
        // A clamp that rendered by dropping the paragraph would satisfy the theory above. What is
        // given up is the overhang past the margin, not the content.
        var pdf = DocxToPdfConverter.Convert(WithIndent("w:right=\"-720\""));

        Assert.Contains("second line", PdfProbe.ExtractText(pdf), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void TheDiagnosisIsKeptAsASafetyNet()
    {
        // Unreachable from ordinary input now, and kept for the same reason as the rowspan one: the
        // clamp rewrites a package, and input it cannot rewrite - or a shape it does not match -
        // should still say what happened rather than surfacing a bare renderer error.
        var described = DocxPdfFailureDiagnosis.Describe(
            new ArgumentException("Paragraph right indent must be a non-negative finite value."));

        Assert.NotNull(described);
        Assert.Contains("not invalid", described, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("w:left=\"720\"")]
    [InlineData("w:right=\"720\"")]
    [InlineData("w:left=\"720\" w:hanging=\"360\"")]      // an ordinary hanging indent
    [InlineData("w:left=\"720\" w:firstLine=\"-360\"")]   // the other ordinary form
    public void OrdinaryIndentsStillConvert(string attributes)
    {
        // Measured, and the message claims it - so this is what stops that claim going stale. A
        // hanging indent is the commonest indent in real documents; if it ever started failing, the
        // message would be telling people something false at the worst moment.
        Assert.NotEmpty(DocxToPdfConverter.Convert(WithIndent(attributes)));
    }

    // ---- what must not be claimed --------------------------------------------------------------------

    [Fact]
    public void AnUnrecognisedFailureKeepsTheGenericWrapper()
    {
        var ex = Assert.Throws<DocumentConversionException>(
            () => DocxToPdfConverter.Convert([1, 2, 3, 4]));

        Assert.Contains("See the inner exception", ex.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("indent", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("Paragraph right indent must be a non-negative finite value.", "negative left or right indent")]
    [InlineData("Paragraph left indent must be a non-negative finite value.", "negative left or right indent")]
    [InlineData("PDF footer zone content must fit inside the page content width.", "wider than the page content area")]
    [InlineData("Combined PDF header/footer content must fit inside the page content width.", "wider than the page content area")]
    [InlineData("PDF header zones must not overlap.", "wider than the page content area")]
    [InlineData("Text contains character U+0421 that is not covered by any embedded font fallback candidate.", "depends on the machine")]
    public void EachRecognisedMessageMapsToItsDiagnosis(string rendererMessage, string expected)
    {
        // The renderer's exact wording, taken from the corpus run rather than invented. If it ever
        // changes, these go red instead of the diagnosis silently ceasing to fire.
        var described = DocxPdfFailureDiagnosis.Describe(new ArgumentException(rendererMessage));

        Assert.NotNull(described);
        Assert.Contains(expected, described, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("Something else entirely.")]
    [InlineData("PDF bookmark link target 'appendix' was not found.")]
    [InlineData("Table horizontal cell padding must leave a positive text width.")]
    [InlineData("")]
    // A message that mentions an indent for a DIFFERENT reason. Without these, matching the whole
    // word "indent" was indistinguishable from matching the specific phrase - mutation testing found
    // that, and it matters because the remedy this class prints ("set the negative indent to 0")
    // would be actively wrong advice for any of them.
    [InlineData("Paragraph indent could not be applied to this style.")]
    [InlineData("List indent level is out of range.")]
    [InlineData("Table cell indent exceeds the column width.")]
    public void EveryOtherFailureGetsNoDiagnosis(string rendererMessage)
    {
        // The third case matters too: it is a real refusal from the same renderer that a DIFFERENT
        // repair handles, so a looser match would put the wrong remedy in front of the reader.
        Assert.Null(DocxPdfFailureDiagnosis.Describe(new ArgumentException(rendererMessage)));
    }

    [Fact]
    public void TheMissingGlyphMessagePointsAtTheRemedy()
    {
        // The renderer's own message says "add a fallback font" without saying this package now
        // takes one, so a reader hits a wall the API can already get them past.
        var described = DocxPdfFailureDiagnosis.Describe(
            new ArgumentException("Text contains character U+0421 that is not covered by any embedded font."));

        Assert.NotNull(described);
        Assert.Contains("PdfFontOptions", described, StringComparison.Ordinal);
        Assert.Contains("depends on the machine", described, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void OrdinaryDocumentsStillConvert()
    {
        var docx = DocxEditor.Create([DocxBlock.Heading("Title", 1), DocxBlock.Paragraph("Body.")]);

        Assert.True(PdfProbe.IsPdf(DocxToPdfConverter.Convert(docx)));
    }
}
