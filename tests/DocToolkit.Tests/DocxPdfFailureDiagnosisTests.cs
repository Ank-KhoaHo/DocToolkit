using System.IO.Compression;
using System.Text;

namespace DocToolkit.Tests;

/// <summary>
/// The reasons a real Word document will not render to PDF that a reader can act on.
///
/// <b>Measured 2026-08-18 over 99 documents carrying real content</b> — the corpus's 111 genuine
/// <c>.doc</c> files converted first, since govdocs1 predates <c>.docx</c>. <b>DOCX → PDF succeeded
/// on 71.7%</b> (OfficeIMO 3.2.6), and 15 of the 28 failures were a negative paragraph indent (8) or
/// header or footer content wider than the page (7). <b>3.5 renders a negative indent</b>, so the
/// first is no longer a refusal and its message and retry were removed (A123).
///
/// <b>Both were legal in Word, which is the thing worth saying.</b> The renderer's own messages are
/// accurate and leave a reader hunting for a mistake in a document that does not contain one.
///
/// <b>The header message is not repaired.</b> Shrinking a header changes a layout somebody chose.
/// Unlike the HTML repairs, no browser behaviour says what the right answer is — that is a
/// decision, and it is filed rather than taken.
///
/// <b>A third message was added later, and it is the one with a remedy.</b> A character no font on
/// the machine can encode now says so, says that this DEPENDS ON THE MACHINE — the same document
/// renders on a host whose fonts cover the script, so it can pass in development and fail in
/// production — and points at <see cref="PdfFontOptions"/>, which did not exist when the first two
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

    // ---- the negative indent: 3.5 renders it, with no retry --------------------------------------

    [Theory]
    [InlineData("w:left=\"-360\"")]
    [InlineData("w:right=\"-360\"")]
    [InlineData("w:right=\"-7\"")]                       // tiny values were refused too
    [InlineData("w:left=\"-360\" w:right=\"-360\"")]
    public void ANegativeIndentNowRenders(string attributes)
    {
        // These asserted the DIAGNOSIS, then a clamp-and-retry made them render on 3.2.6 (71/99 to
        // 75/99). OfficeIMO 3.5 renders them as they stand, so the retry and the message are gone
        // and the assertion is the same: it did not fail.
        Assert.True(PdfProbe.IsPdf(DocxToPdfConverter.Convert(WithIndent(attributes))));
    }

    [Fact]
    public void ANegativeIndentKeepsTheDocumentsText()
    {
        // A render that dropped the paragraph would satisfy the theory above.
        var pdf = DocxToPdfConverter.Convert(WithIndent("w:right=\"-720\""));

        Assert.Contains("second line", PdfProbe.ExtractText(pdf), StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("w:left=\"720\"")]
    [InlineData("w:right=\"720\"")]
    [InlineData("w:left=\"720\" w:hanging=\"360\"")]      // an ordinary hanging indent
    [InlineData("w:left=\"720\" w:firstLine=\"-360\"")]   // the other ordinary form
    public void OrdinaryIndentsStillConvert(string attributes)
    {
        // A hanging indent is the commonest indent in real documents; if an ordinary indent ever
        // started failing, every other test here would be explaining the wrong cause.
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
    // The old negative-indent message: 3.5 never says it, and the retry that answered it is gone.
    [InlineData("Paragraph right indent must be a non-negative finite value.")]
    [InlineData("")]
    // A message that mentions an indent for a DIFFERENT reason. Without these, matching the whole
    // word "indent" was indistinguishable from matching the specific phrase - mutation testing found
    // that, and it matters because a diagnosis that names the wrong cause would be actively wrong
    // advice for any of them.
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
