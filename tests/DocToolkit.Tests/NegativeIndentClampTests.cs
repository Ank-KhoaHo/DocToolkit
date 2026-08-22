using System.IO.Compression;
using System.Text;

namespace DocToolkit.Tests;

/// <summary>
/// Clamping a negative paragraph indent so the document renders.
///
/// <b>Measured over 99 documents carrying real content: DOCX → PDF went from 71/99 to 75/99.</b>
/// Four documents, which is a small return and was expected to be smaller still — the prediction
/// from an earlier probe was two.
///
/// <b>This is the one repair here that takes a liberty.</b> A negative indent is legal in Word,
/// which honours it: content is deliberately set outside the margin in a letterhead or a pull-quote,
/// and clamping pulls it back inside. No reference renderer says that is right — unlike the HTML
/// repairs, where a browser's own behaviour decided the question. It shipped on the maintainer's
/// decision that a document rendering slightly differently beats one not rendering.
/// </summary>
public class NegativeIndentClampTests
{
    private static byte[] WithIndent(string attributes)
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
            xml = xml[..afterOpen] + $"<w:pPr><w:ind {attributes}/></w:pPr>" + xml[afterOpen..];

            entry.Delete();
            var fresh = zip.CreateEntry("word/document.xml", CompressionLevel.Optimal);
            using var writer = new StreamWriter(fresh.Open(), new UTF8Encoding(false));
            writer.Write(xml);
        }

        return ms.ToArray();
    }

    private static string DocumentXml(byte[] docx)
    {
        using var ms = new MemoryStream(docx);
        using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
        using var reader = new StreamReader(zip.GetEntry("word/document.xml")!.Open(), Encoding.UTF8);
        return reader.ReadToEnd();
    }

    // ---- it changes nothing it does not have to ----------------------------------------------------

    [Theory]
    [InlineData("w:left=\"720\"")]
    [InlineData("w:right=\"720\"")]
    [InlineData("w:left=\"720\" w:hanging=\"360\"")]
    [InlineData("w:left=\"720\" w:firstLine=\"-360\"")]
    public void APackageWithNothingToClampComesBackByReference(string attributes)
    {
        // ReferenceEquals: a document that renders today is never repackaged, so it cannot change.
        // A negative firstLine and a hanging indent are the ORDINARY forms and are not refused.
        var docx = WithIndent(attributes);

        Assert.Same(docx, NegativeIndentClamp.Apply(docx));
    }

    [Fact]
    public void ADocumentWithNoIndentAtAllIsUntouched()
    {
        var docx = DocxEditor.Create([DocxBlock.Paragraph("No indent here.")]);

        Assert.Same(docx, NegativeIndentClamp.Apply(docx));
    }

    // ---- what it clamps, and only that -------------------------------------------------------------

    [Theory]
    [InlineData("w:left=\"-360\"")]
    [InlineData("w:right=\"-360\"")]
    [InlineData("w:right=\"-7\"")]
    public void ANegativeIndentBecomesZero(string attributes)
    {
        var clamped = NegativeIndentClamp.Apply(WithIndent(attributes));

        Assert.DoesNotContain("=\"-", DocumentXml(clamped), StringComparison.Ordinal);
        Assert.Contains("w:ind", DocumentXml(clamped), StringComparison.Ordinal);
    }

    [Fact]
    public void AHangingIndentSurvivesAlongsideAClampedOne()
    {
        // The mixed case, and the one most likely to be broken by a careless pattern: w:hanging is
        // POSITIVE and w:firstLine may be negative, and neither is what the renderer refuses.
        var clamped = NegativeIndentClamp.Apply(WithIndent("w:left=\"-360\" w:hanging=\"360\" w:firstLine=\"-180\""));
        var xml = DocumentXml(clamped);

        Assert.Contains("w:hanging=\"360\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:firstLine=\"-180\"", xml, StringComparison.Ordinal);
        Assert.Contains("w:left=\"0\"", xml, StringComparison.Ordinal);
    }

    [Fact]
    public void OnlyParagraphIndentsAreTouched()
    {
        // Those attribute names appear on other elements - table indents, cell margins - and only
        // the paragraph indent is refused. A looser pattern would rewrite parts of the document that
        // were never the problem, which is the difference between a repair and damage.
        var docx = DocxEditor.Create([DocxBlock.Paragraph("Text.")]);

        using var ms = new MemoryStream();
        ms.Write(docx, 0, docx.Length);
        ms.Position = 0;
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = zip.GetEntry("word/document.xml")!;
            string xml;
            using (var r = new StreamReader(entry.Open(), Encoding.UTF8)) xml = r.ReadToEnd();
            xml = xml.Replace("<w:body>", "<w:body><w:tbl><w:tblPr><w:tblInd w:left=\"-500\"/></w:tblPr></w:tbl>");
            entry.Delete();
            var fresh = zip.CreateEntry("word/document.xml", CompressionLevel.Optimal);
            using var w = new StreamWriter(fresh.Open(), new UTF8Encoding(false));
            w.Write(xml);
        }

        var result = NegativeIndentClamp.Apply(ms.ToArray());

        Assert.Contains("w:tblInd w:left=\"-500\"", DocumentXml(result), StringComparison.Ordinal);
    }

    // ---- end to end ---------------------------------------------------------------------------------

    [Fact]
    public void TheDocumentRendersAndKeepsItsText()
    {
        var pdf = DocxToPdfConverter.Convert(WithIndent("w:right=\"-720\""));

        Assert.True(PdfProbe.IsPdf(pdf));
        Assert.Contains("second line", PdfProbe.ExtractText(pdf), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void OrdinaryDocumentsStillRender()
    {
        var docx = DocxEditor.Create([DocxBlock.Heading("TITLE", 1), DocxBlock.Paragraph("BODY")]);

        var text = PdfProbe.ExtractText(DocxToPdfConverter.Convert(docx));
        Assert.Contains("TITLE", text, StringComparison.Ordinal);
        Assert.Contains("BODY", text, StringComparison.Ordinal);
    }

    // ---- ClampOrNull, and the premise the retry guard rests on -------------------------------

    /// <summary>
    /// Nothing to clamp is reported as <c>null</c>, and it is reported about the array it was
    /// GIVEN - not about whatever the caller started with.
    /// </summary>
    /// <remarks>
    /// <b>This pins the premise that the caller got wrong.</b> Until 2026-08-22
    /// <c>DocxToPdfConverter</c> clamped a list-substituted copy of the document and then compared
    /// the result against the ORIGINAL, so for any document containing a list the references could
    /// never match, the "nothing to clamp" guard could not fire, and a doomed second render was
    /// paid anyway.
    ///
    /// <para>Both assertions below are that premise: substitution really does hand back a
    /// different object, and the clamp really does hand back the same one. Neither is obvious, and
    /// the bug lived in the gap between them.</para>
    /// </remarks>
    [Fact]
    public void ClampOrNull_NothingToClamp_IsNull_EvenWhenTheInputIsNotTheOriginalDocument()
    {
        // The real fixture, not a stand-in: substitution only acts on a package that has a
        // numbering part, which DocxEditor.Create does not produce.
        var original = File.ReadAllBytes(Path.Join("assets", "word-bullets.docx"));
        var prepared = ListMarkerSubstitution.Apply(original);

        // The premise that broke the old comparison: substitution returns a NEW array.
        Assert.NotSame(original, prepared);

        // And the premise the guard needs: nothing to clamp means null, about `prepared`.
        Assert.Null(NegativeIndentClamp.ClampOrNull(prepared));
    }

    [Fact]
    public void ClampOrNull_SomethingToClamp_ReturnsTheClampedCopy()
    {
        // The positive control. Without it, a ClampOrNull that returned null unconditionally would
        // pass the test above and look correct.
        var docx = WithIndent("w:left=\"-360\"");

        var clamped = NegativeIndentClamp.ClampOrNull(docx);

        Assert.NotNull(clamped);
        Assert.NotSame(docx, clamped);
        Assert.DoesNotContain("-360", DocumentXml(clamped!));
    }
}
