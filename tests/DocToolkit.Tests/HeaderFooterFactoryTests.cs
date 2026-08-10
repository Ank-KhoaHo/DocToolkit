using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace DocToolkit.Tests;

public class HeaderFooterFactoryTests
{
    private static (WordprocessingDocument Doc, MainDocumentPart Main, MemoryStream Stream) NewDocument()
    {
        var ms = new MemoryStream();
        var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document);
        var main = doc.AddMainDocumentPart();
        main.Document = new Document(new Body());
        return (doc, main, ms);
    }

    [Fact]
    public void NoHeaderMeansNoPartsAndNoReferences()
    {
        var (doc, main, _) = NewDocument();
        using (doc)
        {
            var references = HeaderFooterFactory.CreateReferences(main, PageSetup.A4);

            Assert.Empty(references);
            Assert.Empty(main.HeaderParts);
            Assert.Empty(main.FooterParts);
        }
    }

    [Fact]
    public void AHeaderProducesOnePartAndOneDefaultReference()
    {
        var (doc, main, _) = NewDocument();
        using (doc)
        {
            var page = PageSetup.A4.WithHeader(DocxHeader.Text("Contoso"));

            var references = HeaderFooterFactory.CreateReferences(main, page);

            var reference = Assert.IsType<HeaderReference>(Assert.Single(references));
            Assert.Equal(HeaderFooterValues.Default, reference.Type!.Value);
            var part = Assert.Single(main.HeaderParts);
            Assert.Contains("Contoso", part.Header!.InnerText, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void APageNumberBecomesARealField()
    {
        var (doc, main, _) = NewDocument();
        using (doc)
        {
            var page = PageSetup.A4.WithFooter(DocxHeader.Of(
                HeaderAlignment.Right,
                DocxHeaderSegment.Text("Page "),
                DocxHeaderSegment.PageNumber,
                DocxHeaderSegment.Text(" of "),
                DocxHeaderSegment.PageCount));

            HeaderFooterFactory.CreateReferences(main, page);

            var footer = Assert.Single(main.FooterParts).Footer!;
            var fields = footer.Descendants<SimpleField>().ToList();

            Assert.Equal(2, fields.Count);
            Assert.Equal(" PAGE ", fields[0].Instruction!.Value);
            Assert.Equal(" NUMPAGES ", fields[1].Instruction!.Value);
        }
    }

    [Fact]
    public void AlignmentBecomesAJustificationOnTheParagraph()
    {
        var (doc, main, _) = NewDocument();
        using (doc)
        {
            var page = PageSetup.A4.WithHeader(DocxHeader.Text("mid", HeaderAlignment.Center));

            HeaderFooterFactory.CreateReferences(main, page);

            var paragraph = Assert.Single(main.HeaderParts).Header!.Descendants<Paragraph>().Single();
            Assert.Equal(
                JustificationValues.Center,
                paragraph.ParagraphProperties!.Justification!.Val!.Value);
        }
    }

    // InnerText reads Text.Text directly and cannot see the Space attribute, so an assertion on
    // text alone would pass even if the trailing space were lost. "Page " + a page-number field
    // renders as the literal string "Page3" without SpaceProcessingModeValues.Preserve on the
    // w:t element, and only the serialized XML can catch that regression.
    [Fact]
    public void ALiteralSegmentEndingInASpacePreservesItInTheSerializedXml()
    {
        var (doc, main, _) = NewDocument();
        using (doc)
        {
            var page = PageSetup.A4.WithHeader(DocxHeader.Of(
                HeaderAlignment.Left,
                DocxHeaderSegment.Text("Page "),
                DocxHeaderSegment.PageNumber));

            HeaderFooterFactory.CreateReferences(main, page);

            var part = Assert.Single(main.HeaderParts);
            Assert.Contains("xml:space=\"preserve\"", part.Header!.OuterXml, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void AFirstPageFooterAloneProducesFirstAndDefaultReferencesButNoFirstHeader()
    {
        var (doc, main, _) = NewDocument();
        using (doc)
        {
            var page = PageSetup.A4
                .WithHeader(DocxHeader.Text("running"))
                .WithFirstPage(header: null, footer: DocxHeader.Text("only on page one"));

            var references = HeaderFooterFactory.CreateReferences(main, page);

            Assert.Single(references.OfType<HeaderReference>());
            var footerReference = Assert.Single(references.OfType<FooterReference>());
            Assert.Equal(HeaderFooterValues.First, footerReference.Type!.Value);
            Assert.Single(main.HeaderParts);
            Assert.Single(main.FooterParts);
        }
    }
}
