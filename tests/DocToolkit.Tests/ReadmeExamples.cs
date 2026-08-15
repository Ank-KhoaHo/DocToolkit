using Xunit;

namespace DocToolkit.Tests;

/// <summary>
/// The README code blocks, as tests. They are injected into the three READMEs by
/// scripts/gen-readme-snippets.py, so a block that stops compiling breaks the build and a
/// block that compiles but is WRONG fails its assertion.
///
/// Separate from DocumentationExamples.cs deliberately: that file feeds the DocFX guides
/// through &lt;code source&gt;, this one feeds markdown through a generator. One file serving
/// two inclusion mechanisms is how a change for one silently reshapes the other.
///
/// Setup ABOVE the region, assertions BELOW it - the reader sees only the capability.
/// </summary>
public class ReadmeExamples
{
    private static async Task<byte[]> PdfAsync(string heading) =>
        await HtmlToPdfConverter.ConvertAsync($"<h1>{heading}</h1>");

    [Fact]
    public void PresentationReplaceImageExample()
    {
        byte[] pptx = PptxFixtures.DeckWithPlaceholderBox("{{chart}}");
        byte[] chartPngBytes = ImageFixtures.Png(40, 30);

        #region readme-pptx-replace-image
        byte[] filled = PresentationEditor.ReplaceImage(pptx, "{{chart}}", chartPngBytes);
        #endregion

        Assert.NotEmpty(filled);
        Assert.DoesNotContain("{{chart}}", string.Join(" ", PresentationEditor.ExtractText(filled)));
    }

    [Fact]
    public async Task PdfMetadataExample()
    {
        byte[] bundle = await PdfAsync("Invoice");

        #region readme-pdf-metadata
        byte[] stamped = PdfEditor.WithMetadata(bundle, new PdfMetadata
        {
            Title = "Invoice INV-2026-0042",
            Author = "Contoso Ltd",
        });

        PdfMetadata info = PdfEditor.ReadMetadata(stamped);
        #endregion

        Assert.Equal("Invoice INV-2026-0042", info.Title);
        Assert.Equal("Contoso Ltd", info.Author);
    }

    [Fact]
    public void PageSetupExample()
    {
        var blocks = new[] { DocxBlock.Paragraph("Body") };

        #region readme-page-setup
        var page = PageSetup.A4
            .WithHeader(DocxHeader.Text("Contoso Ltd"))
            .WithFooter(DocxHeader.Of(HeaderAlignment.Right,
                DocxHeaderSegment.Text("Page "), DocxHeaderSegment.PageNumber,
                DocxHeaderSegment.Text(" of "), DocxHeaderSegment.PageCount));

        byte[] docx = DocxEditor.Create(blocks, page);
        #endregion

        Assert.Contains("Contoso Ltd", DocxEditor.ExtractText(docx, includeHeadersAndFooters: true),
            StringComparison.Ordinal);
    }

    [Fact]
    public async Task HtmlToPdfPageSetupExample()
    {
        string html = "<h1>Invoice</h1>";
        var blocks = new[] { DocxBlock.Paragraph("Body") };

        #region readme-html-to-pdf-page
        byte[] pdf  = await HtmlToPdfConverter.ConvertAsync(html, PageSetup.Letter);
        byte[] docx = DocxEditor.Create(blocks, PageSetup.Letter);
        #endregion

        var pdfBoxes = PdfProbe.MediaBoxes(pdf);
        Assert.NotEmpty(pdfBoxes);
        Assert.All(pdfBoxes, box =>
        {
            Assert.Equal(612, box.Width, 0);
            Assert.Equal(792, box.Height, 0);
        });

        var docxBoxes = PdfProbe.MediaBoxes(DocxToPdfConverter.Convert(docx));
        Assert.NotEmpty(docxBoxes);
        Assert.All(docxBoxes, box =>
        {
            Assert.Equal(612, box.Width, 0);
            Assert.Equal(792, box.Height, 0);
        });
    }

    [Fact]
    public async Task PdfUtilitiesExample()
    {
        byte[] pdf = PdfEditor.Merge([await PdfAsync("Alpha"), await PdfAsync("Bravo")]);
        byte[] cover = await PdfAsync("Cover");
        byte[] invoice = await PdfAsync("Invoice");
        byte[] terms = await PdfAsync("Terms");
        byte[] appendix = await PdfAsync("Appendix");

        #region readme-pdf-utilities
        int pages = PdfEditor.PageCount(pdf);

        // Join several into one, in the order given.
        byte[] bundle = PdfEditor.Merge([cover, invoice, terms]);

        // And take a range back out. firstPage is 1-based, the way a reader numbers pages.
        byte[] justTheInvoice = PdfEditor.ExtractPages(bundle, firstPage: 2, count: 1);

        // Or keep everything except a range - the complement of ExtractPages.
        byte[] withoutTheCover = PdfEditor.RemovePages(bundle, firstPage: 1, count: 1);

        // Turn a page that came out sideways. Relative, so calling it twice leaves you at 180.
        byte[] upright = PdfEditor.RotatePages(bundle, firstPage: 3, count: 1, degrees: 90);

        // Put the pages in a different order - a permutation of every page, not a subset.
        byte[] resequenced = PdfEditor.ReorderPages(bundle, [3, 1, 2]);

        // Slot another document in. atPage is where its first page lands; PageCount + 1 appends.
        byte[] withAppendix = PdfEditor.InsertPages(bundle, appendix, atPage: 2);

        // Read a PDF's text back out, one string per page — pageText[0] is page 1.
        IReadOnlyList<string> pageText = PdfEditor.ExtractText(bundle);
        #endregion

        Assert.Equal(2, pages);

        Assert.Equal(3, PdfEditor.PageCount(bundle));
        Assert.Contains("Cover", PdfProbe.ExtractText(PdfEditor.ExtractPages(bundle, 1, 1)), StringComparison.Ordinal);
        Assert.Contains("Invoice", PdfProbe.ExtractText(PdfEditor.ExtractPages(bundle, 2, 1)), StringComparison.Ordinal);
        Assert.Contains("Terms", PdfProbe.ExtractText(PdfEditor.ExtractPages(bundle, 3, 1)), StringComparison.Ordinal);

        var justTheInvoiceText = PdfProbe.ExtractText(justTheInvoice);
        Assert.Contains("Invoice", justTheInvoiceText, StringComparison.Ordinal);
        Assert.DoesNotContain("Cover", justTheInvoiceText, StringComparison.Ordinal);
        Assert.DoesNotContain("Terms", justTheInvoiceText, StringComparison.Ordinal);

        var withoutTheCoverText = PdfProbe.ExtractText(withoutTheCover);
        Assert.DoesNotContain("Cover", withoutTheCoverText, StringComparison.Ordinal);
        Assert.Contains("Invoice", withoutTheCoverText, StringComparison.Ordinal);
        Assert.Contains("Terms", withoutTheCoverText, StringComparison.Ordinal);

        Assert.Equal([90], PdfProbe.PageRotations(upright));

        Assert.Contains("Terms", PdfProbe.ExtractText(PdfEditor.ExtractPages(resequenced, 1, 1)), StringComparison.Ordinal);
        Assert.Contains("Cover", PdfProbe.ExtractText(PdfEditor.ExtractPages(resequenced, 2, 1)), StringComparison.Ordinal);
        Assert.Contains("Invoice", PdfProbe.ExtractText(PdfEditor.ExtractPages(resequenced, 3, 1)), StringComparison.Ordinal);

        Assert.Equal(4, PdfEditor.PageCount(withAppendix));
        Assert.Contains("Appendix", PdfProbe.ExtractText(PdfEditor.ExtractPages(withAppendix, 2, 1)), StringComparison.Ordinal);

        Assert.Equal(3, pageText.Count);
        Assert.Contains("Cover", pageText[0], StringComparison.Ordinal);
        Assert.Contains("Invoice", pageText[1], StringComparison.Ordinal);
        Assert.Contains("Terms", pageText[2], StringComparison.Ordinal);
    }
}
