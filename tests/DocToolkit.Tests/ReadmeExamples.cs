using Xunit;
using Xunit.Abstractions;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

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
    private readonly ITestOutputHelper _output;

    public ReadmeExamples(ITestOutputHelper output) => _output = output;

    private static async Task<byte[]> PdfAsync(string heading) =>
        await HtmlToPdfConverter.ConvertAsync($"<h1>{heading}</h1>");

    /// <summary>
    /// The landing page's first code block. Deliberately the smallest thing that shows the shape of
    /// the whole library: one call in, bytes out, no configuration, no disposal, nothing to wire up.
    /// A README that opens with setup code loses the reader before the capability appears.
    /// </summary>
    [Fact]
    public async Task QuickStartExample()
    {
        #region readme-quickstart
        // HTML in, a Word document and a PDF out. No browser, no LibreOffice, nothing to install.
        byte[] docx = await HtmlToDocxConverter.ConvertAsync("<h1>Invoice 2026-114</h1>");
        byte[] pdf = await HtmlToPdfConverter.ConvertAsync("<h1>Invoice 2026-114</h1>");

        // Read a document back, edit one you were given, and lock it when you are done.
        string text = DocxEditor.ExtractText(docx);
        byte[] locked = PdfEditor.Protect(pdf, new PdfProtection { UserPassword = "s3cret" });
        #endregion

        Assert.Equal("Invoice 2026-114", text);
        Assert.True(PdfProbe.IsPdf(pdf));
        // The control that makes "locked" mean something: it can no longer be opened.
        Assert.Throws<DocumentConversionException>(() => PdfEditor.PageCount(locked));
    }

    [Fact]
    public void PresentationReplaceImageExample()
    {
        byte[] pptx = PptxFixtures.DeckWithPlaceholderBox("{{chart}}");
        var chartPath = Path.Join(Directory.GetCurrentDirectory(), "chart.png");
        File.WriteAllBytes(chartPath, ImageFixtures.Png(40, 30));

        try
        {
            #region readme-pptx-replace-image
            byte[] filled = PresentationEditor.ReplaceImage(pptx, "{{chart}}", File.ReadAllBytes("chart.png"));
            #endregion

            Assert.NotEmpty(filled);
            Assert.DoesNotContain("{{chart}}", string.Join(" ", PresentationEditor.ExtractText(filled)));
        }
        finally
        {
            File.Delete(chartPath);
        }
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
    public void DocumentMetadataExample()
    {
        byte[] docx = DocxEditor.Create(new[] { DocxBlock.Paragraph("Body") });

        #region readme-document-metadata
        byte[] stamped = DocxEditor.WithMetadata(docx, new DocumentMetadata
        {
            Title = "Q3 board report",
            Creator = "Contoso Ltd",
        });

        DocumentMetadata info = DocxEditor.ReadMetadata(stamped);
        #endregion

        Assert.Equal("Q3 board report", info.Title);
        Assert.Equal("Contoso Ltd", info.Creator);
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
        byte[] pdf = await HtmlToPdfConverter.ConvertAsync(html, PageSetup.Letter);
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
    public async Task PdfProtectionExample()
    {
        byte[] statement = await PdfAsync("Statement");

        #region readme-pdf-protection
        // A password to OPEN the document. Without it the file cannot be read at all.
        byte[] locked = PdfEditor.Protect(statement, new PdfProtection
        {
            UserPassword = "s3cret",
            AllowPrinting = false,
        });

        // An OWNER password leaves the document readable and asks readers to honour the
        // restrictions. It is not a lock - use UserPassword when content must not be read.
        byte[] restricted = PdfEditor.Protect(statement, new PdfProtection
        {
            OwnerPassword = "admin",
            AllowCopying = false,
        });

        // The other PdfEditor operations refuse an encrypted document, so unprotect it first.
        // If the document has an owner password, that is the one required here.
        byte[] opened = PdfEditor.Unprotect(locked, "s3cret");
        #endregion

        // The assertion that matters: the locked copy really is locked.
        Assert.Throws<DocumentConversionException>(() => PdfEditor.PageCount(locked));

        // ...and the control, so the line above cannot pass against a converter that breaks
        // everything: the owner-password copy is still readable, and the unprotected one works.
        Assert.Equal(1, PdfEditor.PageCount(restricted));
        Assert.Equal(PdfEditor.PageCount(statement), PdfEditor.PageCount(opened));
        Assert.Contains("Statement", PdfProbe.ExtractText(opened), StringComparison.Ordinal);
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

    [Fact]
    public async Task RemoteImagesExample()
    {
        using var probe = new LoopbackProbe(_output);
        string html = $"""<p><img src="{probe.BaseUrl}/logo.bmp" alt="logo" /></p>""";
        var ct = CancellationToken.None;

        // Gated like every other opt-in-download test in this suite: HtmlToOpenXml's ParseBody is
        // not proven safe to run concurrently with itself once RemoteImageOptions routes through it.
        await RemoteDownloadGate.RunAsync(async () =>
        {
            #region readme-remote-images
            // The ONLY API family that makes an outbound request: downloads and embeds images the markup
            // names. It still succeeds in an air-gapped environment - a host that will not answer just leaves
            // that image out of the result, after a per-image timeout, rather than failing the conversion.
            byte[] docx = await HtmlToDocxConverter.ConvertAsync(html, allowRemoteImageDownload: true, ct);
            byte[] pdf = await HtmlToPdfConverter.ConvertAsync(html, allowRemoteImageDownload: true, ct);

            // RemoteImageOptions bounds that opt-in instead of leaving it wide open. Every default here is
            // already the restrictive one, so `new RemoteImageOptions()` is far narrower than the bool form.
            byte[] bounded = await HtmlToDocxConverter.ConvertAsync(html, new RemoteImageOptions(), ct);
            #endregion

            Assert.NotEmpty(docx);
            Assert.NotEmpty(PdfProbe.MediaBoxes(pdf));
            Assert.NotEmpty(bounded);

            // AllowPrivateAddresses defaults to false, so even with the opt-in requested the loopback
            // probe's image is left out rather than embedded - this is the "still succeeds" claim above.
            Assert.Equal(0, DocxFixtures.Read(docx, main => main.ImageParts.Count()));
            Assert.Equal(0, DocxFixtures.Read(bounded, main => main.ImageParts.Count()));
        });

        await probe.AssertSilentAsync(nameof(RemoteImagesExample));
    }

    [Fact]
    public async Task FillRowsExample()
    {
        byte[] docx = await HtmlToDocxConverter.ConvertAsync(
            """
            <p>Customer: {{customer}}</p>
            <table border="1">
              <tr><th>Description</th><th>Qty</th><th>Total</th></tr>
              <tr><td>{{item.Desc}}</td><td>{{item.Qty}}</td><td>{{item.Total}}</td></tr>
            </table>
            """);

        #region readme-fill-rows
        byte[] filled = DocxEditor.FillRows(docx, "item", new[]
        {
            new Dictionary<string, string> { ["Desc"] = "Widget", ["Qty"] = "2", ["Total"] = "19.98" },
            new Dictionary<string, string> { ["Desc"] = "Gadget", ["Qty"] = "5", ["Total"] = "45.00" },
        });

        // then the document-level scalars
        filled = DocxEditor.ReplaceText(filled, new Dictionary<string, string> { ["{{customer}}"] = "Contoso Ltd" });
        #endregion

        var text = DocxEditor.ExtractText(filled);
        Assert.Contains("Customer: Contoso Ltd", text, StringComparison.Ordinal);
        Assert.DoesNotContain("{{customer}}", text, StringComparison.Ordinal);
        Assert.DoesNotContain("{{item.", text, StringComparison.Ordinal);

        var table = DocxEditor.ReadTable(filled, 0);
        Assert.Equal(new[] { "Description", "Qty", "Total" }, table[0]);
        Assert.Equal(new[] { "Widget", "2", "19.98" }, table[1]);
        Assert.Equal(new[] { "Gadget", "5", "45.00" }, table[2]);
    }

    [Fact]
    public async Task ReadTableExample()
    {
        byte[] docx = await HtmlToDocxConverter.ConvertAsync(
            """
            <table border="1">
              <tr><th>Description</th><th>Qty</th><th>Total</th></tr>
              <tr><td>{{item.Desc}}</td><td>{{item.Qty}}</td><td>{{item.Total}}</td></tr>
            </table>
            """);

        byte[] filled = DocxEditor.FillRows(docx, "item", new[]
        {
            new Dictionary<string, string> { ["Desc"] = "Widget", ["Qty"] = "2", ["Total"] = "19.98" },
            new Dictionary<string, string> { ["Desc"] = "Gadget", ["Qty"] = "5", ["Total"] = "45.00" },
        });

        #region readme-read-table
        int tables = DocxEditor.TableCount(filled);
        IReadOnlyList<IReadOnlyList<string>> rows = DocxEditor.ReadTable(filled, 0);
        // rows[0] is the header row: ["Description", "Qty", "Total"]
        // rows[1] is: ["Widget", "2", "19.98"]
        #endregion

        Assert.Equal(1, tables);
        Assert.Equal(new[] { "Description", "Qty", "Total" }, rows[0]);
        Assert.Equal(new[] { "Widget", "2", "19.98" }, rows[1]);
        Assert.Equal(new[] { "Gadget", "5", "45.00" }, rows[2]);
    }

    [Fact]
    public async Task RemoteOptInExample()
    {
        using var probe = new LoopbackProbe(_output);
        string html = $"""<p><img src="{probe.BaseUrl}/logo.bmp" alt="logo" /></p>""";

        // Gated like every other opt-in-download test in this suite: HtmlToOpenXml's ParseBody is
        // not proven safe to run concurrently with itself once RemoteImageOptions routes through it.
        await RemoteDownloadGate.RunAsync(async () =>
        {
            #region readme-remote-opt-in
            byte[] docx = await HtmlToDocxConverter.ConvertAsync(html, allowRemoteImageDownload: true);

            // Bounded instead of wide open: timeout, byte cap, host allow-list and a block on
            // loopback/private/link-local addresses, all on by default. Not a complete SSRF defence — see
            // the package README.
            byte[] bounded = await HtmlToDocxConverter.ConvertAsync(html, new RemoteImageOptions());
            #endregion

            // AllowPrivateAddresses defaults to false, so even with the opt-in requested the loopback
            // probe's image is left out rather than embedded - the conversion still succeeds either way.
            Assert.Equal(0, DocxFixtures.Read(docx, main => main.ImageParts.Count()));
            Assert.Equal(0, DocxFixtures.Read(bounded, main => main.ImageParts.Count()));
        });

        await probe.AssertSilentAsync(nameof(RemoteOptInExample));
    }

    [Fact]
    public async Task PageSetupOptionsExample()
    {
        string html = "<h1>Invoice</h1>";
        var blocks = new[] { DocxBlock.Paragraph("Body") };

        #region readme-page-setup-options
        byte[] pdf = await HtmlToPdfConverter.ConvertAsync(
            html,
            PageSetup.A4.Landscape().WithMargins(36));

        byte[] docx = DocxEditor.Create(blocks, PageSetup.Letter);
        #endregion

        var pdfBoxes = PdfProbe.MediaBoxes(pdf);
        Assert.NotEmpty(pdfBoxes);
        Assert.All(pdfBoxes, box =>
        {
            // A4 landscape swaps the dimensions: 841.9 x 595.3pt, rounded to the nearest point.
            Assert.Equal(842, box.Width, 0);
            Assert.Equal(595, box.Height, 0);
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
    public void ReplaceImageExample()
    {
        var logoPath = Path.Join(Directory.GetCurrentDirectory(), "logo.png");
        File.WriteAllBytes(logoPath, ImageFixtures.Png(width: 4, height: 3));
        byte[] sigBytes = ImageFixtures.Png(width: 6, height: 2);

        byte[] docx = DocxFixtures.Build(
            DocxFixtures.P(DocxFixtures.R("Logo: {{logo}} end")),
            DocxFixtures.P(DocxFixtures.R("Signed: {{signature}} (authorised)")));

        try
        {
            #region readme-replace-image
            byte[] withLogo = DocxEditor.ReplaceImage(docx, "{{logo}}", File.ReadAllBytes("logo.png"));

            // or at a chosen width; the height scales to keep the aspect ratio
            byte[] signed = DocxEditor.ReplaceImage(withLogo, "{{signature}}", sigBytes, widthPoints: 90);
            #endregion

            var text = DocxEditor.ExtractText(signed);
            Assert.Contains("Logo: ", text, StringComparison.Ordinal);
            Assert.Contains(" end", text, StringComparison.Ordinal);
            Assert.Contains("Signed: ", text, StringComparison.Ordinal);
            Assert.Contains(" (authorised)", text, StringComparison.Ordinal);
            Assert.DoesNotContain("{{logo}}", text, StringComparison.Ordinal);
            Assert.DoesNotContain("{{signature}}", text, StringComparison.Ordinal);

            var extents = DocxFixtures.Read(signed, main => main.Document!.Body!.Descendants<DW.Extent>().ToList());
            Assert.Equal(2, extents.Count);

            // widthPoints: 90 against a 6x2 source -> 90pt wide, 30pt tall to keep the 3:1 ratio.
            Assert.Equal(90L * 12700, extents[1].Cx!.Value);
            Assert.Equal(30L * 12700, extents[1].Cy!.Value);
        }
        finally
        {
            File.Delete(logoPath);
        }
    }
}
