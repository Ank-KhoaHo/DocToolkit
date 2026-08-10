using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace DocToolkit.Tests;

/// <summary>
/// What page setup actually lands in the produced file.
///
/// These assert on the <c>w:sectPr</c> element rather than on a round trip, deliberately: a test
/// that writes a document and reads its text back passes identically against a document with no
/// page setup at all, which is exactly the defect this feature exists to fix.
/// </summary>
public class PageSetupOutputTests
{
    private static readonly DocxBlock[] Blocks = { DocxBlock.Paragraph("Hello.") };

    /// <summary>
    /// The <c>w:sectPr</c> of the main document part, or null if there is none. Asserting on null
    /// is what makes "no page setup at all" a visible failure.
    /// </summary>
    private static SectionProperties? SectionPropertiesOf(byte[] docx)
    {
        using var ms = new MemoryStream(docx);
        using var doc = WordprocessingDocument.Open(ms, false);
        return doc.MainDocumentPart!.Document!.Body!.Elements<SectionProperties>().SingleOrDefault();
    }

    [Fact]
    public void Create_WithNoPageSetup_EmitsA4()
    {
        var sectPr = SectionPropertiesOf(DocxEditor.Create(Blocks));

        Assert.NotNull(sectPr);
        var size = sectPr!.GetFirstChild<PageSize>();
        Assert.NotNull(size);
        Assert.Equal(11906U, size!.Width!.Value);
        Assert.Equal(16838U, size.Height!.Value);
    }

    [Fact]
    public void Create_WithLetter_EmitsLetter()
    {
        var sectPr = SectionPropertiesOf(DocxEditor.Create(Blocks, PageSetup.Letter));

        var size = sectPr!.GetFirstChild<PageSize>();
        Assert.Equal(12240U, size!.Width!.Value);
        Assert.Equal(15840U, size.Height!.Value);
    }

    [Fact]
    public void Create_WritesTheMarginsInTwentiethsOfAPoint()
    {
        var page = PageSetup.A4.WithMargins(10, 20, 30, 40);

        var sectPr = SectionPropertiesOf(DocxEditor.Create(Blocks, page));

        var margin = sectPr!.GetFirstChild<PageMargin>();
        Assert.NotNull(margin);
        Assert.Equal(200, margin!.Top!.Value);
        Assert.Equal(400U, margin.Right!.Value);
        Assert.Equal(600, margin.Bottom!.Value);
        Assert.Equal(800U, margin.Left!.Value);
    }

    // Word reads the dimensions, but its page-setup UI and several renderers read w:orient. A
    // landscape page whose orient still says portrait is a document that disagrees with itself.
    [Fact]
    public void Create_WithLandscape_SwapsTheDimensionsAndSaysSo()
    {
        var sectPr = SectionPropertiesOf(DocxEditor.Create(Blocks, PageSetup.A4.Landscape()));

        var size = sectPr!.GetFirstChild<PageSize>();
        Assert.Equal(16838U, size!.Width!.Value);
        Assert.Equal(11906U, size.Height!.Value);
        Assert.Equal(PageOrientationValues.Landscape, size.Orient!.Value);
    }

    [Fact]
    public void Create_WithPortraitPage_SaysPortrait()
    {
        var sectPr = SectionPropertiesOf(DocxEditor.Create(Blocks, PageSetup.A4));

        Assert.Equal(
            PageOrientationValues.Portrait,
            sectPr!.GetFirstChild<PageSize>()!.Orient!.Value);
    }

    // sectPr anywhere but last makes Word declare the file corrupt. Nothing else catches this: the
    // document is schema-valid either way and every text-reading test still passes.
    [Fact]
    public void Create_PutsSectionPropertiesLastInTheBody()
    {
        byte[] docx = DocxEditor.Create(
            new[] { DocxBlock.Heading("Title", 1), DocxBlock.Paragraph("Body.") });

        using var ms = new MemoryStream(docx);
        using var doc = WordprocessingDocument.Open(ms, false);
        var body = doc.MainDocumentPart!.Document!.Body!;

        Assert.IsType<SectionProperties>(body.LastChild);
    }

    [Fact]
    public void Create_WithNoBlocks_StillEmitsPageSetup()
    {
        var sectPr = SectionPropertiesOf(DocxEditor.Create(Array.Empty<DocxBlock>()));

        Assert.NotNull(sectPr);
    }

    [Fact]
    public async Task CreateAsync_HonoursThePageSetup()
    {
        using var destination = new MemoryStream();

        await DocxEditor.CreateAsync(Blocks, PageSetup.Letter, destination);

        var size = SectionPropertiesOf(destination.ToArray())!.GetFirstChild<PageSize>();
        Assert.Equal(12240U, size!.Width!.Value);
    }

    [Fact]
    public async Task CreateToFileAsync_HonoursThePageSetup()
    {
        string path = Path.Join(Path.GetTempPath(), $"{Guid.NewGuid():N}.docx");
        try
        {
            await DocxEditor.CreateToFileAsync(Blocks, PageSetup.Letter, path);

            var size = SectionPropertiesOf(await File.ReadAllBytesAsync(path))!
                .GetFirstChild<PageSize>();
            Assert.Equal(12240U, size!.Width!.Value);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Create_WithNullPageSetup_ThrowsArgumentNullException()
    {
        var ex = Assert.Throws<ArgumentNullException>(() => DocxEditor.Create(Blocks, null!));

        Assert.Equal("page", ex.ParamName);
    }

    // Fractional points are rounded, not truncated: 100.03 pt is 2000.6 twentieths, and truncating
    // would lose a whole twentieth on a value that was only ever an approximation anyway.
    [Fact]
    public void Create_RoundsFractionalPointsToTheNearestTwentieth()
    {
        var page = PageSetup.Custom(100.03, 200);

        var size = SectionPropertiesOf(DocxEditor.Create(Blocks, page))!.GetFirstChild<PageSize>();

        Assert.Equal(2001U, size!.Width!.Value);
    }

    private const string Html = "<h1>Report</h1><p>Body text.</p>";

    [Fact]
    public async Task HtmlToDocx_WithNoPageSetup_EmitsA4()
    {
        var sectPr = SectionPropertiesOf(await HtmlToDocxConverter.ConvertAsync(Html));

        Assert.NotNull(sectPr);
        Assert.Equal(11906U, sectPr!.GetFirstChild<PageSize>()!.Width!.Value);
    }

    [Fact]
    public async Task HtmlToDocx_WithLetter_EmitsLetter()
    {
        var sectPr = SectionPropertiesOf(
            await HtmlToDocxConverter.ConvertAsync(Html, PageSetup.Letter));

        Assert.Equal(12240U, sectPr!.GetFirstChild<PageSize>()!.Width!.Value);
    }

    [Fact]
    public async Task HtmlToDocx_PutsSectionPropertiesLastInTheBody()
    {
        byte[] docx = await HtmlToDocxConverter.ConvertAsync(Html);

        using var ms = new MemoryStream(docx);
        using var doc = WordprocessingDocument.Open(ms, false);

        Assert.IsType<SectionProperties>(doc.MainDocumentPart!.Document!.Body!.LastChild);
    }

    [Fact]
    public async Task HtmlToDocxAsync_ToStream_HonoursThePageSetup()
    {
        using var destination = new MemoryStream();

        await HtmlToDocxConverter.ConvertAsync(Html, PageSetup.Letter, destination);

        Assert.Equal(
            12240U,
            SectionPropertiesOf(destination.ToArray())!.GetFirstChild<PageSize>()!.Width!.Value);
    }

    [Fact]
    public async Task HtmlToDocx_ToFile_HonoursThePageSetup()
    {
        string path = Path.Join(Path.GetTempPath(), $"{Guid.NewGuid():N}.docx");
        try
        {
            await HtmlToDocxConverter.ConvertToFileAsync(Html, PageSetup.Letter, path);

            Assert.Equal(
                12240U,
                SectionPropertiesOf(await File.ReadAllBytesAsync(path))!
                    .GetFirstChild<PageSize>()!.Width!.Value);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public async Task HtmlToDocx_WithNullPageSetup_ThrowsArgumentNullException()
    {
        // Typed rather than a bare null!: three two-argument overloads now take a reference
        // type in that position, so `null!` alone cannot pick one.
        PageSetup nullPage = null!;

        var ex = await Assert.ThrowsAsync<ArgumentNullException>(
            () => HtmlToDocxConverter.ConvertAsync(Html, nullPage));

        Assert.Equal("page", ex.ParamName);
    }

    // The MediaBox is compared at whole-point precision: A4's 595.2756 pt lands as 595.3 in the
    // sectPr and OfficeIMO may write 595 or 595.25 into the PDF. What matters is that it is A4
    // rather than Letter, and 1 pt says exactly that.
    private static void AssertPageSize(byte[] pdf, double expectedWidth, double expectedHeight)
    {
        var boxes = PdfProbe.MediaBoxes(pdf);

        Assert.NotEmpty(boxes);
        Assert.All(boxes, box =>
        {
            Assert.Equal(expectedWidth, box.Width, 0);
            Assert.Equal(expectedHeight, box.Height, 0);
        });
    }

    [Fact]
    public async Task HtmlToPdf_WithNoPageSetup_RendersA4()
    {
        AssertPageSize(await HtmlToPdfConverter.ConvertAsync(Html), 595, 842);
    }

    [Fact]
    public async Task HtmlToPdf_WithLetter_RendersLetter()
    {
        AssertPageSize(await HtmlToPdfConverter.ConvertAsync(Html, PageSetup.Letter), 612, 792);
    }

    [Fact]
    public async Task HtmlToPdf_WithLandscape_RendersTheSwappedDimensions()
    {
        AssertPageSize(
            await HtmlToPdfConverter.ConvertAsync(Html, PageSetup.A4.Landscape()), 842, 595);
    }

    [Fact]
    public async Task HtmlToPdfAsync_ToStream_HonoursThePageSetup()
    {
        using var destination = new MemoryStream();

        await HtmlToPdfConverter.ConvertAsync(Html, PageSetup.Letter, destination);

        AssertPageSize(destination.ToArray(), 612, 792);
    }

    [Fact]
    public async Task HtmlToPdf_ToFile_HonoursThePageSetup()
    {
        string path = Path.Join(Path.GetTempPath(), $"{Guid.NewGuid():N}.pdf");
        try
        {
            await HtmlToPdfConverter.ConvertToFileAsync(Html, PageSetup.Letter, path);

            AssertPageSize(await File.ReadAllBytesAsync(path), 612, 792);
        }
        finally
        {
            File.Delete(path);
        }
    }

    // DocxToPdfConverter takes a document that already states its own paper, so it gets no
    // PageSetup overload. This asserts it honours what it is handed - which is what makes the
    // absent overload correct rather than an omission.
    [Fact]
    public void DocxToPdf_HonoursTheSectionPropertiesItIsHanded()
    {
        byte[] docx = DocxEditor.Create(Blocks, PageSetup.Letter);

        AssertPageSize(DocxToPdfConverter.Convert(docx), 612, 792);
    }

    // A square page. `WidthPoints > HeightPoints` mutated to `>=` flips a square from Portrait to
    // Landscape, and nothing else in the suite has a square page to notice.
    [Fact]
    public void Create_WithASquarePage_SaysPortrait()
    {
        var square = PageSetup.Custom(400, 400);

        var sectPr = SectionPropertiesOf(DocxEditor.Create(Blocks, square));

        Assert.Equal(
            PageOrientationValues.Portrait,
            sectPr!.GetFirstChild<PageSize>()!.Orient!.Value);
    }


    /// <summary>
    /// <c>SectionPropertiesFactory.Build</c> opens with <c>ArgumentNullException.ThrowIfNull</c>.
    /// Mutation deleted that line and nothing failed - no test had ever passed null, because every
    /// caller inside the library passes a PageSetup it just constructed.
    ///
    /// Deleting it is not harmless. Without the guard the next line dereferences the argument and
    /// the caller gets a NullReferenceException from inside a factory they did not know existed,
    /// instead of an exception naming the parameter they got wrong.
    /// </summary>
    [Fact]
    public void BuildingSectionPropertiesFromANullPageIsRejectedByName()
    {
        using var ms = new MemoryStream();
        using var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document);
        var main = doc.AddMainDocumentPart();
        main.Document = new Document(new Body());

        var ex = Assert.Throws<ArgumentNullException>(
            () => SectionPropertiesFactory.Build(main, null!));

        Assert.Equal("page", ex.ParamName);
    }

    /// <summary>
    /// Mirrors <see cref="BuildingSectionPropertiesFromANullPageIsRejectedByName"/> for the other
    /// new parameter: <c>main</c> is checked first, so a null main is rejected by its own name
    /// rather than surfacing as a NullReferenceException once <c>HeaderFooterFactory</c> tries to
    /// use it.
    /// </summary>
    [Fact]
    public void BuildingSectionPropertiesFromANullMainPartIsRejectedByName()
    {
        var ex = Assert.Throws<ArgumentNullException>(
            () => SectionPropertiesFactory.Build(null!, PageSetup.A4));

        Assert.Equal("main", ex.ParamName);
    }


    // =============================================================================================
    // Page setup and remote images together.
    // =============================================================================================
    //
    // The public overloads made these mutually exclusive: (html, page) always converts OFFLINE, and
    // (html, RemoteImageOptions) always lays out on A4 - it passes PageSetup.A4 to the internal
    // builder regardless. The builder has taken both since page setup shipped; nothing exposed the
    // combination.
    //
    // That is a silent drop either way round, which is the shape of failure this library treats as
    // worse than an exception: ask for Letter and a CDN allow-list, and one of the two is discarded
    // without a word.

    [Fact]
    public async Task ConvertAsync_WithPageAndRemoteImageOptions_HonoursThePage()
    {
        var remote = new RemoteImageOptions();
        remote.AllowedHosts.Add("assets.example.invalid");

        var docx = await HtmlToDocxConverter.ConvertAsync("<p>Hi</p>", PageSetup.Letter, remote);

        var size = SectionPropertiesOf(docx)!.GetFirstChild<PageSize>();

        // 12240 twentieths = 612 pt = US Letter. A4 would be 11906.
        Assert.Equal(12240U, size!.Width!.Value);
    }

    [Fact]
    public async Task ConvertAsync_WithPageAndRemoteImageOptions_StreamOverloadHonoursThePage()
    {
        var remote = new RemoteImageOptions();
        remote.AllowedHosts.Add("assets.example.invalid");

        await using var destination = new MemoryStream();
        await HtmlToDocxConverter.ConvertAsync("<p>Hi</p>", PageSetup.Letter, remote, destination);

        var size = SectionPropertiesOf(destination.ToArray())!.GetFirstChild<PageSize>();

        Assert.Equal(12240U, size!.Width!.Value);
    }

    // =============================================================================================
    // Null html, on the overloads that take a page.
    // =============================================================================================
    //
    // Every OTHER overload rejects null html. The page ones did not: they went straight to the
    // internal builder, which validated `page` and not `html`, so ConvertAsync(null, PageSetup.A4)
    // returned an empty document. A caller passing null got a document rather than an error - the
    // silent shape this library treats as worse than throwing.
    //
    // Found by a dependency-injection test, not by these: the DI default-page work routed the
    // parameterless overload through the page one, and the service's existing null-html test
    // stopped passing.

    [Fact]
    public async Task DocxPageOverloads_RejectNullHtml()
    {
        await using var destination = new MemoryStream();

        await Assert.ThrowsAsync<ArgumentNullException>(
            () => HtmlToDocxConverter.ConvertAsync(null!, PageSetup.A4));
        await Assert.ThrowsAsync<ArgumentNullException>(
            () => HtmlToDocxConverter.ConvertAsync(null!, PageSetup.A4, destination));
    }

    [Fact]
    public async Task PdfPageOverloads_RejectNullHtml()
    {
        await using var destination = new MemoryStream();

        await Assert.ThrowsAsync<ArgumentNullException>(
            () => HtmlToPdfConverter.ConvertAsync(null!, PageSetup.A4));
        await Assert.ThrowsAsync<ArgumentNullException>(
            () => HtmlToPdfConverter.ConvertAsync(null!, PageSetup.A4, destination));
    }

    [Fact]
    public async Task PageAndOptionsOverloads_RejectNullHtml()
    {
        await Assert.ThrowsAsync<ArgumentNullException>(
            () => HtmlToDocxConverter.ConvertAsync(null!, PageSetup.A4, new RemoteImageOptions()));
        await Assert.ThrowsAsync<ArgumentNullException>(
            () => HtmlToPdfConverter.ConvertAsync(null!, PageSetup.A4, new RemoteImageOptions()));
    }

    [Fact]
    public async Task ConvertToFilePageOverloads_RejectNullHtml()
    {
        var path = Path.Join(Path.GetTempPath(), $"doctoolkit-{Guid.NewGuid():N}.out");

        await Assert.ThrowsAsync<ArgumentNullException>(
            () => HtmlToDocxConverter.ConvertToFileAsync(null!, PageSetup.A4, path));
        await Assert.ThrowsAsync<ArgumentNullException>(
            () => HtmlToPdfConverter.ConvertToFileAsync(null!, PageSetup.A4, path));

        Assert.False(File.Exists(path), "A rejected conversion still created a file.");
    }

    [Fact]
    public void Create_WithAHeader_EmitsAReferenceBeforeThePageSize()
    {
        var page = PageSetup.A4.WithHeader(DocxHeader.Text("Contoso"));

        var sectPr = SectionPropertiesOf(DocxEditor.Create(Blocks, page));

        Assert.NotNull(sectPr);
        var children = sectPr!.ChildElements.ToList();
        var referenceIndex = children.FindIndex(c => c is HeaderReference);
        var sizeIndex = children.FindIndex(c => c is PageSize);

        Assert.True(referenceIndex >= 0, "No w:headerReference was emitted.");

        // The schema is order-sensitive and Word calls a misordered sectPr corrupt - the same trap
        // already recorded for sectPr needing to be the last child of w:body.
        Assert.True(
            referenceIndex < sizeIndex,
            $"w:headerReference must precede w:pgSz; got indexes {referenceIndex} and {sizeIndex}.");
    }

    [Fact]
    public void Create_WithNoFirstPage_EmitsNoTitlePg()
    {
        var page = PageSetup.A4.WithHeader(DocxHeader.Text("Contoso"));

        var sectPr = SectionPropertiesOf(DocxEditor.Create(Blocks, page));

        Assert.Null(sectPr!.GetFirstChild<TitlePage>());
    }

    [Fact]
    public void Create_WithAFirstPage_EmitsTitlePgAfterThePageMargin()
    {
        var page = PageSetup.A4
            .WithHeader(DocxHeader.Text("running"))
            .WithFirstPage(DocxHeader.Text("cover"), null);

        var sectPr = SectionPropertiesOf(DocxEditor.Create(Blocks, page));

        Assert.NotNull(sectPr!.GetFirstChild<TitlePage>());

        var children = sectPr.ChildElements.ToList();
        Assert.True(
            children.FindIndex(c => c is TitlePage) > children.FindIndex(c => c is PageMargin),
            "w:titlePg must follow w:pgMar.");
    }

    [Fact]
    public async Task HtmlToDocx_HonoursTheHeaderToo()
    {
        var page = PageSetup.A4.WithHeader(DocxHeader.Text("FromHtml"));

        var docx = await HtmlToDocxConverter.ConvertAsync("<p>Body</p>", page);

        Assert.Contains("FromHtml", DocxEditor.ExtractText(docx, includeHeadersAndFooters: true),
            StringComparison.Ordinal);
    }

    [Fact]
    public void EveryVariantTogetherIsSchemaValid()
    {
        var page = PageSetup.Letter
            .WithHeader(DocxHeader.Text("H", HeaderAlignment.Center))
            .WithFooter(DocxHeader.Of(HeaderAlignment.Right, DocxHeaderSegment.PageNumber))
            .WithFirstPage(DocxHeader.Text("first"), DocxHeader.Text("first foot"));

        using var ms = new MemoryStream(DocxEditor.Create(Blocks, page));
        using var doc = WordprocessingDocument.Open(ms, false);

        var errors = new DocumentFormat.OpenXml.Validation.OpenXmlValidator().Validate(doc).ToList();

        Assert.True(errors.Count == 0,
            "Schema errors: " + string.Join(" | ", errors.Select(e => e.Description)));
    }
}
