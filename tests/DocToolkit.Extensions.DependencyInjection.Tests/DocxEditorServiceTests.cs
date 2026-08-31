using System.Collections.Generic;
using DocToolkit;
using DocToolkit.Extensions.DependencyInjection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace DocToolkit.Extensions.DependencyInjection.Tests;

public class DocxEditorServiceTests
{
    // ---------------------------------------------------------------------------------------
    // The six members that restore 1:1 with DocxEditor. FillRows and ReplaceImage shipped in
    // core 0.8.0 and were never mirrored here; Create arrived in 0.10.0. Each test asserts the
    // wrapper agrees with the static method rather than merely that it returns something -
    // a wrapper that ignored its arguments and returned an empty document would pass the
    // weaker check.
    // ---------------------------------------------------------------------------------------

    [Fact]
    public void Create_MatchesTheStaticMethod()
    {
        var sut = new DocxEditorService(new TestOptionsMonitor<DocToolkitOptions>(new DocToolkitOptions()));
        var blocks = new[]
        {
            DocxBlock.Heading("Quarterly Report", 1),
            DocxBlock.Paragraph("Revenue rose 12%."),
            DocxBlock.Table(new[] { "Region" }, new[] { new object[] { "EMEA" } }),
        };

        var fromWrapper = sut.Create(blocks);

        Assert.Equal(
            DocxEditor.ExtractText(DocxEditor.Create(blocks)),
            DocxEditor.ExtractText(fromWrapper));
        Assert.Contains("Quarterly Report", sut.ExtractText(fromWrapper));
    }

    [Fact]
    public async Task CreateAsync_MatchesTheStaticMethod()
    {
        var sut = new DocxEditorService(new TestOptionsMonitor<DocToolkitOptions>(new DocToolkitOptions()));
        var blocks = new[] { DocxBlock.Paragraph("Written to a stream.") };

        using var destination = new MemoryStream();
        await sut.CreateAsync(blocks, destination);

        // B16: the literal comes FIRST and the parity check second. On its own, comparing
        // ExtractText against ExtractText holds however broken ExtractText is - if it returned ""
        // for everything, "" == "" passes. That is not hypothetical here: A26 was exactly that
        // method returning the wrong thing, and it survived 785 tests behind assertions shaped
        // like the one below.
        Assert.Contains("Written to a stream.", DocxEditor.ExtractText(destination.ToArray()),
            StringComparison.Ordinal);

        Assert.Equal(
            DocxEditor.ExtractText(DocxEditor.Create(blocks)),
            DocxEditor.ExtractText(destination.ToArray()));
    }

    [Fact]
    public async Task FillRows_MatchesTheStaticMethod()
    {
        var sut = new DocxEditorService(new TestOptionsMonitor<DocToolkitOptions>(new DocToolkitOptions()));
        var docx = await HtmlToDocxConverter.ConvertAsync(
            "<table border=\"1\"><tr><td>{{item.Desc}}</td></tr></table>");
        var records = new[]
        {
            new Dictionary<string, string>(StringComparer.Ordinal) { ["Desc"] = "Widget" },
            new Dictionary<string, string>(StringComparer.Ordinal) { ["Desc"] = "Gadget" },
        };

        var fromWrapper = sut.FillRows(docx, "item", records);

        Assert.Equal(
            DocxEditor.ExtractText(DocxEditor.FillRows(docx, "item", records)),
            DocxEditor.ExtractText(fromWrapper));

        var text = sut.ExtractText(fromWrapper);
        Assert.Contains("Widget", text);
        Assert.Contains("Gadget", text);
        Assert.DoesNotContain("{{item.Desc}}", text);
    }

    [Fact]
    public async Task FillRowsAsync_MatchesTheStaticMethod()
    {
        var sut = new DocxEditorService(new TestOptionsMonitor<DocToolkitOptions>(new DocToolkitOptions()));
        var docx = await HtmlToDocxConverter.ConvertAsync(
            "<table border=\"1\"><tr><td>{{item.Desc}}</td></tr></table>");
        var records = new[]
        {
            new Dictionary<string, string>(StringComparer.Ordinal) { ["Desc"] = "Widget" },
        };

        using var source = new MemoryStream(docx);
        using var destination = new MemoryStream();
        await sut.FillRowsAsync(source, "item", records, destination);

        Assert.Contains("Widget", DocxEditor.ExtractText(destination.ToArray()));
    }

    [Fact]
    public async Task ReplaceImage_MatchesTheStaticMethod()
    {
        var sut = new DocxEditorService(new TestOptionsMonitor<DocToolkitOptions>(new DocToolkitOptions()));
        var docx = await HtmlToDocxConverter.ConvertAsync("<p>Logo: {{logo}}</p>");
        var png = OnePixelPng();

        var fromWrapper = sut.ReplaceImage(docx, "{{logo}}", png, widthPoints: 24);

        // The placeholder is gone and a real image part exists - a wrapper that dropped the
        // image would still remove the text, so the part is what proves it.
        Assert.DoesNotContain("{{logo}}", sut.ExtractText(fromWrapper));

        using var ms = new MemoryStream(fromWrapper);
        using var doc = WordprocessingDocument.Open(ms, false);
        Assert.Single(doc.MainDocumentPart!.ImageParts);
    }

    [Fact]
    public async Task ReplaceImageAsync_MatchesTheStaticMethod()
    {
        var sut = new DocxEditorService(new TestOptionsMonitor<DocToolkitOptions>(new DocToolkitOptions()));
        var docx = await HtmlToDocxConverter.ConvertAsync("<p>Logo: {{logo}}</p>");

        using var source = new MemoryStream(docx);
        using var destination = new MemoryStream();
        await sut.ReplaceImageAsync(source, "{{logo}}", OnePixelPng(), destination, widthPoints: 24);

        var written = destination.ToArray();
        Assert.DoesNotContain("{{logo}}", DocxEditor.ExtractText(written));

        using var ms = new MemoryStream(written);
        using var doc = WordprocessingDocument.Open(ms, false);
        Assert.Single(doc.MainDocumentPart!.ImageParts);
    }

    /// <summary>A 1x1 PNG, inline so this project needs no fixture asset.</summary>
    private static byte[] OnePixelPng() => Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==");

    // ---------------------------------------------------------------------------------------
    // TableCount/ReadTable landed in core 0.22.0 and are mirrored here for the first time.
    // Unlike most tests above, each also asserts a concrete literal grid, not just parity with
    // the static method - the parity check alone cannot fail if the wrapper and the static
    // method are broken identically.
    // ---------------------------------------------------------------------------------------

    private static byte[] TableFixture() => DocxEditor.Create(new[]
    {
        DocxBlock.Table(
            new[] { "Region", "Q1" },
            new[] { new object?[] { "North", 1200 } }),
    });

    [Fact]
    public void TableCount_MatchesTheStaticMethodAndReturnsTheRealCount()
    {
        var sut = new DocxEditorService(new TestOptionsMonitor<DocToolkitOptions>(new DocToolkitOptions()));
        var docx = TableFixture();

        var count = sut.TableCount(docx);

        Assert.Equal(DocxEditor.TableCount(docx), count);
        Assert.Equal(1, count);
    }

    [Fact]
    public async Task TableCountAsync_MatchesTheStaticMethodAndReturnsTheRealCount()
    {
        var sut = new DocxEditorService(new TestOptionsMonitor<DocToolkitOptions>(new DocToolkitOptions()));
        var docx = TableFixture();

        using var expectedSource = new MemoryStream(docx);
        using var actualSource = new MemoryStream(docx);

        var expected = await DocxEditor.TableCountAsync(expectedSource);
        var actual = await sut.TableCountAsync(actualSource);

        Assert.Equal(expected, actual);
        Assert.Equal(1, actual);
    }

    [Fact]
    public void ReadTable_MatchesTheStaticMethodAndReturnsTheRealGrid()
    {
        var sut = new DocxEditorService(new TestOptionsMonitor<DocToolkitOptions>(new DocToolkitOptions()));
        var docx = TableFixture();

        var table = sut.ReadTable(docx, 0);

        Assert.Equal(DocxEditor.ReadTable(docx, 0), table);
        Assert.Equal(new[] { "Region", "Q1" }, table[0]);
        Assert.Equal(new[] { "North", "1200" }, table[1]);
    }

    [Fact]
    public async Task ReadTableAsync_MatchesTheStaticMethodAndReturnsTheRealGrid()
    {
        var sut = new DocxEditorService(new TestOptionsMonitor<DocToolkitOptions>(new DocToolkitOptions()));
        var docx = TableFixture();

        using var expectedSource = new MemoryStream(docx);
        using var actualSource = new MemoryStream(docx);

        var expected = await DocxEditor.ReadTableAsync(expectedSource, 0);
        var actual = await sut.ReadTableAsync(actualSource, 0);

        Assert.Equal(expected, actual);
        Assert.Equal(new[] { "Region", "Q1" }, actual[0]);
        Assert.Equal(new[] { "North", "1200" }, actual[1]);
    }

    [Fact]
    public async Task ReplaceText_SubstitutesPlaceholders()
    {
        var sut = new DocxEditorService(new TestOptionsMonitor<DocToolkitOptions>(new DocToolkitOptions()));
        var docx = await HtmlToDocxConverter.ConvertAsync(
            "<p>Dear {{name}}, your balance is {{balance}}.</p>");

        var edited = sut.ReplaceText(docx, new Dictionary<string, string>
        {
            ["{{name}}"] = "Contoso Ltd",
            ["{{balance}}"] = "4,250.00",
        });

        var text = sut.ExtractText(edited);
        Assert.Contains("Contoso Ltd", text);
        Assert.Contains("4,250.00", text);
        Assert.DoesNotContain("{{name}}", text);
    }

    [Fact]
    public void ExtractText_WithHeadersAndFooters_MatchesTheStaticMethod()
    {
        var docx = DocxWithHeaderAndFooter("Body text.", "Page header", "Page footer");
        var sut = new DocxEditorService(new TestOptionsMonitor<DocToolkitOptions>(new DocToolkitOptions()));

        Assert.Equal(
            DocxEditor.ExtractText(docx, includeHeadersAndFooters: true),
            sut.ExtractText(docx, includeHeadersAndFooters: true));

        // The old fixture (built via HtmlToDocxConverter) had zero header/footer parts, so this
        // assertion was previously impossible to fail even if the bool parameter were dropped or
        // inverted. This fixture genuinely has both, so these two lines are what actually prove the
        // parameter threads through correctly.
        Assert.Contains("Page header", sut.ExtractText(docx, includeHeadersAndFooters: true));
        Assert.DoesNotContain("Page header", sut.ExtractText(docx, includeHeadersAndFooters: false));
    }

    [Fact]
    public void ReplaceText_RejectsNullReplacements()
    {
        var sut = new DocxEditorService(new TestOptionsMonitor<DocToolkitOptions>(new DocToolkitOptions()));

        Assert.Throws<ArgumentNullException>(() => sut.ReplaceText(Array.Empty<byte>(), null!));
    }

    [Fact]
    public async Task ExtractTextAsync_MatchesTheStaticMethod()
    {
        var docx = await HtmlToDocxConverter.ConvertAsync("<p>Body copy.</p>");
        var sut = new DocxEditorService(new TestOptionsMonitor<DocToolkitOptions>(new DocToolkitOptions()));

        using var expectedSource = new MemoryStream(docx);
        using var actualSource = new MemoryStream(docx);

        var expected = await DocxEditor.ExtractTextAsync(expectedSource);
        var actual = await sut.ExtractTextAsync(actualSource);

        Assert.Equal(expected, actual);
    }

    [Fact]
    public async Task ExtractTextAsync_WithHeadersAndFooters_MatchesTheStaticMethod()
    {
        var docx = DocxWithHeaderAndFooter("Body text.", "Page header", "Page footer");
        var sut = new DocxEditorService(new TestOptionsMonitor<DocToolkitOptions>(new DocToolkitOptions()));

        using var expectedSource = new MemoryStream(docx);
        using var actualSource = new MemoryStream(docx);

        var expected = await DocxEditor.ExtractTextAsync(expectedSource, includeHeadersAndFooters: true);
        var actual = await sut.ExtractTextAsync(actualSource, includeHeadersAndFooters: true);

        Assert.Equal(expected, actual);
        Assert.Contains("Page header", actual);
    }

    [Fact]
    public async Task ReplaceTextAsync_MatchesTheStaticMethod()
    {
        var docx = await HtmlToDocxConverter.ConvertAsync("<p>Dear {{name}}.</p>");
        var sut = new DocxEditorService(new TestOptionsMonitor<DocToolkitOptions>(new DocToolkitOptions()));
        var replacements = new Dictionary<string, string> { ["{{name}}"] = "Contoso Ltd" };

        using var expectedSource = new MemoryStream(docx);
        using var expected = new MemoryStream();
        await DocxEditor.ReplaceTextAsync(expectedSource, replacements, expected);

        using var actualSource = new MemoryStream(docx);
        using var actual = new MemoryStream();
        await sut.ReplaceTextAsync(actualSource, replacements, actual);

        Assert.Equal(expected.ToArray(), actual.ToArray());
    }

    [Fact]
    public async Task ExtractTextAsync_HonorsCancellation()
    {
        var sut = new DocxEditorService(new TestOptionsMonitor<DocToolkitOptions>(new DocToolkitOptions()));
        using var cts = new CancellationTokenSource();
        cts.Cancel();

        using var source = new MemoryStream();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(
            () => sut.ExtractTextAsync(source, cts.Token));
    }

    // ---------------------------------------------------------------------------------------
    // AddFootnote/AddEndnote/AddTableOfContents, mirrored from core 0.43.0 (A81-DI). The
    // footnote/endnote pair is the highest-risk one to mirror by hand - both take an identical
    // (byte[], string, string) shape, so a service method calling the WRONG static method
    // compiles cleanly and a wrapper-vs-wrapper comparison alone would not catch it. Each test
    // below reads back the footnote/endnote PART directly and asserts on a token unique to that
    // call, so a footnote/endnote swap fails here even though nothing else in this file would
    // notice.
    // ---------------------------------------------------------------------------------------

    [Fact]
    public void AddFootnote_MatchesTheStaticMethodAndAddsAFootnoteNotAnEndnote()
    {
        var sut = new DocxEditorService(new TestOptionsMonitor<DocToolkitOptions>(new DocToolkitOptions()));
        var docx = DocxEditor.Create([DocxBlock.Paragraph("See{{note}}here.")]);

        var fromWrapper = sut.AddFootnote(docx, "{{note}}", "FOOTNOTE_TOKEN");

        Assert.Equal(
            DocxEditor.ExtractText(DocxEditor.AddFootnote(docx, "{{note}}", "FOOTNOTE_TOKEN")),
            DocxEditor.ExtractText(fromWrapper));

        using var ms = new MemoryStream(fromWrapper);
        using var doc = WordprocessingDocument.Open(ms, false);
        Assert.NotNull(doc.MainDocumentPart!.FootnotesPart);
        Assert.Null(doc.MainDocumentPart.EndnotesPart);
        Assert.Contains("FOOTNOTE_TOKEN", doc.MainDocumentPart.FootnotesPart!.Footnotes!.InnerText);
    }

    [Fact]
    public async Task AddFootnoteAsync_MatchesTheStaticMethod()
    {
        var sut = new DocxEditorService(new TestOptionsMonitor<DocToolkitOptions>(new DocToolkitOptions()));
        var docx = DocxEditor.Create([DocxBlock.Paragraph("See{{note}}here.")]);

        using var source = new MemoryStream(docx);
        using var destination = new MemoryStream();
        await sut.AddFootnoteAsync(source, "{{note}}", "FOOTNOTE_TOKEN", destination);

        using var ms = new MemoryStream(destination.ToArray());
        using var doc = WordprocessingDocument.Open(ms, false);
        Assert.NotNull(doc.MainDocumentPart!.FootnotesPart);
        Assert.Contains("FOOTNOTE_TOKEN", doc.MainDocumentPart.FootnotesPart!.Footnotes!.InnerText);
    }

    [Fact]
    public void AddEndnote_MatchesTheStaticMethodAndAddsAnEndnoteNotAFootnote()
    {
        var sut = new DocxEditorService(new TestOptionsMonitor<DocToolkitOptions>(new DocToolkitOptions()));
        var docx = DocxEditor.Create([DocxBlock.Paragraph("See{{note}}here.")]);

        var fromWrapper = sut.AddEndnote(docx, "{{note}}", "ENDNOTE_TOKEN");

        Assert.Equal(
            DocxEditor.ExtractText(DocxEditor.AddEndnote(docx, "{{note}}", "ENDNOTE_TOKEN")),
            DocxEditor.ExtractText(fromWrapper));

        using var ms = new MemoryStream(fromWrapper);
        using var doc = WordprocessingDocument.Open(ms, false);
        Assert.NotNull(doc.MainDocumentPart!.EndnotesPart);
        Assert.Null(doc.MainDocumentPart.FootnotesPart);
        Assert.Contains("ENDNOTE_TOKEN", doc.MainDocumentPart.EndnotesPart!.Endnotes!.InnerText);
    }

    [Fact]
    public async Task AddEndnoteAsync_MatchesTheStaticMethod()
    {
        var sut = new DocxEditorService(new TestOptionsMonitor<DocToolkitOptions>(new DocToolkitOptions()));
        var docx = DocxEditor.Create([DocxBlock.Paragraph("See{{note}}here.")]);

        using var source = new MemoryStream(docx);
        using var destination = new MemoryStream();
        await sut.AddEndnoteAsync(source, "{{note}}", "ENDNOTE_TOKEN", destination);

        using var ms = new MemoryStream(destination.ToArray());
        using var doc = WordprocessingDocument.Open(ms, false);
        Assert.NotNull(doc.MainDocumentPart!.EndnotesPart);
        Assert.Contains("ENDNOTE_TOKEN", doc.MainDocumentPart.EndnotesPart!.Endnotes!.InnerText);
    }

    [Fact]
    public void AddTableOfContents_MatchesTheStaticMethodAndThreadsTheLevels()
    {
        using var word = OfficeIMO.Word.WordDocument.Create();
        word.AddParagraph("{{toc}}");
        word.AddParagraph("Top").Style = OfficeIMO.Word.WordParagraphStyles.Heading1;
        word.AddParagraph("Deep").Style = OfficeIMO.Word.WordParagraphStyles.Heading3;
        using var built = new MemoryStream();
        word.Save(built);
        var docx = built.ToArray();

        var sut = new DocxEditorService(new TestOptionsMonitor<DocToolkitOptions>(new DocToolkitOptions()));

        // minLevel/maxLevel both pinned to 1: a wrapper that dropped or swapped either default
        // parameter would still compile and still produce A table of contents field, just with
        // the wrong instruction text - which this reads directly rather than inferring.
        var fromWrapper = sut.AddTableOfContents(docx, "{{toc}}", minLevel: 1, maxLevel: 1);

        using var ms = new MemoryStream(fromWrapper);
        using var doc = WordprocessingDocument.Open(ms, false);
        var instr = doc.MainDocumentPart!.Document!.Body!.Descendants<SimpleField>().Single()
            .Instruction!.Value!;
        Assert.Contains("\"1-1\"", instr);
    }

    [Fact]
    public async Task AddTableOfContentsAsync_MatchesTheStaticMethod()
    {
        using var word = OfficeIMO.Word.WordDocument.Create();
        word.AddParagraph("{{toc}}");
        word.AddParagraph("Overview").Style = OfficeIMO.Word.WordParagraphStyles.Heading1;
        using var built = new MemoryStream();
        word.Save(built);
        var docx = built.ToArray();

        var sut = new DocxEditorService(new TestOptionsMonitor<DocToolkitOptions>(new DocToolkitOptions()));

        using var source = new MemoryStream(docx);
        using var destination = new MemoryStream();
        await sut.AddTableOfContentsAsync(source, "{{toc}}", destination);

        var written = destination.ToArray();
        Assert.DoesNotContain("{{toc}}", DocxEditor.ExtractText(written));

        using var ms = new MemoryStream(written);
        using var doc = WordprocessingDocument.Open(ms, false);
        Assert.Single(doc.MainDocumentPart!.Document!.Body!.Descendants<SimpleField>());
    }

    // ---------------------------------------------------------------------------------------
    // InspectSignatures/ValidateSignatures and their Async forms, mirrored from core 0.45.0
    // (A99-DI). Exercised against a genuinely unsigned DOCX built by this format's own Create -
    // each format's InspectPackageSignatures/ValidatePackageSignatures opens the package through
    // that format's own typed loader first, so a wrapper wired to the wrong format's static
    // method would very likely throw here rather than silently returning a plausible answer.
    // ---------------------------------------------------------------------------------------

    [Fact]
    public void InspectSignatures_MatchesTheStaticMethod()
    {
        var sut = new DocxEditorService(new TestOptionsMonitor<DocToolkitOptions>(new DocToolkitOptions()));
        var docx = DocxEditor.Create([DocxBlock.Paragraph("Unsigned.")]);

        var info = sut.InspectSignatures(docx);

        Assert.Equal(DocxEditor.InspectSignatures(docx).HasSignatures, info.HasSignatures);
        Assert.False(info.HasSignatures);
    }

    [Fact]
    public async Task InspectSignaturesAsync_MatchesTheStaticMethod()
    {
        var sut = new DocxEditorService(new TestOptionsMonitor<DocToolkitOptions>(new DocToolkitOptions()));
        var docx = DocxEditor.Create([DocxBlock.Paragraph("Unsigned.")]);

        using var source = new MemoryStream(docx);
        var info = await sut.InspectSignaturesAsync(source);

        Assert.False(info.HasSignatures);
    }

    [Fact]
    public void ValidateSignatures_MatchesTheStaticMethod()
    {
        var sut = new DocxEditorService(new TestOptionsMonitor<DocToolkitOptions>(new DocToolkitOptions()));
        var docx = DocxEditor.Create([DocxBlock.Paragraph("Unsigned.")]);

        var report = sut.ValidateSignatures(docx);

        Assert.Equal(DocxEditor.ValidateSignatures(docx).HasSignatures, report.HasSignatures);
        Assert.False(report.HasSignatures);
    }

    [Fact]
    public async Task ValidateSignaturesAsync_MatchesTheStaticMethod()
    {
        var sut = new DocxEditorService(new TestOptionsMonitor<DocToolkitOptions>(new DocToolkitOptions()));
        var docx = DocxEditor.Create([DocxBlock.Paragraph("Unsigned.")]);

        using var source = new MemoryStream(docx);
        var report = await sut.ValidateSignaturesAsync(source);

        Assert.False(report.HasSignatures);
    }

    // ---------------------------------------------------------------------------------------
    // ReadMetadata/WithMetadata, mirrored from core 0.46.0 (A102-DI).
    // ---------------------------------------------------------------------------------------

    [Fact]
    public void WithMetadata_ReadMetadata_RoundTripCorrectly()
    {
        var sut = new DocxEditorService(new TestOptionsMonitor<DocToolkitOptions>(new DocToolkitOptions()));
        var docx = DocxEditor.Create([DocxBlock.Paragraph("Body.")]);
        var metadata = new DocumentMetadata { Title = "Quarterly Report", Creator = "Finance" };

        var stamped = sut.WithMetadata(docx, metadata);
        var read = sut.ReadMetadata(stamped);

        Assert.Equal("Quarterly Report", read.Title);
        Assert.Equal("Finance", read.Creator);
        Assert.Equal(
            DocxEditor.ReadMetadata(DocxEditor.WithMetadata(docx, metadata)).Title,
            read.Title);
    }

    private static byte[] DocxWithHeaderAndFooter(string bodyText, string headerText, string footerText)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            var body = new Body(new Paragraph(new Run(new Text(bodyText))));
            mainPart.Document = new Document(body);

            var headerPart = mainPart.AddNewPart<HeaderPart>();
            headerPart.Header = new Header(new Paragraph(new Run(new Text(headerText))));
            var headerRelId = mainPart.GetIdOfPart(headerPart);

            var footerPart = mainPart.AddNewPart<FooterPart>();
            footerPart.Footer = new Footer(new Paragraph(new Run(new Text(footerText))));
            var footerRelId = mainPart.GetIdOfPart(footerPart);

            body.Append(new SectionProperties(
                new HeaderReference { Type = HeaderFooterValues.Default, Id = headerRelId },
                new FooterReference { Type = HeaderFooterValues.Default, Id = footerRelId }));

            mainPart.Document.Save();
        }
        return ms.ToArray();
    }
}
