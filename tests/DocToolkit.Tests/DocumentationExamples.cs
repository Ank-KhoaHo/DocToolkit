using Xunit;

namespace DocToolkit.Tests;

/// <summary>
/// The usage snippets shown on the API documentation site.
///
/// They live here — in compiled, executed tests — and are pulled into the XML doc comments by DocFX
/// with <c>&lt;code source="..." region="..."/&gt;</c> rather than being retyped into a comment.
///
/// That is the whole point. A snippet nobody compiles is a claim nobody checks, and this repository
/// has had to correct enough of those. If an example stops compiling the build fails; if it stops
/// working this test fails. An example drifting from the API it documents is not possible.
///
/// Keep them short: this is the first code a consumer reads, so each region shows one capability.
/// Setup lines sit ABOVE the region and assertions BELOW it, so neither reaches the reader.
/// </summary>
public class DocumentationExamples
{
    [Fact]
    public async Task HtmlToDocxExample()
    {
        #region HtmlToDocx
        string html = "<h1>Quarterly Report</h1><p>Revenue was up <strong>12%</strong>.</p>";

        byte[] docx = await HtmlToDocxConverter.ConvertAsync(html);
        #endregion

        Assert.Contains("Quarterly Report", DocxEditor.ExtractText(docx), StringComparison.Ordinal);
    }

    [Fact]
    public async Task HtmlToPdfExample()
    {
        #region HtmlToPdf
        string html = "<h1>Invoice 2026-114</h1><p>Due on receipt.</p>";

        // A4 portrait with one-inch margins is the default; PageSetup overrides it.
        byte[] pdf = await HtmlToPdfConverter.ConvertAsync(
            html,
            PageSetup.A4.Landscape().WithMargins(36));
        #endregion

        var box = Assert.Single(PdfProbe.MediaBoxes(pdf));
        Assert.Equal(842, box.Width, 0);
    }

    [Fact]
    public async Task HtmlToPdfOptionsExample()
    {
        string html = "<h1>Invoice 2026-114</h1><p>Due on receipt.</p>";

        #region HtmlToPdfOptions
        // One object when a conversion needs more than one of the three axes. Each property
        // defaults to what the simpler overloads do, so setting only what you care about is safe.
        byte[] pdf = await HtmlToPdfConverter.ConvertAsync(html, new HtmlToPdfOptions
        {
            Page = PageSetup.Letter.WithMargins(36),
            // Fonts = new PdfFontOptions("Noto Sans", File.ReadAllBytes("NotoSans-Regular.ttf")),
            // RemoteImage = new RemoteImageOptions { AllowedHosts = ["cdn.example.com"] },
        });
        #endregion

        var letter = Assert.Single(PdfProbe.MediaBoxes(pdf));
        Assert.Equal(612, letter.Width, 0);
    }

    [Fact]
    public void DocxToPdfExample()
    {
        byte[] docx = DocxEditor.Create(new[] { DocxBlock.Paragraph("Signed.") });

        #region DocxToPdf
        byte[] pdf = DocxToPdfConverter.Convert(docx);
        #endregion

        Assert.True(PdfProbe.IsPdf(pdf));
    }

    [Fact]
    public void DocxCreateExample()
    {
        #region DocxCreate
        byte[] docx = DocxEditor.Create(new[]
        {
            DocxBlock.Heading("Quarterly Report", 1),
            DocxBlock.Paragraph("Revenue was up 12%."),
            DocxBlock.Table(
                new[] { "Region", "Total" },
                new[] { new object?[] { "North", 1200 } }),
        });
        #endregion

        Assert.Contains("North", DocxEditor.ExtractText(docx), StringComparison.Ordinal);
    }

    [Fact]
    public void DocxReplaceTextExample()
    {
        byte[] template = DocxEditor.Create(new[]
        {
            DocxBlock.Paragraph("Dear {{customer}}, your invoice {{number}} is ready."),
        });

        #region DocxReplaceText
        byte[] filled = DocxEditor.ReplaceText(template, new Dictionary<string, string>
        {
            ["{{customer}}"] = "Acme Ltd",
            ["{{number}}"] = "2026-114",
        });
        #endregion

        Assert.Contains("Acme Ltd", DocxEditor.ExtractText(filled), StringComparison.Ordinal);
    }

    [Fact]
    public void DocxReadTableExample()
    {
        byte[] docx = DocxEditor.Create(new[]
        {
            DocxBlock.Table(
                new[] { "Region", "Total" },
                new[] { new object?[] { "North", 1200 } }),
        });

        #region DocxReadTable
        int count = DocxEditor.TableCount(docx);
        IReadOnlyList<IReadOnlyList<string>> table = DocxEditor.ReadTable(docx, 0);
        #endregion

        Assert.Equal(1, count);
        Assert.Equal(new[] { "Region", "Total" }, table[0]);
    }

    [Fact]
    public void DocxToHtmlExample()
    {
        byte[] docx = DocxEditor.Create(new[] { DocxBlock.Heading("Quarterly Report", 1) });

        #region DocxToHtml
        // A complete HTML document, not a fragment, with images embedded as data: URIs so the
        // result is self-contained. Extract the body with a parser if you are embedding it.
        string html = DocxToHtmlConverter.Convert(docx);
        #endregion

        Assert.Contains("<h1", html, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void DocxToMarkdownExample()
    {
        byte[] docx = DocxEditor.Create(new[] { DocxBlock.Heading("Quarterly Report", 1) });

        #region DocxToMarkdown
        // Structure survives: a heading becomes "#", a table becomes a pipe table. ExtractText, by
        // contrast, returns the same words as flat lines with no markup at all.
        string markdown = DocxToMarkdownConverter.Convert(docx);
        #endregion

        Assert.Contains("# Quarterly Report", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void WorkbookCreateExample()
    {
        #region WorkbookCreate
        byte[] xlsx = WorkbookEditor.Create(new[]
        {
            XlsxSheet.Named("Sales", new[]
            {
                new object?[] { "Region", "Q1" },
                new object?[] { "North", 1200 },
                new object?[] { "Total", XlsxFormula.From("=SUM(B2:B2)") },
            }),
        });
        #endregion

        Assert.Equal("1200", WorkbookEditor.ReadCell(xlsx, "Sales", "B3"));
    }

    [Fact]
    public void WorkbookReadSheetExample()
    {
        byte[] xlsx = WorkbookEditor.Create("Sales", new[]
        {
            new object?[] { "Region", "Total" },
            new object?[] { "North", 1200 },
        });

        #region WorkbookReadSheet
        // No need to know the workbook's shape in advance.
        foreach (string name in WorkbookEditor.SheetNames(xlsx))
        {
            IReadOnlyList<IReadOnlyList<string>> grid = WorkbookEditor.ReadSheet(xlsx, name);
            Console.WriteLine($"{name}: {grid.Count} rows x {grid[0].Count} columns");
        }
        #endregion

        Assert.Single(WorkbookEditor.SheetNames(xlsx));
    }

    [Fact]
    public void XlsxToPdfExample()
    {
        byte[] xlsx = WorkbookEditor.Create("Sales", new[]
        {
            new object?[] { "Region", "Total" },
            new object?[] { "North", 1200 },
        });

        #region XlsxToPdf
        byte[] pdf = XlsxToPdfConverter.Convert(xlsx);
        #endregion

        Assert.True(PdfProbe.IsPdf(pdf));
    }

    [Fact]
    public void PresentationCreateExample()
    {
        #region PresentationCreate
        byte[] pptx = PresentationEditor.Create(new[]
        {
            PptxSlide.Titled("Quarterly Report", "Revenue up 12%"),
            PptxSlide.Titled("Outlook", "Hiring 3 engineers"),
        });
        #endregion

        Assert.Equal(2, PresentationEditor.SlideCount(pptx));
    }

    [Fact]
    public void PptxToPdfExample()
    {
        byte[] pptx = PresentationEditor.Create(new[] { PptxSlide.Titled("One", "Body") });

        #region PptxToPdf
        // One page per slide, at the deck's own 16:9 slide geometry rather than a paper size.
        byte[] pdf = PptxToPdfConverter.Convert(pptx);
        #endregion

        Assert.Single(PdfProbe.MediaBoxes(pdf));
    }
}
