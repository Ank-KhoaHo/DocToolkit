using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using V = DocumentFormat.OpenXml.Vml;

namespace DocToolkit.Tests;

/// <summary>
/// Hand-built .docx fixtures.
///
/// The HtmlToOpenXml fixtures used elsewhere emit one <c>w:t</c> per paragraph, so they cannot
/// exercise the split-run path at all — which is the entire reason DocxEditor merges runs. These
/// helpers build the body XML directly so a test can pin down exactly how many runs there are,
/// what formatting each carries, and where a placeholder straddles a run boundary. This mirrors
/// the hand-splitting already done in <see cref="PresentationEditorTests"/>.
/// </summary>
internal static class DocxFixtures
{
    /// <summary>A run carrying <paramref name="text"/>, optionally bold.</summary>
    public static Run R(string text, bool bold = false)
    {
        var run = new Run();
        if (bold) run.AppendChild(new RunProperties(new Bold()));
        run.AppendChild(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        return run;
    }

    /// <summary>A paragraph made of <paramref name="children"/> (runs, hyperlinks, ...).</summary>
    public static Paragraph P(params OpenXmlElement[] children) => new(children);

    /// <summary>
    /// A run holding a legacy VML text box — the shape Word produces for "Insert &gt; Text Box".
    /// The key structural point is that a whole <c>w:p</c> lives inside <c>w:txbxContent</c>,
    /// nested within a run of the *outer* paragraph.
    /// </summary>
    public static Run TextBoxRun(string innerText) => new(
        new Picture(
            new V.Shape(
                new V.TextBox(
                    new TextBoxContent(
                        new Paragraph(new Run(new Text(innerText) { Space = SpaceProcessingModeValues.Preserve })))))
            {
                Id = "TextBox1",
                Type = "#_x0000_t202",
                Style = "position:absolute;width:200pt;height:50pt",
            }));

    /// <summary>A one-cell table row whose cell holds a single paragraph of <paramref name="children"/>.</summary>
    public static TableRow Row(params OpenXmlElement[] children) => new(new TableCell(P(children)));

    /// <summary>A one-cell table row holding <paramref name="cellChildren"/> directly — a cell may
    /// contain a nested table as well as paragraphs, which is how the nesting tests are built.</summary>
    public static TableRow RowOf(params OpenXmlElement[] cellChildren) => new(new TableCell(cellChildren));

    /// <summary>
    /// A table made of <paramref name="rows"/>.
    ///
    /// The <c>w:tblGrid</c> is not decoration: without it the schema validator rejects the first
    /// <c>w:tr</c> as an unexpected child, because <c>tblPr</c> and <c>tblGrid</c> must precede any
    /// row. An earlier version of this helper omitted it and produced tables that passed every test
    /// while being schema-invalid.
    /// </summary>
    /// <remarks>
    /// <b>Takes <c>OpenXmlElement</c>, not <c>TableRow</c>.</b> A <c>w:sdt</c> at row level is an
    /// <c>SdtRow</c>, which is not a <c>TableRow</c> at all — so a <c>TableRow[]</c> signature
    /// cannot express a table with a wrapped row, and A77's tests had to declare a second builder
    /// to say it. One builder, because two would be a second answer to "what does a table look
    /// like" inside the very tests written to close a disagreement about what a table is.
    ///
    /// <para>The column count looks through a row-level control, so a wrapped row still counts
    /// toward the grid width.</para>
    /// </remarks>
    public static Table Tbl(params OpenXmlElement[] rows)
    {
        static int CellCount(OpenXmlElement row) => row switch
        {
            TableRow plain => plain.ChildElements.OfType<TableCell>().Count(),
            SdtRow control => control.SdtContentRow?.Elements<TableRow>()
                                     .Select(r => r.ChildElements.OfType<TableCell>().Count())
                                     .DefaultIfEmpty(0).Max() ?? 0,
            _ => 0,
        };

        var columns = rows.Length == 0 ? 1 : rows.Max(CellCount);

        var grid = new TableGrid();
        for (var i = 0; i < Math.Max(columns, 1); i++) grid.AppendChild(new GridColumn());

        var table = new Table(new TableProperties(), grid);
        foreach (var row in rows) table.AppendChild(row);
        return table;
    }

    /// <summary>A table row of one cell per string, each holding a single paragraph.</summary>
    /// <remarks>
    /// The multi-cell sibling of <see cref="Row"/> and <see cref="RowOf"/>, which both build a
    /// ONE-cell row. A dropped or shifted cell is only observable across several columns, so the
    /// content-control tests need rows wider than one.
    /// </remarks>
    public static TableRow RowOfText(params string[] cells)
    {
        var row = new TableRow();
        foreach (string cell in cells) row.AppendChild(new TableCell(P(new Run(new Text(cell)
        {
            Space = SpaceProcessingModeValues.Preserve,
        }))));

        return row;
    }

    /// <summary>Builds a .docx whose body is exactly <paramref name="bodyChildren"/>.</summary>
    public static byte[] Build(params OpenXmlElement[] bodyChildren)
        => Build(null, null, bodyChildren);

    /// <summary>
    /// Builds a .docx with an optional default header and footer wired up through the section
    /// properties, so the header/footer parts are genuinely referenced by the document.
    /// </summary>
    public static byte[] Build(string? headerText, string? footerText, params OpenXmlElement[] bodyChildren)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var body = new Body(bodyChildren.Select(c => c.CloneNode(true)));
            main.Document = new Document(body);

            var sectionProperties = new SectionProperties();

            if (headerText is not null)
            {
                var headerPart = main.AddNewPart<HeaderPart>();
                headerPart.Header = new Header(new Paragraph(
                    new Run(new Text(headerText) { Space = SpaceProcessingModeValues.Preserve })));
                headerPart.Header.Save();
                sectionProperties.AppendChild(new HeaderReference
                {
                    Type = HeaderFooterValues.Default,
                    Id = main.GetIdOfPart(headerPart),
                });
            }

            if (footerText is not null)
            {
                var footerPart = main.AddNewPart<FooterPart>();
                footerPart.Footer = new Footer(new Paragraph(
                    new Run(new Text(footerText) { Space = SpaceProcessingModeValues.Preserve })));
                footerPart.Footer.Save();
                sectionProperties.AppendChild(new FooterReference
                {
                    Type = HeaderFooterValues.Default,
                    Id = main.GetIdOfPart(footerPart),
                });
            }

            if (sectionProperties.HasChildren) body.AppendChild(sectionProperties);
            main.Document.Save();
        }

        return ms.ToArray();
    }

    /// <summary>Opens <paramref name="docx"/> read-only and hands the body to <paramref name="inspect"/>.</summary>
    public static T ReadBody<T>(byte[] docx, Func<Body, T> inspect)
    {
        using var ms = new MemoryStream(docx);
        using var doc = WordprocessingDocument.Open(ms, false);
        return inspect(doc.MainDocumentPart!.Document!.Body!);
    }

    /// <summary>Opens <paramref name="docx"/> read-only and hands the whole package to <paramref name="inspect"/>.</summary>
    public static T Read<T>(byte[] docx, Func<MainDocumentPart, T> inspect)
    {
        using var ms = new MemoryStream(docx);
        using var doc = WordprocessingDocument.Open(ms, false);
        return inspect(doc.MainDocumentPart!);
    }

    /// <summary>
    /// The text a paragraph owns *itself* — i.e. excluding any paragraph nested inside it via a
    /// text box. <c>Paragraph.InnerText</c> would swallow the text box's content too.
    /// </summary>
    public static string OwnText(Paragraph paragraph) => string.Concat(
        paragraph.Descendants<Text>()
                 .Where(t => t.Ancestors<Paragraph>().FirstOrDefault() == paragraph)
                 .Select(t => t.Text));

    /// <summary>The text of each run this paragraph owns directly, in document order.</summary>
    public static string[] OwnRunTexts(Paragraph paragraph) => paragraph
        .Descendants<Run>()
        .Where(r => r.Ancestors<Paragraph>().FirstOrDefault() == paragraph)
        .Select(r => string.Concat(r.Elements<Text>().Select(t => t.Text)))
        .ToArray();

    /// <summary>Schema-validation errors for the whole package (empty means valid).</summary>
    /// <remarks>
    /// <b>The version is explicit, and <c>Office2007</c> — the parameterless default — is the wrong
    /// one here.</b> That schema predates the <c>w14:</c>/<c>w15:</c> extension elements an ordinary
    /// Word-saved document carries (<c>w15:chartTrackingRefBased</c>, <c>w15:docId</c>), so a
    /// validator built for it cannot see an ordering violation involving one and reports clean. Every
    /// validity assertion in this suite went through this helper, which meant no test in the
    /// repository could observe that class of bug — see
    /// <c>DocxEditorFootnoteEndnoteTocTests.AddTableOfContents_InsertsUpdateFieldsInAValidSlot_OnW15ShapedSettings</c>,
    /// which is red under <c>Office2007</c> and green under this.
    /// </remarks>
    public static IReadOnlyList<ValidationErrorInfo> Validate(byte[] docx)
    {
        using var ms = new MemoryStream(docx);
        using var doc = WordprocessingDocument.Open(ms, false);
        return new OpenXmlValidator(FileFormatVersions.Office2013).Validate(doc).ToList();
    }
}
