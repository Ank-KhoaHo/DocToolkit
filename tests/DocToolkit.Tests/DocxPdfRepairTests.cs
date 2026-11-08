using System.IO.Compression;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using static DocToolkit.Tests.DocxFixtures;

namespace DocToolkit.Tests;

/// <summary>
/// The DOCX repairs (up front, or on refusal through DocxPdfRefusalLadder, as each class says) for constructs OfficeIMO 3.5 refuses and Word accepts (A123).
///
/// <b>Every "refused before" test calls the renderer directly on the unrepaired bytes first.</b> A
/// fixture the renderer already accepts would prove nothing about the repair, and the test would
/// pass whether the repair worked or not. Each pair also has an "untouched" test: an up-front
/// repair runs on every conversion, so the one thing it must never do is rewrite a document that
/// has nothing to repair.
/// </summary>
public class DocxPdfRepairTests
{
    // ---- helpers -----------------------------------------------------------------------------------

    private static string Part(byte[] docx, string name)
    {
        using var ms = new MemoryStream(docx);
        using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
        using var reader = new StreamReader(zip.GetEntry(name)!.Open(), Encoding.UTF8);
        return reader.ReadToEnd();
    }

    private static string PartNamed(byte[] docx, string prefix)
    {
        using var ms = new MemoryStream(docx);
        using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
        var entry = zip.Entries.Single(e => e.FullName.StartsWith(prefix, StringComparison.Ordinal));
        using var reader = new StreamReader(entry.Open(), Encoding.UTF8);
        return reader.ReadToEnd();
    }

    // ---- R1: a hyperlink inside a table cell whose run holds no text -------------------------------

    /// <summary>
    /// <b>The link must be inside a table cell.</b> The same link in a body paragraph renders on raw
    /// 3.5.3 (measured), so a body-paragraph fixture would prove nothing - and the repair leaves
    /// such a link alone.
    /// </summary>
    private static byte[] LinkWithBreakInCell(params Run[] linkRuns) => LinkDocument(true, false, linkRuns);

    /// <summary>
    /// A document with one hyperlink - in a table cell or in a body paragraph - made of
    /// <paramref name="linkRuns"/>. A run holding <c>w:drawing</c> with no image in it is replaced by
    /// a real picture, so the fixture is a valid document.
    /// </summary>
    private static byte[] LinkDocument(bool inCell, bool withPicture, params Run[] linkRuns)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var rel = main.AddHyperlinkRelationship(new Uri("https://example.org/x"), true);
            var children = linkRuns.Select(r => (OpenXmlElement)r.CloneNode(true)).ToList();
            if (withPicture)
            {
                var image = main.AddImagePart(ImagePartType.Png);
                using (var png = new MemoryStream(ImageFixtures.Png())) image.FeedData(png);
                children.Insert(children.Count > 0 ? 1 : 0, new Run(Picture(main.GetIdOfPart(image))));
            }

            var link = new Hyperlink(children) { Id = rel.Id };
            OpenXmlElement content = inCell
                ? Tbl(Row(link, R(" after the link")))
                : P(link, R(" after the link"));
            main.Document = new Document(new Body(content));
            main.Document.Save();
        }

        return ms.ToArray();
    }

    private static Drawing Picture(string relationshipId)
    {
        var picture = new PIC.Picture(
            new PIC.NonVisualPictureProperties(
                new PIC.NonVisualDrawingProperties { Id = 0U, Name = "picture" },
                new PIC.NonVisualPictureDrawingProperties()),
            new PIC.BlipFill(new A.Blip { Embed = relationshipId }, new A.Stretch(new A.FillRectangle())),
            new PIC.ShapeProperties(
                new A.Transform2D(new A.Offset { X = 0L, Y = 0L }, new A.Extents { Cx = 990000L, Cy = 792000L }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }));

        return new Drawing(new DW.Inline(
            new DW.Extent { Cx = 990000L, Cy = 792000L },
            new DW.DocProperties { Id = 1U, Name = "picture" },
            new A.Graphic(new A.GraphicData(picture) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })));
    }

    private static byte[] LinkWithBreak() => LinkWithBreakInCell(R("linked words"), new Run(new Break()));

    private static void AssertEmptyTextRefusal(Exception? refusal)
    {
        Assert.IsType<ArgumentException>(refusal);
        Assert.Contains("cannot be empty or whitespace", refusal.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void HyperlinkBlankRunSplit_RefusedBefore_RendersAfter()
    {
        var docx = LinkWithBreak();
        // Raw 3.5.3: ArgumentException "Parameter 'text' cannot be empty or whitespace."
        AssertEmptyTextRefusal(RawRender.Refusal(docx));

        var text = PdfProbe.ExtractText(DocxToPdfConverter.Convert(docx));

        Assert.Contains("linked words", text);
        Assert.Contains("after the link", text);
    }

    [Fact]
    public void HyperlinkBlankRunSplit_ALinkOfOnlySpaces_RefusedBefore_RendersAfter()
    {
        var docx = LinkWithBreakInCell(new Run(new Text(" ") { Space = SpaceProcessingModeValues.Preserve }));
        AssertEmptyTextRefusal(RawRender.Refusal(docx));

        var text = PdfProbe.ExtractText(DocxToPdfConverter.Convert(docx));

        Assert.Contains("after the link", text);
    }

    [Fact]
    public void HyperlinkBlankRunSplit_MovesTheBreakOutOfTheLink()
    {
        var repaired = HyperlinkBlankRunSplit.Apply(LinkWithBreak());

        var xml = Part(repaired, "word/document.xml");
        Assert.Contains("<w:br", xml);
        var link = xml[xml.IndexOf("<w:hyperlink", StringComparison.Ordinal)..xml.IndexOf("</w:hyperlink>", StringComparison.Ordinal)];
        Assert.DoesNotContain("<w:br", link);
        Assert.Contains("linked words", link);
    }

    [Fact]
    public void HyperlinkBlankRunSplit_NothingToDo_ReturnsTheSameArray()
    {
        var docx = LinkWithBreakInCell(R("just a link"));

        Assert.Same(docx, HyperlinkBlankRunSplit.Apply(docx));
    }

    /// <summary>
    /// The shape html-000214 and html-000438 carry, x-ified: a BODY paragraph link whose first run is
    /// only a no-break space and a space. Raw 3.5.3 refuses it, so the repair is not limited to
    /// table cells.
    /// </summary>
    private static byte[] BodyLinkStartingWithBlankRun() =>
        LinkDocument(false, false, new Run(new Text("  ") { Space = SpaceProcessingModeValues.Preserve }), R("x."));

    [Fact]
    public void HyperlinkBlankRunSplit_ABodyParagraphLink_RefusedBefore_RendersAfter()
    {
        var docx = BodyLinkStartingWithBlankRun();
        AssertEmptyTextRefusal(RawRender.Refusal(docx));

        var repaired = HyperlinkBlankRunSplit.Apply(docx);

        Assert.NotSame(docx, repaired);
        Assert.Null(RawRender.Refusal(repaired));
        var text = PdfProbe.ExtractText(DocxToPdfConverter.Convert(docx));
        Assert.Contains("x.", text);
        Assert.Contains("after the link", text);
    }

    [Fact]
    public void HyperlinkBlankRunSplit_IsAStepOfTheLadder_AndNothingToDoGivesNull()
    {
        var refusal = new ArgumentException("Parameter 'text' cannot be empty or whitespace.");
        Assert.True(DocxPdfRefusalLadder.Knows(refusal));

        var docx = BodyLinkStartingWithBlankRun();
        Assert.NotNull(DocxPdfRefusalLadder.NextOrNull(refusal, docx, new HashSet<int>()));

        var plain = LinkDocument(false, false, R("just a link"));
        Assert.Null(DocxPdfRefusalLadder.NextOrNull(refusal, plain, new HashSet<int>()));
    }

    [Fact]
    public void HyperlinkBlankRunSplit_APictureOnlyLinkInACell_IsLeftAlone()
    {
        // Measured: a picture-only link in a cell renders on raw 3.5.3, and unwrapping it would lose the link.
        var docx = LinkDocument(true, true);
        Assert.Null(RawRender.Refusal(docx));

        Assert.Same(docx, HyperlinkBlankRunSplit.Apply(docx));
    }

    [Fact]
    public void HyperlinkBlankRunSplit_ThePictureStaysInsideALink()
    {
        var docx = LinkDocument(true, true, R("before"), new Run(new Break()), R("after"));

        var xml = Part(HyperlinkBlankRunSplit.Apply(docx), "word/document.xml");

        var doc = System.Xml.Linq.XDocument.Parse(xml);
        System.Xml.Linq.XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        Assert.All(doc.Descendants(w + "drawing"), d => Assert.NotNull(d.Ancestors(w + "hyperlink").FirstOrDefault()));
    }

    [Fact]
    public void HyperlinkBlankRunSplit_BothPiecesCarryTheSameTarget()
    {
        var xml = Part(HyperlinkBlankRunSplit.Apply(
            LinkWithBreakInCell(R("before"), new Run(new Break()), R("after"))), "word/document.xml");

        var doc = System.Xml.Linq.XDocument.Parse(xml);
        System.Xml.Linq.XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        System.Xml.Linq.XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var targets = doc.Descendants(w + "hyperlink").Select(h => (string?)h.Attribute(r + "id")).ToList();
        Assert.Equal(2, targets.Count);
        Assert.NotNull(targets[0]);
        Assert.Equal(targets[0], targets[1]);
    }

    [Fact]
    public void HyperlinkBlankRunSplit_AFieldInsideACellLink_IsNotSplit()
    {
        var docx = LinkWithBreakInCell(
            new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
            new Run(new FieldCode(" PAGE ")),
            new Run(new Break()),
            new Run(new FieldChar { FieldCharType = FieldCharValues.End }));

        Assert.Same(docx, HyperlinkBlankRunSplit.Apply(docx));
    }

    [Fact]
    public void DocxPartRewrite_ADoctypeIsNotParsed_AndTheSameArrayComesBack()
    {
        var plain = Build(P(R("A\u009DB")));
        using var ms = new MemoryStream();
        ms.Write(plain);
        ms.Position = 0;
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = zip.GetEntry("word/document.xml")!;
            string xml;
            using (var reader = new StreamReader(entry.Open(), Encoding.UTF8)) xml = reader.ReadToEnd();
            var root = xml.IndexOf("<w:document", StringComparison.Ordinal);
            xml = xml[..root] + "<!DOCTYPE w:document [<!ENTITY boom \"expanded\">]>" + xml[root..];
            // The entity is referenced from the text, so a parser that allows DTDs expands it.
            xml = xml.Replace("B</w:t>", "B&boom;</w:t>", StringComparison.Ordinal);
            entry.Delete();
            using var writer = new StreamWriter(zip.CreateEntry("word/document.xml").Open(), new UTF8Encoding(false));
            writer.Write(xml);
        }

        var docx = ms.ToArray();
        var before = Part(docx, "word/document.xml");
        Assert.Contains("\u009D", before, StringComparison.Ordinal);
        Assert.Contains("&boom;", before, StringComparison.Ordinal);

        // A parser that allows DTDs would expand the entity, strip the control character and return
        // a NEW array - so this fails without DtdProcessing.Prohibit. With it the part is skipped:
        // no throw, no expansion, and the control character is left for the renderer to refuse,
        // which is a better answer than a parse failure from here.
        var result = ControlCharacterStrip.Apply(docx);

        Assert.Same(docx, result);
        Assert.Equal(before, Part(result, "word/document.xml"));
    }

    // ---- R5: control characters --------------------------------------------------------------------

    private static byte[] ControlsInBodyHeaderAndFootnote() =>
        DocxEditor.AddFootnote(
            Build("Head A\u009DB", null, P(R("Body A\u009DB {{n}}"))),
            "{{n}}",
            "Note A\u009DB");

    [Fact]
    public void ControlCharacterStrip_RemovesControlsFromBodyHeaderAndFootnote()
    {
        var repaired = ControlCharacterStrip.Apply(ControlsInBodyHeaderAndFootnote());

        foreach (var prefix in new[] { "word/document.xml", "word/header", "word/footnotes.xml" })
        {
            var xml = prefix.EndsWith(".xml") ? Part(repaired, prefix) : PartNamed(repaired, prefix);
            // Not DoesNotContain(string): its culture-aware comparison treats U+009D as ignorable.
            Assert.False(xml.Contains('\u009D'), $"{prefix} still holds U+009D");
        }

        Assert.Contains("Head AB", PartNamed(repaired, "word/header"));
        Assert.Contains("Note AB", Part(repaired, "word/footnotes.xml"));
    }

    [Fact]
    public void ControlCharacterStrip_KeepsTabAndLineBreakCharacters()
    {
        var docx = Build(P(R("a\u007Fb\tc\u009Fd")));

        var xml = Part(ControlCharacterStrip.Apply(docx), "word/document.xml");

        Assert.Contains(">ab\tcd<", xml);
    }

    [Fact]
    public void ControlCharacterStrip_RefusedBefore_RendersAfter()
    {
        // The render half needs a real font: the refusal is about glyph coverage, which only
        // exists once fonts are supplied, and the dtk-nofonts image has none. The repair-level
        // test above runs everywhere.
        var font = PdfFontOptionsTests.FindFont();
        if (font is null) return;
        var fonts = new PdfFontOptions("Supplied", font);
        var docx = ControlsInBodyHeaderAndFootnote();
        // Raw 3.5.3: PdfTextEncodingPreflightException "Text contains character U+009D ... not covered
        // by any embedded font fallback candidate".
        var refusal = RawRender.Refusal(docx, fonts);
        Assert.NotNull(refusal);
        Assert.Equal("PdfTextEncodingPreflightException", refusal.GetType().Name);
        Assert.Contains("U+009D", refusal.Message, StringComparison.Ordinal);

        var pdf = DocxToPdfConverter.Convert(docx, fonts);

        var text = PdfProbe.ExtractText(pdf);
        Assert.Contains("Body AB", text);
        Assert.Contains("Head AB", text);
        Assert.Contains("Note AB", text);
    }

    [Fact]
    public void ControlCharacterStrip_NothingToDo_ReturnsTheSameArray()
    {
        var docx = Build("Header", "Footer", P(R("Plain text\twith a tab")));

        Assert.Same(docx, ControlCharacterStrip.Apply(docx));
    }

    // ---- R7: a percentage width of a few fiftieths of a percent ------------------------------------

    /// <summary>
    /// A 1% table (<c>tblW w="50" type="pct"</c>) with cell spacing, nested in a cell of a 100% table.
    /// </summary>
    /// <remarks>
    /// <b>Minimal shape, bisected from a real document (html-000079) down to this.</b> Both parts are
    /// needed: the 1% width alone renders, and so does cell spacing alone. Nesting is needed too, as
    /// a top-level 1% table with spacing renders. Cell margins do not matter - the real document had
    /// them zeroed, this fixture does not.
    /// </remarks>
    private static byte[] TinyPercentSpacedTableInCell()
    {
        var inner = Tbl(Row(R("tiny percent table")));
        var properties = inner.GetFirstChild<TableProperties>()!;
        properties.TableWidth = new TableWidth { Width = "50", Type = TableWidthUnitValues.Pct };
        properties.TableCellSpacing = new TableCellSpacing { Width = "90", Type = TableWidthUnitValues.Dxa };

        var outer = Tbl(new TableRow(new TableCell(P(R("before")), inner, P(R("after")))));
        outer.GetFirstChild<TableProperties>()!.TableWidth =
            new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct };
        return Build(outer);
    }

    [Fact]
    public void TinyPercentWidthRepair_RefusedBefore_RendersAfter()
    {
        var docx = TinyPercentSpacedTableInCell();
        // Raw 3.5.3: ArgumentException "Table preferred width must be a positive finite value.
        // (Parameter 'PreferredWidth')".
        var refusal = RawRender.Refusal(docx);
        Assert.IsType<ArgumentException>(refusal);
        Assert.Contains("preferred width must be a positive finite value", refusal.Message, StringComparison.Ordinal);

        // The repair alone is enough here; a document that ALSO trips the cell-padding refusal (the
        // real html-000079 does) needs the padding ladder as well.
        Assert.Null(RawRender.Refusal(TinyPercentWidthRepair.Apply(docx)));
        var text = PdfProbe.ExtractText(DocxToPdfConverter.Convert(docx));
        Assert.Contains("tiny percent table", text);
        Assert.Contains("after", text);
    }

    [Fact]
    public void TinyPercentWidthRepair_TurnsTheTableWidthAuto()
    {
        var xml = Part(TinyPercentWidthRepair.Apply(TinyPercentSpacedTableInCell()), "word/document.xml");

        var tblW = xml[xml.IndexOf("<w:tblW", StringComparison.Ordinal)..];
        tblW = tblW[..tblW.IndexOf("/>", StringComparison.Ordinal)];
        Assert.Contains("w:type=\"pct\"", tblW); // the 100% outer table is first and untouched
        var inner = xml[(xml.IndexOf("<w:tblW", StringComparison.Ordinal) + 1)..];
        inner = inner[inner.IndexOf("<w:tblW", StringComparison.Ordinal)..];
        inner = inner[..inner.IndexOf("/>", StringComparison.Ordinal)];
        Assert.Contains("w:w=\"0\"", inner);
        Assert.Contains("w:type=\"auto\"", inner);
    }

    [Fact]
    public void TinyPercentWidthRepair_AlsoFixesACell()
    {
        var table = Tbl(new TableRow(
            new TableCell(new TableCellProperties(new TableCellWidth { Width = "50", Type = TableWidthUnitValues.Pct }), P(R("tiny percent cell"))),
            new TableCell(P(R("the rest")))));

        var xml = Part(TinyPercentWidthRepair.Apply(Build(table)), "word/document.xml");

        var tcW = xml[xml.IndexOf("<w:tcW", StringComparison.Ordinal)..];
        tcW = tcW[..tcW.IndexOf("/>", StringComparison.Ordinal)];
        Assert.Contains("w:w=\"0\"", tcW);
        Assert.Contains("w:type=\"auto\"", tcW);
    }

    [Fact]
    public void TinyPercentWidthRepair_NothingToDo_ReturnsTheSameArray()
    {
        var table = Tbl(Row(R("half width table")));
        table.GetFirstChild<TableProperties>()!.TableWidth =
            new TableWidth { Width = "2500", Type = TableWidthUnitValues.Pct };
        var docx = Build(table);

        Assert.Same(docx, TinyPercentWidthRepair.Apply(docx));
    }

    // ---- R6: a top margin under 360 twips under a header --------------------------------------------

    /// <summary>
    /// <paramref name="docx"/> with its last section given these margins (twips) and an A4 page. The
    /// fixture builder writes the header and footer references but no <c>w:pgMar</c>.
    /// </summary>
    private static byte[] WithMargins(byte[] docx, int top, uint header, int bottom, uint footer)
    {
        using var ms = new MemoryStream();
        ms.Write(docx);
        using (var doc = WordprocessingDocument.Open(ms, true))
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            var section = body.Elements<SectionProperties>().LastOrDefault() ?? body.AppendChild(new SectionProperties());
            section.RemoveAllChildren<PageMargin>();
            section.AppendChild(new PageSize { Width = 11906U, Height = 16838U });
            section.AppendChild(new PageMargin
            {
                Top = top,
                Bottom = bottom,
                Header = header,
                Footer = footer,
                Left = 1440U,
                Right = 1440U,
                Gutter = 0U,
            });
            doc.MainDocumentPart.Document.Save();
        }

        return ms.ToArray();
    }

    private static byte[] HeaderPage(int top, uint header = 0, string? headerText = "Header words") =>
        WithMargins(Build(headerText, null, P(R("Body words"))), top, header, 1440, 720);

    private static string MarginAttribute(byte[] docx, string name)
    {
        var xml = Part(docx, "word/document.xml");
        var margin = xml[xml.IndexOf("<w:pgMar", StringComparison.Ordinal)..];
        margin = margin[..margin.IndexOf("/>", StringComparison.Ordinal)];
        var start = margin.IndexOf($"w:{name}=\"", StringComparison.Ordinal) + name.Length + 4;
        return margin[start..margin.IndexOf('"', start)];
    }

    [Fact]
    public void HeaderOffsetRepair_TopMarginUnder360WithAHeader_RefusedBefore_IsRaisedTo360()
    {
        var docx = HeaderPage(288);
        // Raw 3.5.3: ArgumentException "PDF header offset must not exceed the top margin when header
        // content is enabled", with a header distance of 0.
        var refusal = RawRender.Refusal(docx);
        Assert.IsType<ArgumentException>(refusal);
        Assert.Contains("header offset must not exceed the top margin", refusal.Message, StringComparison.Ordinal);

        var repaired = HeaderOffsetRepair.Apply(docx);

        Assert.Equal("360", MarginAttribute(repaired, "top"));
        Assert.Null(RawRender.Refusal(repaired));
        var text = PdfProbe.ExtractText(DocxToPdfConverter.Convert(docx));
        Assert.Contains("Body words", text);
        Assert.Contains("Header words", text);
    }

    /// <summary>
    /// doc-000111's shape: the header is declared on the FIRST section (top 576, fine), and the last
    /// section declares none, inherits it, and has a top margin of 288.
    /// </summary>
    private static byte[] HeaderOnFirstSectionOnly_LastSectionTop288()
    {
        var docx = HeaderPage(576);
        using var ms = new MemoryStream();
        ms.Write(docx);
        using (var doc = WordprocessingDocument.Open(ms, true))
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            var last = body.Elements<SectionProperties>().Last();
            var first = new Paragraph(
                new ParagraphProperties((SectionProperties)last.CloneNode(true)),
                new Run(new Text("First section words")));
            body.InsertAt(first, 0);
            last.RemoveAllChildren<HeaderReference>();
            last.GetFirstChild<PageMargin>()!.Top = 288;
            doc.MainDocumentPart.Document.Save();
        }

        return ms.ToArray();
    }

    [Fact]
    public void HeaderOffsetRepair_ASectionInheritingTheHeader_RefusedBefore_IsRaised()
    {
        var docx = HeaderOnFirstSectionOnly_LastSectionTop288();
        var refusal = RawRender.Refusal(docx);
        Assert.IsType<ArgumentException>(refusal);
        Assert.Contains("header offset must not exceed the top margin", refusal.Message, StringComparison.Ordinal);

        var repaired = HeaderOffsetRepair.Apply(docx);

        Assert.NotSame(docx, repaired);
        Assert.Null(RawRender.Refusal(repaired));
        var text = PdfProbe.ExtractText(DocxToPdfConverter.Convert(docx));
        Assert.Contains("First section words", text);
        Assert.Contains("Body words", text);
    }

    [Fact]
    public void HeaderOffsetRepair_AHeaderOfOnlyALegacyPageNumber_RefusedBefore_IsRaised()
    {
        // doc-000111's headers hold only <w:pgNum/>, which a document converted from .doc carries.
        var docx = HeaderPage(288, 0, headerText: "");
        using var ms = new MemoryStream();
        ms.Write(docx);
        using (var doc = WordprocessingDocument.Open(ms, true))
        {
            var header = doc.MainDocumentPart!.HeaderParts.Single();
            header.Header!.RemoveAllChildren();
            header.Header.AppendChild(new Paragraph(new Run(new PageNumber())));
            header.Header.Save();
        }

        var legacy = ms.ToArray();
        Assert.Contains("header offset must not exceed the top margin", RawRender.Refusal(legacy)!.Message, StringComparison.Ordinal);

        Assert.Equal("360", MarginAttribute(HeaderOffsetRepair.Apply(legacy), "top"));
        Assert.Contains("Body words", PdfProbe.ExtractText(DocxToPdfConverter.Convert(legacy)));
    }

    [Fact]
    public void HeaderOffsetRepair_IsAStepOfTheLadder_AndNothingToDoGivesNull()
    {
        var refusal = new ArgumentException("PDF header offset must not exceed the top margin when header content is enabled.");
        Assert.True(DocxPdfRefusalLadder.Knows(refusal));

        Assert.NotNull(DocxPdfRefusalLadder.NextOrNull(refusal, HeaderPage(288), new HashSet<int>()));
        Assert.Null(DocxPdfRefusalLadder.NextOrNull(refusal, HeaderPage(576), new HashSet<int>()));
    }

    [Fact]
    public void HeaderOffsetRepair_TheFloorIs360_359IsRefused_360IsNot()
    {
        Assert.Contains("header offset", RawRender.Refusal(HeaderPage(359))!.Message, StringComparison.Ordinal);

        var accepted = HeaderPage(360);

        Assert.Null(RawRender.Refusal(accepted));
        Assert.Same(accepted, HeaderOffsetRepair.Apply(accepted));
    }

    [Fact]
    public void HeaderOffsetRepair_ATopMarginOfZero_IsRaisedToo()
    {
        var docx = HeaderPage(0);
        Assert.Contains("header offset", RawRender.Refusal(docx)!.Message, StringComparison.Ordinal);

        Assert.Equal("360", MarginAttribute(HeaderOffsetRepair.Apply(docx), "top"));
    }

    [Fact]
    public void HeaderOffsetRepair_NoHeaderNoChange()
    {
        // A small top margin is fine without a header: measured, raw 3.5.3 renders it.
        var docx = WithMargins(Build(P(R("Body words"))), 288, 0, 288, 0);
        Assert.Null(RawRender.Refusal(docx));

        Assert.Same(docx, HeaderOffsetRepair.Apply(docx));
    }

    [Fact]
    public void HeaderOffsetRepair_AnEmptyHeader_IsLeftAlone()
    {
        // Measured: "header content is enabled" is about content. A header holding no text renders
        // on raw 3.5.3 with a top margin of 288, so it must not be edited.
        var docx = HeaderPage(288, 0, headerText: "");
        Assert.Null(RawRender.Refusal(docx));

        Assert.Same(docx, HeaderOffsetRepair.Apply(docx));
    }

    [Fact]
    public void HeaderOffsetRepair_AHeaderDistanceAboveTheTopMargin_IsLeftAlone()
    {
        // The plan expected 3.5 to refuse this; measured on raw 3.5.3 it renders (even 8000 over
        // 1440), so editing it would change a document that renders today.
        var docx = HeaderPage(576, 720);
        Assert.Null(RawRender.Refusal(docx));

        Assert.Same(docx, HeaderOffsetRepair.Apply(docx));
        var text = PdfProbe.ExtractText(DocxToPdfConverter.Convert(docx));
        Assert.Contains("Header words", text);
        Assert.Contains("Body words", text);
    }

    [Fact]
    public void HeaderOffsetRepair_ANegativeTopMargin_IsLeftAlone()
    {
        // 3.5 refuses a negative margin at any size and for another reason, so -360 would not help.
        var docx = HeaderPage(-288);
        var refusal = RawRender.Refusal(docx);
        Assert.IsType<ArgumentOutOfRangeException>(refusal);
        Assert.Contains("must be a finite non-negative number", refusal.Message, StringComparison.Ordinal);

        Assert.Same(docx, HeaderOffsetRepair.Apply(docx));
    }

    // ---- R9: a fixed-width table wider than the text area ------------------------------------------

    private const double A4WidthPoints = 595.3;

    /// <summary>
    /// A one-row, two-column table on an A4 page with 1440 twip side margins (text area 9026). The
    /// second cell is right-aligned, so its text sits at the right edge of the table: a table drawn
    /// off the page puts that text off the page.
    /// </summary>
    private static Table TwoColumnTable(bool fixedLayout, string? widthType, int? width, int[] grid, int?[] cellWidths)
    {
        var properties = new TableProperties();
        if (widthType is not null)
            properties.TableWidth = new TableWidth { Width = (width ?? 0).ToString(), Type = widthType == "pct" ? TableWidthUnitValues.Pct : TableWidthUnitValues.Dxa };
        if (fixedLayout) properties.TableLayout = new TableLayout { Type = TableLayoutValues.Fixed };

        var tableGrid = new TableGrid();
        foreach (var column in grid) tableGrid.AppendChild(new GridColumn { Width = column.ToString() });

        var row = new TableRow();
        for (var i = 0; i < cellWidths.Length; i++)
        {
            var paragraph = P(R(i == 0 ? "LEFT" : "RIGHT"));
            if (i == cellWidths.Length - 1)
                paragraph.PrependChild(new ParagraphProperties(new Justification { Val = JustificationValues.Right }));
            var cell = new TableCell(paragraph);
            if (cellWidths[i] is { } w)
                cell.PrependChild(new TableCellProperties(new TableCellWidth { Width = w.ToString(), Type = TableWidthUnitValues.Dxa }));
            row.AppendChild(cell);
        }

        return new Table(properties, tableGrid, row);
    }

    private static byte[] WideTablePage(params OpenXmlElement[] body) =>
        WithMargins(Build(body), 1440, 720, 1440, 720);

    /// <summary>The wide table of the plan: fixed layout, 12000 twips over a 9026 twip text area.</summary>
    private static byte[] TwelveThousandTwipTable() =>
        WideTablePage(TwoColumnTable(true, "dxa", 12000, [8000, 4000], [8000, 4000]));

    private static double XOf(byte[] pdf, string text) =>
        PdfProbe.TextRuns(pdf).Single(r => r.Text.Contains(text, StringComparison.Ordinal)).X;

    private static List<int> TableWidths(byte[] docx, string element)
    {
        var xml = System.Xml.Linq.XDocument.Parse(Part(docx, "word/document.xml"));
        System.Xml.Linq.XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var table = xml.Descendants(w + "tbl").First();
        return table.Descendants(w + element)
            .Where(e => element == "gridCol" || (string?)e.Attribute(w + "type") == "dxa")
            .Select(e => int.Parse((string)e.Attribute(w + "w")!)).ToList();
    }

    [Fact]
    public void WideTableScale_AFixedTableWiderThanTheTextAreaIsScaledToIt()
    {
        var docx = TwelveThousandTwipTable();
        // Not a refusal: raw 3.5.3 renders it, and draws the right-aligned text outside the page.
        Assert.Null(RawRender.Refusal(docx));
        Assert.True(XOf(RawRender.Pdf(docx), "RIGHT") > A4WidthPoints);

        var repaired = WideTableScale.Apply(docx);

        var grid = TableWidths(repaired, "gridCol");
        Assert.True(grid.Sum() <= 9026, $"grid {string.Join("+", grid)}");
        Assert.True(TableWidths(repaired, "tcW").Sum() <= 9026);
        Assert.True(TableWidths(repaired, "tblW").Single() <= 9026);
        // The ratio of the columns is kept (8000:4000 = 2:1), to within truncation.
        Assert.InRange((double)grid[0] / grid[1], 1.99, 2.01);

        Assert.True(XOf(DocxToPdfConverter.Convert(docx), "RIGHT") < A4WidthPoints);
    }

    [Fact]
    public void WideTableScale_ADxaWidthAloneIsEnough_WhateverTheGrid()
    {
        // Measured: a dxa table width of 12000 over a grid of 3000 + 3000 is drawn 12000 wide.
        var docx = WideTablePage(TwoColumnTable(false, "dxa", 12000, [3000, 3000], [3000, 3000]));
        Assert.True(XOf(RawRender.Pdf(docx), "RIGHT") > A4WidthPoints);

        Assert.True(TableWidths(WideTableScale.Apply(docx), "tblW").Single() <= 9026);
    }

    [Fact]
    public void WideTableScale_AFixedTableWithNoWidthButAWideGrid_IsScaled()
    {
        var docx = WideTablePage(TwoColumnTable(true, null, null, [6000, 6000], [6000, 6000]));
        Assert.True(XOf(RawRender.Pdf(docx), "RIGHT") > A4WidthPoints);

        Assert.True(TableWidths(WideTableScale.Apply(docx), "gridCol").Sum() <= 9026);
    }

    [Fact]
    public void WideTableScale_APercentTableIsUntouched()
    {
        // Measured: a 100% table with a wide grid renders inside the page, with or without a fixed layout.
        foreach (var fixedLayout in new[] { false, true })
        {
            var docx = WideTablePage(TwoColumnTable(fixedLayout, "pct", 5000, [6000, 6000], [6000, 6000]));
            Assert.True(XOf(RawRender.Pdf(docx), "RIGHT") < A4WidthPoints);

            Assert.Same(docx, WideTableScale.Apply(docx));
        }
    }

    [Fact]
    public void WideTableScale_AnAutofitTableWithAWideGridOnly_IsUntouched()
    {
        var docx = WideTablePage(TwoColumnTable(false, null, null, [6000, 6000], [6000, 6000]));
        Assert.True(XOf(RawRender.Pdf(docx), "RIGHT") < A4WidthPoints);

        Assert.Same(docx, WideTableScale.Apply(docx));
    }

    [Fact]
    public void WideTableScale_ADxaWidthThatFitsOverAWiderGrid_IsUntouched()
    {
        var docx = WideTablePage(TwoColumnTable(true, "dxa", 9000, [6000, 6000], [6000, 6000]));
        Assert.True(XOf(RawRender.Pdf(docx), "RIGHT") < A4WidthPoints);

        Assert.Same(docx, WideTableScale.Apply(docx));
    }

    [Fact]
    public void WideTableScale_ANestedTableIsUntouched()
    {
        // Only top-level body tables are scaled. The outer table fits; the wide one is inside a cell.
        var inner = TwoColumnTable(true, "dxa", 12000, [6000, 6000], [6000, 6000]);
        var outer = new Table(new TableProperties(), new TableGrid(new GridColumn { Width = "9000" }),
            new TableRow(new TableCell(P(R("outer")), inner, P())));
        var docx = WideTablePage(outer);

        Assert.Same(docx, WideTableScale.Apply(docx));
    }

    [Fact]
    public void WideTableScale_OnlyTheOuterTableChanges_WhenBothAreWide()
    {
        var inner = TwoColumnTable(true, "dxa", 12000, [6000, 6000], [6000, 6000]);
        var outer = new Table(
            new TableProperties(new TableWidth { Width = "12000", Type = TableWidthUnitValues.Dxa }),
            new TableGrid(new GridColumn { Width = "12000" }),
            new TableRow(new TableCell(new TableCellProperties(new TableCellWidth { Width = "12000", Type = TableWidthUnitValues.Dxa }),
                P(R("outer")), inner, P())));
        var docx = WideTablePage(outer);

        var xml = System.Xml.Linq.XDocument.Parse(Part(WideTableScale.Apply(docx), "word/document.xml"));
        System.Xml.Linq.XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var tables = xml.Descendants(w + "tbl").ToList();
        Assert.Equal("9026", (string?)tables[0].Element(w + "tblPr")!.Element(w + "tblW")!.Attribute(w + "w"));
        Assert.Equal("12000", (string?)tables[1].Element(w + "tblPr")!.Element(w + "tblW")!.Attribute(w + "w"));
    }

    [Fact]
    public void WideTableScale_NarrowTable_ReturnsTheSameArray()
    {
        var docx = WideTablePage(TwoColumnTable(true, "dxa", 8000, [4000, 4000], [4000, 4000]));
        Assert.True(XOf(RawRender.Pdf(docx), "RIGHT") < A4WidthPoints);

        Assert.Same(docx, WideTableScale.Apply(docx));
    }

    [Fact]
    public void WideTableScale_NoSectionProperties_ReturnsTheSameArray()
    {
        // Without a page size there is no text area to scale to.
        var docx = Build(TwoColumnTable(true, "dxa", 12000, [6000, 6000], [6000, 6000]));

        Assert.Same(docx, WideTableScale.Apply(docx));
    }

    // ---- R9 (final review): each table is measured against ITS OWN section -------------------------

    /// <summary>
    /// A body of <paramref name="table"/>, then a paragraph whose <c>w:pPr/w:sectPr</c> closes the first
    /// section (width, height, side margin in twips), then a last paragraph and the body's own last
    /// <c>w:sectPr</c> (the second section). In Word only the LAST section's properties are a direct
    /// child of <c>w:body</c>.
    /// </summary>
    private static byte[] TwoSections(Table table, (uint W, uint H, uint Margin) first, (uint W, uint H, uint Margin) last)
    {
        static SectionProperties Section((uint W, uint H, uint Margin) s) => new(
            new PageSize { Width = s.W, Height = s.H, Orient = s.W > s.H ? PageOrientationValues.Landscape : PageOrientationValues.Portrait },
            new PageMargin { Top = 1440, Bottom = 1440, Left = s.Margin, Right = s.Margin, Header = 720U, Footer = 720U, Gutter = 0U });

        var docx = Build(table, new Paragraph(new ParagraphProperties(Section(first))), P(R("second section")));

        using var ms = new MemoryStream();
        ms.Write(docx);
        using (var doc = WordprocessingDocument.Open(ms, true))
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            foreach (var old in body.Elements<SectionProperties>().ToList()) old.Remove();
            body.AppendChild(Section(last));
            doc.MainDocumentPart.Document.Save();
        }

        return ms.ToArray();
    }

    [Fact]
    public void AWideTableInALandscapeSectionBeforeAPortraitLastSection_IsUntouched()
    {
        // Landscape 15840 wide with 1080 side margins: a text area of 13680. The 13000 table fits it.
        // The body's own w:sectPr is the portrait one (text area 9026), which is NOT this table's.
        var docx = TwoSections(
            TwoColumnTable(true, "dxa", 13000, [6500, 6500], [6500, 6500]),
            (15840, 12240, 1080),
            (11906, 16838, 1440));

        Assert.Same(docx, WideTableScale.Apply(docx));
    }

    [Fact]
    public void ATableTooWideForAPortraitSectionBeforeALandscapeLastSection_IsScaled()
    {
        // The mirror: the table is in the portrait (9026) section, the last section is landscape.
        var docx = TwoSections(
            TwoColumnTable(true, "dxa", 12000, [6000, 6000], [6000, 6000]),
            (11906, 16838, 1440),
            (15840, 12240, 1440));

        var repaired = WideTableScale.Apply(docx);

        Assert.NotSame(docx, repaired);
        Assert.True(TableWidths(repaired, "tblW").Single() <= 9026);
        Assert.True(TableWidths(repaired, "gridCol").Sum() <= 9026);
    }

    [Fact]
    public void WideTableScale_ATableAfterTheLastParagraphSectionUsesTheBodySection()
    {
        // Two tables: one in the portrait section, one after it (so in the landscape last section).
        // 12000 fits the landscape text area (12960) and not the portrait one.
        var docx = TwoSections(
            TwoColumnTable(true, "dxa", 12000, [6000, 6000], [6000, 6000]),
            (11906, 16838, 1440),
            (15840, 12240, 1440));
        using var ms = new MemoryStream();
        ms.Write(docx);
        using (var doc = WordprocessingDocument.Open(ms, true))
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            body.InsertBefore(
                TwoColumnTable(true, "dxa", 12000, [6000, 6000], [6000, 6000]),
                body.Elements<SectionProperties>().Last());
            doc.MainDocumentPart.Document.Save();
        }

        var xml = System.Xml.Linq.XDocument.Parse(Part(WideTableScale.Apply(ms.ToArray()), "word/document.xml"));
        System.Xml.Linq.XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var widths = xml.Root!.Element(w + "body")!.Elements(w + "tbl")
            .Select(t => (string?)t.Element(w + "tblPr")!.Element(w + "tblW")!.Attribute(w + "w")).ToList();
        Assert.Equal(new[] { "9026", "12000" }, widths);
    }

    // ---- R9 (final review): arithmetic that cannot throw or wrap -----------------------------------

    private static byte[] HugeGridTable() =>
        WideTablePage(TwoColumnTable(true, null, null, [2000000000, 2000000000], [null, null]));

    [Fact]
    public void WideTableScale_AGridWhoseSumOverflowsAnInt_DoesNotThrow_AndScalesToTheTextArea()
    {
        var docx = HugeGridTable();

        var repaired = WideTableScale.Apply(docx);

        var grid = TableWidths(repaired, "gridCol");
        Assert.All(grid, g => Assert.InRange(g, 1, 9026));
        Assert.True(grid.Sum(g => (long)g) <= 9026);
    }

    [Fact]
    public void WideTableScale_AGridWhoseSumOverflowsAnInt_ConvertNeverThrowsOverflow()
    {
        var thrown = Record.Exception(() => DocxToPdfConverter.Convert(HugeGridTable()));

        Assert.True(thrown is null or DocumentConversionException, $"escaped: {thrown}");
    }

    [Fact]
    public void WideTableScale_CellWidthsWhoseSumOverflowsAnInt_DoNotThrow()
    {
        var docx = WideTablePage(TwoColumnTable(true, null, null, [6000, 6000], [2000000000, 2000000000]));

        var thrown = Record.Exception(() => WideTableScale.Apply(docx));

        Assert.Null(thrown);
    }

    // ---- an unreadable part is skipped, not thrown -------------------------------------------------

    /// <summary>
    /// <paramref name="docx"/> plus an unreferenced <c>word/footer3.xml</c> whose Deflate data is
    /// garbage (a block header of type 3, which Deflate reserves): reading it throws
    /// <see cref="InvalidDataException"/>.
    /// </summary>
    private static byte[] WithCorruptFooter(byte[] docx)
    {
        using var ms = new MemoryStream();
        ms.Write(docx);
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Update, leaveOpen: true))
        {
            using var writer = new StreamWriter(zip.CreateEntry("word/footer3.xml", CompressionLevel.Optimal).Open(), new UTF8Encoding(false));
            writer.Write("<w:ftr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:p><w:r><w:t>footer text long enough to deflate</w:t></w:r></w:p></w:ftr>");
        }

        var bytes = ms.ToArray();
        var name = "word/footer3.xml"u8;
        var at = bytes.AsSpan().IndexOf(name);
        Assert.True(at >= 30, "local header not found");
        var extra = BitConverter.ToUInt16(bytes, at - 30 + 28);
        bytes[at + name.Length + extra] = 0xFF;
        return bytes;
    }

    [Fact]
    public void ControlCharacterStrip_AnUnreadablePartIsSkipped_NotThrown()
    {
        var docx = WithCorruptFooter(Build("Header", "Footer", P(R("Plain text"))));
        // The fixture is a real one: reading the part does throw.
        using (var zip = new ZipArchive(new MemoryStream(docx), ZipArchiveMode.Read))
            Assert.Throws<InvalidDataException>(() => new StreamReader(zip.GetEntry("word/footer3.xml")!.Open()).ReadToEnd());

        Assert.Same(docx, ControlCharacterStrip.Apply(docx));
    }

    [Fact]
    public void ControlCharacterStrip_AnUnreadablePart_DoesNotStopTheRepairOfTheOthers()
    {
        var docx = WithCorruptFooter(Build(P(R("Body A\u009DB"))));

        var repaired = ControlCharacterStrip.Apply(docx);

        Assert.False(Part(repaired, "word/document.xml").Contains('\u009D'));
    }
}
