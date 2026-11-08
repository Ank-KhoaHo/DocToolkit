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
/// The up-front DOCX repairs for constructs OfficeIMO 3.5 refuses and Word accepts (A123).
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

    /// <summary>What the renderer says about these bytes as they are, or null if it renders them.</summary>
    private static Exception? RawRefusal(byte[] docx, PdfFontOptions? fonts = null)
    {
        var buffer = new MemoryStream();
        buffer.Write(docx, 0, docx.Length);
        buffer.Position = 0;
        using var input = buffer;
        using var word = WordDocument.Load(input);
        try
        {
            word.ToPdfBytes(new WordToPdfOptions { PdfOptions = fonts?.ToPdfOptions() });
            return null;
        }
        catch (Exception ex)
        {
            return ex;
        }
    }

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
        AssertEmptyTextRefusal(RawRefusal(docx));

        var text = PdfProbe.ExtractText(DocxToPdfConverter.Convert(docx));

        Assert.Contains("linked words", text);
        Assert.Contains("after the link", text);
    }

    [Fact]
    public void HyperlinkBlankRunSplit_ALinkOfOnlySpaces_RefusedBefore_RendersAfter()
    {
        var docx = LinkWithBreakInCell(new Run(new Text(" ") { Space = SpaceProcessingModeValues.Preserve }));
        AssertEmptyTextRefusal(RawRefusal(docx));

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

    [Fact]
    public void HyperlinkBlankRunSplit_ABodyParagraphLinkWithABreak_ReturnsTheSameArray()
    {
        // Renders on raw 3.5.3, so it must not be rewritten.
        var docx = LinkDocument(false, false, R("linked words"), new Run(new Break()));
        Assert.Null(RawRefusal(docx));

        Assert.Same(docx, HyperlinkBlankRunSplit.Apply(docx));
    }

    [Fact]
    public void HyperlinkBlankRunSplit_APictureOnlyLinkInACell_IsLeftAlone()
    {
        // Measured: a picture-only link in a cell renders on raw 3.5.3, and unwrapping it would lose the link.
        var docx = LinkDocument(true, true);
        Assert.Null(RawRefusal(docx));

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
        var plain = Build(P(R("AB")));
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
            entry.Delete();
            using var writer = new StreamWriter(zip.CreateEntry("word/document.xml").Open(), new UTF8Encoding(false));
            writer.Write(xml);
        }

        var docx = ms.ToArray();

        // Skipped, not thrown and not expanded: the control character is still there to be refused
        // by the renderer, which is the better answer than a parse failure from here.
        Assert.Same(docx, ControlCharacterStrip.Apply(docx));
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
        var refusal = RawRefusal(docx, fonts);
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
        var refusal = RawRefusal(docx);
        Assert.IsType<ArgumentException>(refusal);
        Assert.Contains("preferred width must be a positive finite value", refusal.Message, StringComparison.Ordinal);

        // The repair alone is enough here; a document that ALSO trips the cell-padding refusal (the
        // real html-000079 does) needs the padding ladder as well.
        Assert.Null(RawRefusal(TinyPercentWidthRepair.Apply(docx)));
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
}
