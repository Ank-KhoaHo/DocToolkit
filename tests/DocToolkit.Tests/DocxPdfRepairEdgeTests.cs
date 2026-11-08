using System.IO.Compression;
using System.Text;
using System.Xml.Linq;

namespace DocToolkit.Tests;

/// <summary>
/// The edges of the DOCX repairs that <see cref="DocxPdfRepairTests"/> does not reach: a part that
/// is missing, malformed or only partly filled in, and the branch of each repair that decides
/// "nothing to do" (A123).
///
/// <b>These documents are written as XML, not built with the SDK, on purpose.</b> The SDK writes
/// every attribute a schema wants, and the edges here are the ones a real producer leaves out: no
/// <c>w:pgMar</c>, a width that is not a number, a relationship with no <c>Id</c>. A repair that
/// reaches the same array it was handed must hand it back as the SAME array, so each "left alone"
/// test asserts that with <see cref="Assert.Same"/>, and each "repaired" test asserts the value the
/// repair wrote.
/// </summary>
public class DocxPdfRepairEdgeTests
{
    private const string WNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string RNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string HeaderType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
    private const string Page = "<w:pgSz w:w=\"11906\" w:h=\"16838\"/><w:pgMar w:left=\"1440\" w:right=\"1440\"/>";

    // ---- helpers -----------------------------------------------------------------------------------

    private static byte[] Pkg(params (string Name, string Xml)[] parts)
    {
        using var ms = new MemoryStream();
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var (name, xml) in parts)
            {
                using var writer = new StreamWriter(zip.CreateEntry(name, CompressionLevel.Optimal).Open(), new UTF8Encoding(false));
                writer.Write(xml);
            }
        }

        return ms.ToArray();
    }

    private static string Doc(string body) =>
        $"<w:document xmlns:w=\"{WNs}\" xmlns:r=\"{RNs}\"><w:body>{body}</w:body></w:document>";

    private static byte[] DocOnly(string body) => Pkg(("word/document.xml", Doc(body)));

    private static string Part(byte[] docx, string name)
    {
        using var zip = new ZipArchive(new MemoryStream(docx), ZipArchiveMode.Read);
        using var reader = new StreamReader(zip.GetEntry(name)!.Open(), Encoding.UTF8);
        return reader.ReadToEnd();
    }

    private static XNamespace W => WNs;

    // ---- ControlCharacterStrip ---------------------------------------------------------------------

    [Fact]
    public void ControlCharacterStrip_RemovesC1AndDelete_KeepsWhitespaceAndLatin1()
    {
        // U+007F..U+009F are controls; U+00A0 and the whitespace characters XML allows are not. The
        // carriage return is a character reference, because a literal one is normalised to a line
        // feed by the parser before the repair ever sees it.
        var docx = DocOnly("<w:p><w:r><w:t xml:space=\"preserve\">a\tb\nc&#13;d\u007Fe\u0080f\u009Fg\u00A0h</w:t></w:r></w:p>");

        var text = XDocument.Parse(Part(ControlCharacterStrip.Apply(docx), "word/document.xml")).Descendants(W + "t").Single().Value;

        // The carriage return is kept as whitespace; it reads back as a line feed because the test
        // parses the written part, and XML normalises line ends on parse.
        Assert.Equal("a\tb\nc\ndefg\u00A0h", text);
    }

    // ---- DrawingRunSplit ---------------------------------------------------------------------------

    [Fact]
    public void DrawingRunSplit_ContentAfterADrawing_StaysWithThatDrawing()
    {
        var docx = DocOnly(
            "<w:p><w:r><w:rPr><w:b/></w:rPr><w:drawing><a/></w:drawing><w:t>after</w:t><w:drawing><b/></w:drawing></w:r></w:p>");

        var runs = XDocument.Parse(Part(DrawingRunSplit.Apply(docx), "word/document.xml")).Descendants(W + "r").ToList();

        Assert.Equal(2, runs.Count);
        Assert.Equal(["rPr", "drawing", "t"], runs[0].Elements().Select(e => e.Name.LocalName));
        Assert.Equal(["rPr", "drawing"], runs[1].Elements().Select(e => e.Name.LocalName));
        Assert.Equal("after", runs[0].Element(W + "t")!.Value);
    }

    // ---- EmptyCellMinWidth -------------------------------------------------------------------------

    [Theory]
    [InlineData("<w:tcW w:w=\"0\" w:type=\"dxa\"/>")]
    [InlineData("<w:tcW w:w=\"wide\" w:type=\"dxa\"/>")]
    [InlineData("<w:tcW w:type=\"dxa\"/>")]
    [InlineData("<w:tcW w:w=\"900\" w:type=\"auto\"/>")]
    [InlineData("<w:tcW w:w=\"900\"/>")]
    public void EmptyCellMinWidth_ADxaWidthThatIsNotPositive_IsGivenTheMinimum(string width)
    {
        var docx = DocOnly($"<w:tbl><w:tr><w:tc><w:tcPr>{width}</w:tcPr><w:p/></w:tc></w:tr></w:tbl>");

        var cell = XDocument.Parse(Part(EmptyCellMinWidth.Apply(docx), "word/document.xml")).Descendants(W + "tcW").Single();

        Assert.Equal("144", (string?)cell.Attribute(W + "w"));
        Assert.Equal("dxa", (string?)cell.Attribute(W + "type"));
    }

    [Fact]
    public void EmptyCellMinWidth_ADxaWidthThatIsPositive_IsLeftAlone()
    {
        var docx = DocOnly("<w:tbl><w:tr><w:tc><w:tcPr><w:tcW w:w=\"900\" w:type=\"dxa\"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>");

        Assert.Same(docx, EmptyCellMinWidth.Apply(docx));
    }

    // ---- HyperlinkBlankRunSplit --------------------------------------------------------------------

    [Fact]
    public void HyperlinkBlankRunSplit_TwoTextRunsBeforeABlankOne_StayInOneLink()
    {
        var docx = DocOnly(
            "<w:p><w:hyperlink r:id=\"rId1\"><w:r><w:t>one</w:t></w:r><w:r><w:t>two</w:t></w:r><w:r><w:t> </w:t></w:r></w:hyperlink></w:p>");

        var paragraph = XDocument.Parse(Part(HyperlinkBlankRunSplit.Apply(docx), "word/document.xml")).Descendants(W + "p").Single();
        var links = paragraph.Elements(W + "hyperlink").ToList();

        Assert.Single(links);
        Assert.Equal(2, links[0].Elements(W + "r").Count());
        Assert.Equal("onetwo", string.Concat(links[0].Descendants(W + "t").Select(t => t.Value)));
        // The blank run follows the link as a plain run, outside it.
        Assert.Equal(2, paragraph.Elements().Count());
        Assert.Equal(W + "r", paragraph.Elements().Last().Name);
    }

    [Fact]
    public void HyperlinkBlankRunSplit_ARunHoldingAPictureOrTextIsNotBlank_ARunOfTabsAndSpacesIs()
    {
        var docx = DocOnly(
            "<w:p><w:hyperlink r:id=\"rId1\">"
            + "<w:r><w:rPr/><w:tab/><w:t> </w:t></w:r>"
            + "<w:r><w:drawing><a/></w:drawing></w:r>"
            + "<w:r><w:t>words</w:t></w:r>"
            + "</w:hyperlink></w:p>");

        var paragraph = XDocument.Parse(Part(HyperlinkBlankRunSplit.Apply(docx), "word/document.xml")).Descendants(W + "p").Single();

        // Only the first run is blank, so it leaves the link and the other two stay in one.
        Assert.Equal([W + "r", W + "hyperlink"], paragraph.Elements().Select(e => e.Name));
        Assert.Equal(2, paragraph.Element(W + "hyperlink")!.Elements(W + "r").Count());
    }

    // ---- IndentFitToColumn -------------------------------------------------------------------------

    [Theory]
    // No margins and no columns: the whole page width is the column.
    [InlineData("<w:pgSz w:w=\"11906\"/>", "w:left=\"20000\"", "left", "11186")]
    // Explicit columns: the narrowest wins, a width that is not a number counts for nothing, and the
    // right indent is taken off as well (3000 - 720 - 100).
    [InlineData("<w:pgSz w:w=\"11906\"/><w:pgMar w:left=\"1440\" w:right=\"1440\"/><w:cols w:num=\"2\"><w:col w:w=\"3000\"/><w:col w:w=\"5000\"/><w:col w:w=\"x\"/></w:cols>",
        "w:start=\"4000\" w:end=\"100\"", "start", "2180")]
    // Columns with a count and a gap but no widths: (9026 - 2 * 360) / 3 = 2768, less 720.
    [InlineData("<w:pgSz w:w=\"11906\"/><w:pgMar w:left=\"1440\" w:right=\"1440\"/><w:cols w:num=\"3\" w:space=\"360\"/>",
        "w:left=\"3000\"", "left", "2048")]
    // A column count with no gap takes Word's default of 720: (9026 - 720) / 2 = 4153, less 720.
    [InlineData("<w:pgSz w:w=\"11906\"/><w:pgMar w:left=\"1440\" w:right=\"1440\"/><w:cols w:num=\"2\"/>",
        "w:left=\"5000\"", "left", "3433")]
    public void IndentFitToColumn_AnIndentWiderThanTheColumn_IsFittedToIt(string section, string indent, string attribute, string expected)
    {
        var docx = DocOnly($"<w:p><w:pPr><w:ind {indent}/></w:pPr></w:p><w:sectPr>{section}</w:sectPr>");

        var ind = XDocument.Parse(Part(IndentFitToColumn.Apply(docx), "word/document.xml")).Descendants(W + "ind").Single();

        Assert.Equal(expected, (string?)ind.Attribute(W + attribute));
    }

    [Fact]
    public void IndentFitToColumn_TheHangingIndentIsFittedWithTheLeftOne_AndOthersAreLeftAlone()
    {
        var docx = DocOnly(
            "<w:p><w:pPr><w:ind w:left=\"20000\" w:hanging=\"15000\"/></w:pPr></w:p>"
            + "<w:p><w:pPr><w:ind w:left=\"100\"/></w:pPr></w:p>"
            + "<w:p><w:pPr><w:ind w:hanging=\"15000\"/></w:pPr></w:p>"
            + "<w:sectPr><w:pgSz w:w=\"11906\"/></w:sectPr>");

        var inds = XDocument.Parse(Part(IndentFitToColumn.Apply(docx), "word/document.xml")).Descendants(W + "ind").ToList();

        Assert.Equal("11186", (string?)inds[0].Attribute(W + "left"));
        Assert.Equal("11186", (string?)inds[0].Attribute(W + "hanging"));
        Assert.Equal("100", (string?)inds[1].Attribute(W + "left"));
        Assert.Equal("15000", (string?)inds[2].Attribute(W + "hanging"));
    }

    [Fact]
    public void IndentFitToColumn_AColumnNoWiderThanTheTextRoom_ReturnsTheSameArray()
    {
        // 2000 - 700 - 700 = 600, which is under the 720 the text itself needs.
        var docx = DocOnly(
            "<w:p><w:pPr><w:ind w:left=\"20000\"/></w:pPr></w:p>"
            + "<w:sectPr><w:pgSz w:w=\"2000\"/><w:pgMar w:left=\"700\" w:right=\"700\"/></w:sectPr>");

        Assert.Same(docx, IndentFitToColumn.Apply(docx));
    }

    // ---- HeaderOffsetRepair ------------------------------------------------------------------------

    private static string SectionWithHeader(string id, string margins = "<w:pgMar w:top=\"288\" w:left=\"1440\" w:right=\"1440\"/>") =>
        $"<w:sectPr><w:headerReference w:type=\"default\" r:id=\"{id}\"/><w:headerReference w:type=\"first\"/><w:headerReference w:type=\"even\" r:id=\"rIdWithNoContent\"/>{margins}</w:sectPr>";

    private static string Rels(params string[] relationships) =>
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" + string.Concat(relationships) + "</Relationships>";

    private static string HeaderWithText(string text) => $"<w:hdr xmlns:w=\"{WNs}\"><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:hdr>";

    [Fact]
    public void HeaderOffsetRepair_AHeaderPartWithNoRelationshipsPart_IsLeftAlone()
    {
        var docx = Pkg(("word/document.xml", Doc("<w:p/>" + SectionWithHeader("rId1"))), ("word/header1.xml", HeaderWithText("Words")));

        Assert.Same(docx, HeaderOffsetRepair.Apply(docx));
    }

    [Fact]
    public void HeaderOffsetRepair_AHeaderThatIsNotWellFormed_IsLeftAlone()
    {
        var docx = Pkg(
            ("word/document.xml", Doc("<w:p/>" + SectionWithHeader("rId1"))),
            ("word/_rels/document.xml.rels", Rels($"<Relationship Id=\"rId1\" Type=\"{HeaderType}\" Target=\"header1.xml\"/>")),
            ("word/header1.xml", "<w:hdr"));

        Assert.Same(docx, HeaderOffsetRepair.Apply(docx));
    }

    [Fact]
    public void HeaderOffsetRepair_OnlyTheHeaderRelationshipThatCanBeResolvedCounts()
    {
        // Four relationships that cannot name a header with content, then the one that can. The
        // absolute target ("/word/header1.xml") is the form Word itself writes.
        var docx = Pkg(
            ("word/document.xml", Doc("<w:p/>" + SectionWithHeader("rId9"))),
            ("word/_rels/document.xml.rels", Rels(
                "<Relationship Id=\"rId1\"/>",
                $"<Relationship Id=\"rId2\" Type=\"{HeaderType}\"/>",
                $"<Relationship Id=\"rId3\" Type=\"{HeaderType}\" Target=\"missing.xml\"/>",
                $"<Relationship Type=\"{HeaderType}\" Target=\"header1.xml\"/>",
                $"<Relationship Id=\"rId9\" Type=\"{HeaderType}\" Target=\"/word/header1.xml\"/>")),
            ("word/header1.xml", HeaderWithText("Words")));

        var repaired = HeaderOffsetRepair.Apply(docx);

        Assert.NotSame(docx, repaired);
        Assert.Equal("360", (string?)XDocument.Parse(Part(repaired, "word/document.xml")).Descendants(W + "pgMar").Single().Attribute(W + "top"));
    }

    [Fact]
    public void HeaderOffsetRepair_AHeaderNoSectionReferences_IsLeftAlone()
    {
        var docx = Pkg(
            ("word/document.xml", Doc("<w:p/><w:sectPr><w:pgMar w:top=\"288\"/></w:sectPr>")),
            ("word/_rels/document.xml.rels", Rels($"<Relationship Id=\"rId1\" Type=\"{HeaderType}\" Target=\"header1.xml\"/>")),
            ("word/header1.xml", HeaderWithText("Words")));

        Assert.Same(docx, HeaderOffsetRepair.Apply(docx));
    }

    [Fact]
    public void HeaderOffsetRepair_ATopMarginThatIsNotANumber_OrAMissingOne_IsLeftAlone()
    {
        var docx = Pkg(
            ("word/document.xml", Doc(
                "<w:p><w:pPr>" + SectionWithHeader("rId1", "<w:pgMar w:top=\"tall\"/>") + "</w:pPr></w:p>"
                + "<w:sectPr><w:pgSz w:w=\"11906\"/></w:sectPr>")),
            ("word/_rels/document.xml.rels", Rels($"<Relationship Id=\"rId1\" Type=\"{HeaderType}\" Target=\"header1.xml\"/>")),
            ("word/header1.xml", HeaderWithText("Words")));

        Assert.Same(docx, HeaderOffsetRepair.Apply(docx));
    }

    // ---- ListMarkerSubstitution --------------------------------------------------------------------

    private const string Numbering = "<w:numbering xmlns:w=\"" + WNs + "\"><w:abstractNum w:abstractNumId=\"0\">{0}</w:abstractNum></w:numbering>";

    private static string LevelXml(byte[] docx) => Part(docx, "word/numbering.xml");

    [Fact]
    public void ListMarkerSubstitution_ANumberingPartThatIsNotWellFormed_StillGetsItsCharactersSubstituted()
    {
        // The font fix needs a parse and is skipped; the character substitution is plain text and is not.
        var docx = Pkg(
            ("word/document.xml", Doc("<w:p/>")),
            ("word/numbering.xml", $"<w:numbering xmlns:w=\"{WNs}\"><w:lvl><w:lvlText w:val=\"\uF0B7\"/><w:lvlText w:val=\"\uFB01\"/>"));

        var xml = LevelXml(ListMarkerSubstitution.Apply(docx));

        Assert.Contains("w:val=\"\u2022\"", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("\uF0B7", xml, StringComparison.Ordinal);
        // A character just above the private-use block is not one.
        Assert.Contains("w:val=\"\uFB01\"", xml, StringComparison.Ordinal);
    }

    [Fact]
    public void ListMarkerSubstitution_OnlyALevelWithASymbolFaceAndAPrivateUseMarkerGetsArial()
    {
        var levels = string.Concat(
            // 0: no marker text at all
            "<w:lvl w:ilvl=\"0\"><w:rPr><w:rFonts w:ascii=\"Symbol\"/></w:rPr></w:lvl>",
            // 1: an ordinary marker
            "<w:lvl w:ilvl=\"1\"><w:lvlText w:val=\"o\"/><w:rPr><w:rFonts w:ascii=\"Symbol\"/></w:rPr></w:lvl>",
            // 2: a private-use marker and no font
            "<w:lvl w:ilvl=\"2\"><w:lvlText w:val=\"\uF0B7\"/></w:lvl>",
            // 3: a private-use marker and a font that is not a symbol one
            "<w:lvl w:ilvl=\"3\"><w:lvlText w:val=\"\uF0B7\"/><w:rPr><w:rFonts w:ascii=\"Calibri\"/></w:rPr></w:lvl>",
            // 4: the case that is repaired, with Wingdings
            "<w:lvl w:ilvl=\"4\"><w:lvlText w:val=\"\uF0A7\"/><w:rPr><w:rFonts w:ascii=\"Wingdings\" w:hAnsi=\"Wingdings\"/></w:rPr></w:lvl>",
            // 5: a private-use character the table does not know
            "<w:lvl w:ilvl=\"5\"><w:lvlText w:val=\"\uF0FF\"/></w:lvl>");
        // No XML declaration, which is also what Word writes for a part it produced by hand.
        var docx = Pkg(("word/document.xml", Doc("<w:p/>")), ("word/numbering.xml", string.Format(Numbering, levels)));

        var repaired = XDocument.Parse(LevelXml(ListMarkerSubstitution.Apply(docx)));
        var byLevel = repaired.Descendants(W + "lvl").ToDictionary(l => (string)l.Attribute(W + "ilvl")!);

        Assert.Equal("Symbol", (string?)byLevel["0"].Descendants(W + "rFonts").Single().Attribute(W + "ascii"));
        Assert.Equal("Symbol", (string?)byLevel["1"].Descendants(W + "rFonts").Single().Attribute(W + "ascii"));
        Assert.Equal("Calibri", (string?)byLevel["3"].Descendants(W + "rFonts").Single().Attribute(W + "ascii"));

        var arial = byLevel["4"].Descendants(W + "rFonts").Single();
        foreach (var slot in new[] { "ascii", "hAnsi", "cs", "eastAsia" })
            Assert.Equal("Arial", (string?)arial.Attribute(W + slot));

        // Every private-use marker is replaced, including the one nothing maps: it gets the bullet.
        Assert.Equal("\u2022", (string?)byLevel["5"].Element(W + "lvlText")!.Attribute(W + "val"));
        Assert.DoesNotContain(repaired.Descendants().SelectMany(e => e.Attributes()), a => a.Value.Any(c => c is >= '\uE000' and <= '\uF8FF'));
    }

    // ---- ListMarkerSuffixSpace ---------------------------------------------------------------------

    [Fact]
    public void ListMarkerSuffixSpace_ALevelWithNeitherASuffixNorMarkerText_GetsTheSuffixAppended()
    {
        var docx = Pkg(
            ("word/document.xml", Doc("<w:p/>")),
            ("word/numbering.xml", string.Format(Numbering,
                "<w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/></w:lvl>"
                + "<w:lvl w:ilvl=\"1\"><w:suff w:val=\"tab\"/></w:lvl>"
                + "<w:lvl w:ilvl=\"2\"><w:suff w:val=\"nothing\"/></w:lvl>")));

        var levels = XDocument.Parse(Part(ListMarkerSuffixSpace.Apply(docx), "word/numbering.xml")).Descendants(W + "lvl").ToList();

        Assert.Equal(["start", "suff"], levels[0].Elements().Select(e => e.Name.LocalName));
        Assert.Equal("space", (string?)levels[0].Element(W + "suff")!.Attribute(W + "val"));
        Assert.Equal("space", (string?)levels[1].Element(W + "suff")!.Attribute(W + "val"));
        Assert.Equal("nothing", (string?)levels[2].Element(W + "suff")!.Attribute(W + "val"));
    }

    // ---- WideTableScale ----------------------------------------------------------------------------

    private static string Section => $"<w:sectPr>{Page}</w:sectPr>";

    private static List<string?> GridWidths(byte[] docx) =>
        XDocument.Parse(Part(docx, "word/document.xml")).Descendants(W + "gridCol").Select(g => (string?)g.Attribute(W + "w")).ToList();

    [Fact]
    public void WideTableScale_ADocumentWithNoBody_ReturnsTheSameArray()
    {
        var docx = Pkg(("word/document.xml", $"<w:document xmlns:w=\"{WNs}\"/>"));

        Assert.Same(docx, WideTableScale.Apply(docx));
    }

    [Fact]
    public void WideTableScale_ATableWithNoPropertiesAndNoGrid_HasNoWidthToScale()
    {
        var docx = DocOnly("<w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>" + Section);

        Assert.Same(docx, WideTableScale.Apply(docx));
    }

    [Fact]
    public void WideTableScale_AnAutoLayoutTableWithNoLayoutElement_IsLeftToTheRenderer()
    {
        var docx = DocOnly(
            "<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr><w:tblGrid><w:gridCol w:w=\"12000\"/></w:tblGrid></w:tbl>" + Section);

        Assert.Same(docx, WideTableScale.Apply(docx));
    }

    [Fact]
    public void WideTableScale_AFixedTableWhoseGridFits_IsLeftAlone()
    {
        var docx = DocOnly(
            "<w:tbl><w:tblPr><w:tblLayout w:type=\"fixed\"/></w:tblPr><w:tblGrid><w:gridCol w:w=\"4000\"/><w:gridCol w:w=\"4000\"/></w:tblGrid></w:tbl>" + Section);

        Assert.Same(docx, WideTableScale.Apply(docx));
    }

    [Fact]
    public void WideTableScale_AFixedTableWithAWideGridAndNoRows_IsScaledByItsGrid()
    {
        var docx = DocOnly(
            "<w:tbl><w:tblPr><w:tblLayout w:type=\"fixed\"/></w:tblPr><w:tblGrid><w:gridCol w:w=\"6000\"/><w:gridCol w:w=\"12000\"/></w:tblGrid></w:tbl>"
            + Section);

        // 9026 / 18000 = 0.50144..., truncated per column: 3008 + 6017 = 9025, never more than the text area.
        Assert.Equal(["3008", "6017"], GridWidths(WideTableScale.Apply(docx)));
    }

    // ---- DocxPdfFailureDiagnosis -------------------------------------------------------------------

    private sealed class NoMessageException : Exception
    {
        public override string Message => null!;
    }

    [Fact]
    public void DocxPdfFailureDiagnosis_AnExceptionWithNoMessage_IsNotRecognised_AndDoesNotThrow()
    {
        Assert.Null(DocxPdfFailureDiagnosis.Describe(new NoMessageException()));
    }
}
