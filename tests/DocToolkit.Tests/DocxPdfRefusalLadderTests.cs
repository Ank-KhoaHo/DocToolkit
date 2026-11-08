using System.IO.Compression;
using System.Text;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

namespace DocToolkit.Tests;

/// <summary>
/// The refusal ladder (A123): repairs that run only after OfficeIMO 3.5 refuses a DOCX with their
/// message, so a document that renders as it stands is never changed.
///
/// <b>Every fixture is first rendered by the unrepaired renderer and must be refused with the
/// step's own message.</b> The plan's first fixtures were mostly NOT refused by 3.5.3, and a repair
/// tested on a document the renderer accepts proves nothing. The shapes here were bisected from the
/// measured documents (see each step's remarks), and where the plan's shape rendered, the real
/// trigger is noted on the fixture.
///
/// <b>The packages are written by hand as WordprocessingML</b>, not through the OpenXml SDK: the
/// triggers are particular twip values, tab stops and numbering levels, and the SDK's fixtures do
/// not carry a styles or numbering part.
/// </summary>
public class DocxPdfRefusalLadderTests
{
    private const string PaddingRefusal = "cell padding must leave a positive text width";
    private const string IndentRefusal = "indents must leave a positive text width";
    private const string MarkerRefusal = "Inline element width exceeds the available paragraph width";
    private const string WidthRefusal = "preferred width must be a positive finite value";

    // ---- helpers -----------------------------------------------------------------------------------

    private const string Ns = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"";

    private const string A4 =
        "<w:sectPr><w:pgSz w:w=\"11906\" w:h=\"16838\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" "
        + "w:left=\"1440\" w:header=\"720\" w:footer=\"720\" w:gutter=\"0\"/></w:sectPr>";

    /// <summary>Two columns of 4,752 twips on a page whose text area is 10,224 wide (4752 + 720 + 4752).</summary>
    private const string TwoColumns =
        "<w:sectPr><w:pgSz w:w=\"13104\" w:h=\"16838\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" "
        + "w:left=\"1440\" w:header=\"720\" w:footer=\"720\" w:gutter=\"0\"/>"
        + "<w:cols w:num=\"2\" w:space=\"720\" w:equalWidth=\"0\"><w:col w:w=\"4752\" w:space=\"720\"/><w:col w:w=\"4752\"/></w:cols></w:sectPr>";

    private static byte[] Package(string body, string? styles = null, string? numbering = null, string section = A4)
    {
        const string Declaration = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
        const string Rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        const string Word = "application/vnd.openxmlformats-officedocument.wordprocessingml";

        var parts = new List<(string Name, string Content)>
        {
            ("[Content_Types].xml", Declaration
                + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                + $"<Override PartName=\"/word/document.xml\" ContentType=\"{Word}.document.main+xml\"/>"
                + (styles is null ? "" : $"<Override PartName=\"/word/styles.xml\" ContentType=\"{Word}.styles+xml\"/>")
                + (numbering is null ? "" : $"<Override PartName=\"/word/numbering.xml\" ContentType=\"{Word}.numbering+xml\"/>")
                + "</Types>"),
            ("_rels/.rels", Declaration
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + $"<Relationship Id=\"rId1\" Type=\"{Rel}/officeDocument\" Target=\"word/document.xml\"/></Relationships>"),
            ("word/_rels/document.xml.rels", Declaration
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + (styles is null ? "" : $"<Relationship Id=\"rId1\" Type=\"{Rel}/styles\" Target=\"styles.xml\"/>")
                + (numbering is null ? "" : $"<Relationship Id=\"rId2\" Type=\"{Rel}/numbering\" Target=\"numbering.xml\"/>")
                + "</Relationships>"),
            ("word/document.xml", Declaration + $"<w:document {Ns}><w:body>{body}{section}</w:body></w:document>"),
        };
        if (styles is not null) parts.Add(("word/styles.xml", Declaration + $"<w:styles {Ns}>{styles}</w:styles>"));
        if (numbering is not null) parts.Add(("word/numbering.xml", Declaration + $"<w:numbering {Ns}>{numbering}</w:numbering>"));

        using var ms = new MemoryStream();
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var (name, content) in parts)
            {
                using var writer = new StreamWriter(zip.CreateEntry(name).Open(), new UTF8Encoding(false));
                writer.Write(content);
            }
        }

        return ms.ToArray();
    }

    private static string Run(string text) => $"<w:r><w:t>{text}</w:t></w:r>";

    private static string Para(string text, string properties = "") =>
        "<w:p>" + (properties.Length == 0 ? "" : $"<w:pPr>{properties}</w:pPr>") + Run(text) + "</w:p>";

    private const string EmptyPara = "<w:p/>";

    private static string Cell(string content, string properties = "") =>
        "<w:tc>" + (properties.Length == 0 ? "" : $"<w:tcPr>{properties}</w:tcPr>") + content + "</w:tc>";

    private static string SideMargins(int twips, string extra = "") =>
        $"<w:tblCellMar>{extra}<w:left w:w=\"{twips}\" w:type=\"dxa\"/><w:right w:w=\"{twips}\" w:type=\"dxa\"/></w:tblCellMar>";

    private static string AutoTable(int columns, string margins, Func<int, int, string> cell, int rows = 2) =>
        "<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/>" + margins + "</w:tblPr><w:tblGrid>"
        + string.Concat(Enumerable.Repeat("<w:gridCol w:w=\"0\"/>", columns)) + "</w:tblGrid>"
        + string.Concat(Enumerable.Range(0, rows).Select(r =>
            "<w:tr>" + string.Concat(Enumerable.Range(0, columns).Select(c => cell(r, c))) + "</w:tr>"))
        + "</w:tbl>";

    /// <summary>
    /// Nine automatic columns, the ninth empty in every row, with 30-twip side margins. The shape of
    /// html-000062, where an HTML-to-DOCX writer left a spacer cell. Bisected: the padding must be
    /// small (a 108-twip margin is wider than the 144-twip minimum and is not enough), and nine
    /// columns is where the empty one is refused (three columns with the same margin render).
    /// </summary>
    private static string NineColumnsWithAnEmptyOne(string margins = "") =>
        AutoTable(9, margins.Length == 0 ? SideMargins(30) : margins,
            (r, c) => Cell(c == 8 ? EmptyPara : Para($"c{c}{r}"))) + Para("after the table");

    /// <summary>
    /// Fifty automatic columns of one letter each, with 108-twip side margins: each column is
    /// narrower than its padding. Every cell holds text, so there is no empty cell for the width step.
    /// </summary>
    private static string FiftyNarrowColumns(string margins) =>
        AutoTable(50, margins, (_, _) => Cell(Para("x")), rows: 1) + Para("after the table");

    private static string IndentedParagraph(string indent) =>
        Para("indented paragraph text", $"<w:ind {indent}/>") + Para("plain paragraph");

    private static string ZeroWidthPercentTables() =>
        "<w:tbl><w:tblPr><w:tblW w:w=\"5000\" w:type=\"pct\"/></w:tblPr><w:tblGrid><w:gridCol/></w:tblGrid><w:tr>"
        + Cell(Para("before") + "<w:tbl><w:tblPr><w:tblW w:w=\"50\" w:type=\"pct\"/>"
            + "<w:tblCellSpacing w:w=\"90\" w:type=\"dxa\"/></w:tblPr><w:tblGrid><w:gridCol/></w:tblGrid><w:tr>"
            + Cell(Para("tiny percent table")) + "</w:tr></w:tbl>" + Para("after the nested one"))
        + "</w:tr></w:tbl>" + Para("end");

    private static void AssertRefused(byte[] docx, string message)
    {
        var refusal = RawRender.Refusal(docx);
        Assert.IsType<ArgumentException>(refusal);
        Assert.Contains(message, refusal.Message, StringComparison.Ordinal);
    }

    /// <summary>
    /// Asserts the refusal where this host reproduces it. R11 and R13 sit on a font-metrics edge (the
    /// measured Windows host, with Arial and Calibri, and both Linux test images refuse these
    /// fixtures; a host with other fonts may not), so the refusal is checked when it happens and the
    /// repair is asserted unconditionally either way.
    /// </summary>
    private static void AssertRefusedWhereItReproduces(byte[] docx, string message)
    {
        if (RawRender.Refusal(docx) is not { } refusal) return;

        Assert.IsType<ArgumentException>(refusal);
        Assert.Contains(message, refusal.Message, StringComparison.Ordinal);
    }

    private static string Part(byte[] docx, string name)
    {
        using var ms = new MemoryStream(docx);
        using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
        using var reader = new StreamReader(zip.GetEntry(name)!.Open(), Encoding.UTF8);

        // The writer puts a space before a self-closing slash; the assertions are written without it.
        return reader.ReadToEnd().Replace(" />", "/>", StringComparison.Ordinal);
    }

    private static string Between(string xml, string open, string close)
    {
        var start = xml.IndexOf(open, StringComparison.Ordinal);
        Assert.True(start >= 0, $"{open} not found");
        return xml[start..(xml.IndexOf(close, start, StringComparison.Ordinal) + close.Length)];
    }

    // ---- R11: an empty column ----------------------------------------------------------------------

    [Fact]
    public void AnEmptyColumnRenders_AndKeepsItsPadding()
    {
        var docx = Package(NineColumnsWithAnEmptyOne());
        // Raw 3.5.3: "Table horizontal cell padding must leave a positive text width."
        AssertRefusedWhereItReproduces(docx, PaddingRefusal);

        var repaired = EmptyCellMinWidth.Apply(docx);

        Assert.NotSame(docx, repaired);
        Assert.Null(RawRender.Refusal(repaired));
        // The padding is exactly what it was: this step gives a width and touches no margin.
        Assert.Equal(
            Between(Part(docx, "word/document.xml"), "<w:tblCellMar>", "</w:tblCellMar>"),
            Between(Part(repaired, "word/document.xml"), "<w:tblCellMar>", "</w:tblCellMar>"));

        var text = PdfProbe.ExtractText(DocxToPdfConverter.Convert(docx));
        Assert.Contains("c70", text);
        Assert.Contains("after the table", text);
    }

    [Theory]
    [InlineData("")]                                                 // no tcW at all
    [InlineData("<w:tcW w:w=\"0\" w:type=\"auto\"/>")]               // automatic
    [InlineData("<w:tcW w:w=\"0\" w:type=\"dxa\"/>")]                // zero twips
    [InlineData("<w:tcW w:w=\"700\" w:type=\"auto\"/>")]             // a width, but not in twips
    public void AnEmptyUnsizedCellGets144Twips(string width)
    {
        var docx = Package(AutoTable(2, "", (_, c) => c == 1 ? Cell(EmptyPara, width) : Cell(Para("text"))));

        var cell = Between(Part(EmptyCellMinWidth.Apply(docx), "word/document.xml"), "<w:tcW", "/>");

        Assert.Contains("w:w=\"144\"", cell);
        Assert.Contains("w:type=\"dxa\"", cell);
    }

    [Fact]
    public void ACellWithNoTcPr_GetsOneThatStartsTheCell()
    {
        var xml = Part(
            EmptyCellMinWidth.Apply(Package(AutoTable(2, "", (_, c) => c == 1 ? Cell(EmptyPara) : Cell(Para("text"))))),
            "word/document.xml");

        Assert.Contains("<w:tc><w:tcPr><w:tcW w:w=\"144\" w:type=\"dxa\"/></w:tcPr><w:p/></w:tc>", xml);
    }

    [Theory]
    [InlineData("<w:tcW w:w=\"700\" w:type=\"dxa\"/>")]              // already a width in twips
    [InlineData("<w:tcW w:w=\"1000\" w:type=\"pct\"/>")]             // a percentage is never replaced
    public void AnEmptyCellThatIsSized_IsUntouched(string width)
    {
        var docx = Package(AutoTable(2, "", (_, c) => c == 1 ? Cell(EmptyPara, width) : Cell(Para("text"))));

        Assert.Same(docx, EmptyCellMinWidth.Apply(docx));
    }

    [Fact]
    public void ACellWithText_IsUntouched()
    {
        var docx = Package(AutoTable(2, "", (_, _) => Cell(Para("text"))));

        Assert.Same(docx, EmptyCellMinWidth.Apply(docx));
    }

    // ---- R4: padding wider than a column -----------------------------------------------------------

    [Fact]
    public void PaddingWiderThanAColumnRenders()
    {
        // Bisected: 50 columns of 9,026 / 50 = 180 twips are narrower than 2 x 108 of padding. Raw
        // 3.5.3 refuses it; the same table at zero padding renders.
        var docx = Package(FiftyNarrowColumns(SideMargins(108, "<w:top w:w=\"55\" w:type=\"dxa\"/>")));
        AssertRefused(docx, PaddingRefusal);
        Assert.Same(docx, EmptyCellMinWidth.Apply(docx)); // no empty cell, so the width step has nothing to do

        var repaired = CellPaddingZero.Apply(docx);

        Assert.Null(RawRender.Refusal(repaired));
        var margins = Between(Part(repaired, "word/document.xml"), "<w:tblCellMar>", "</w:tblCellMar>");
        Assert.Contains("<w:left w:w=\"0\" w:type=\"dxa\"", margins);
        Assert.Contains("<w:right w:w=\"0\" w:type=\"dxa\"", margins);
        Assert.Contains("<w:top w:w=\"55\" w:type=\"dxa\"", margins); // only the sides

        Assert.Contains("after the table", PdfProbe.ExtractText(DocxToPdfConverter.Convert(docx)));
    }

    [Fact]
    public void PaddingIsZeroedInCellsAndStylesToo_InEverySpelling()
    {
        var styles = "<w:style w:type=\"table\" w:styleId=\"Wide\"><w:name w:val=\"Wide\"/><w:tblPr>"
            + SideMargins(108).Replace("w:left", "w:start").Replace("w:right", "w:end") + "</w:tblPr></w:style>";
        var cell = "<w:tc><w:tcPr><w:tcMar><w:left w:w=\"90\" w:type=\"dxa\"/><w:right w:w=\"90\" w:type=\"dxa\"/>"
            + "</w:tcMar></w:tcPr><w:p/></w:tc>";
        var docx = Package("<w:tbl><w:tblPr/><w:tblGrid><w:gridCol/></w:tblGrid><w:tr>" + cell + "</w:tr></w:tbl>", styles);

        var repaired = CellPaddingZero.Apply(docx);

        Assert.DoesNotContain("w:w=\"90\"", Part(repaired, "word/document.xml"));
        var styleXml = Part(repaired, "word/styles.xml");
        Assert.DoesNotContain("w:w=\"108\"", styleXml);
        Assert.Contains("<w:start w:w=\"0\" w:type=\"dxa\"", styleXml);
        Assert.Contains("<w:end w:w=\"0\" w:type=\"dxa\"", styleXml);
    }

    [Fact]
    public void ATableWithNoExplicitPadding_IsUntouchedByTheZeroing()
    {
        var docx = Package(AutoTable(3, "", (_, _) => Cell(Para("x"))));

        Assert.Same(docx, CellPaddingZero.Apply(docx));
    }

    [Fact]
    public void PaddingAlreadyZero_IsUntouched()
    {
        var docx = Package(AutoTable(3, SideMargins(0), (_, _) => Cell(Para("x"))));

        Assert.Same(docx, CellPaddingZero.Apply(docx));
    }

    // ---- R12: an indent wider than its column ------------------------------------------------------

    [Fact]
    public void AnIndentWiderThanItsColumnRenders()
    {
        // doc-000111: left 6480 / hanging 6480 in a section of two 4,752-twip columns. The page is
        // widened so the section itself is accepted and the INDENT is what 3.5 refuses.
        var docx = Package(IndentedParagraph("w:left=\"6480\" w:right=\"-576\" w:hanging=\"6480\""), section: TwoColumns);
        AssertRefused(docx, IndentRefusal);

        var repaired = IndentFitToColumn.Apply(docx);

        Assert.Null(RawRender.Refusal(repaired));
        // 4752 - 720 = 4032 (the negative right indent counts as zero); the hanging indent follows.
        var indent = Between(Part(repaired, "word/document.xml"), "<w:ind", "/>");
        Assert.Contains("w:left=\"4032\"", indent);
        Assert.Contains("w:hanging=\"4032\"", indent);
        Assert.Contains("w:right=\"-576\"", indent);

        Assert.Contains("indented paragraph text", PdfProbe.ExtractText(DocxToPdfConverter.Convert(docx)));
    }

    [Fact]
    public void AStartIndentIsFitToo()
    {
        // Raw 3.5.3 does NOT refuse the w:start / w:end spelling (measured), so there is no refusal
        // to show here; what is pinned is that the step reads it, for Strict and newer Word, and
        // leaves the end indent alone.
        var docx = Package(IndentedParagraph("w:start=\"6480\" w:end=\"-576\" w:hanging=\"6480\""), section: TwoColumns);

        var indent = Between(Part(IndentFitToColumn.Apply(docx), "word/document.xml"), "<w:ind", "/>");

        Assert.Contains("w:start=\"4032\"", indent);
        Assert.Contains("w:hanging=\"4032\"", indent);
        Assert.Contains("w:end=\"-576\"", indent);
    }

    [Fact]
    public void ARightIndentLeavesLessRoomForTheLeftOne()
    {
        var docx = Package(IndentedParagraph("w:left=\"6480\" w:right=\"1000\""), section: TwoColumns);

        Assert.Contains("w:left=\"3032\"", Part(IndentFitToColumn.Apply(docx), "word/document.xml"));
    }

    [Fact]
    public void IndentsInStylesAndNumberingAreFitToo()
    {
        var styles = "<w:style w:type=\"paragraph\" w:styleId=\"Wide\"><w:name w:val=\"Wide\"/><w:pPr><w:ind w:left=\"9999\"/></w:pPr></w:style>";
        var numbering = "<w:abstractNum w:abstractNumId=\"0\"><w:lvl w:ilvl=\"0\"><w:pPr><w:ind w:left=\"8000\" w:hanging=\"8000\"/></w:pPr></w:lvl></w:abstractNum>";
        var docx = Package(Para("text"), styles, numbering, TwoColumns);

        var repaired = IndentFitToColumn.Apply(docx);

        Assert.Contains("w:left=\"4032\"", Part(repaired, "word/styles.xml"));
        Assert.Contains("w:left=\"4032\" w:hanging=\"4032\"", Part(repaired, "word/numbering.xml"));
    }

    [Fact]
    public void AnIndentThatFits_IsUntouched()
    {
        var docx = Package(IndentedParagraph("w:left=\"4032\" w:hanging=\"360\""), section: TwoColumns);

        Assert.Same(docx, IndentFitToColumn.Apply(docx));
    }

    [Fact]
    public void NoSectionToMeasure_IsUntouched()
    {
        var docx = Package(IndentedParagraph("w:left=\"6480\""), section: "");

        Assert.Same(docx, IndentFitToColumn.Apply(docx));
    }

    // ---- R13: a list marker whose tab runs past the cell -------------------------------------------

    /// <summary>
    /// doc-000006 reduced to what matters (bisected): a numbered paragraph in the first cell of a
    /// fixed-layout table, in a style that numbers it and CLEARS the level's own tab stop at 1,080
    /// while setting a centre stop at 4,680 - past the 4,698-twip cell's 4,482-twip text area.
    /// <b>Document defaults with a font size are needed, and it is a measured oddity:</b> without
    /// them the same package renders, which is why the plan's version of this fixture was not
    /// refused. At 11 pt with no font named it is refused in both test images; naming a font
    /// (Arial, Calibri) makes the fonts image render it, and 14 pt renders everywhere, so the
    /// refusal is a font-metrics edge and this fixture sits deliberately on the refused side of it.
    /// </summary>
    private static byte[] ListMarkerPastTheCell(string suffix)
    {
        const string defaults = "<w:docDefaults><w:rPrDefault><w:rPr>"
            + "<w:sz w:val=\"22\"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:before=\"0\" w:after=\"0\" "
            + "w:line=\"240\" w:lineRule=\"auto\"/></w:pPr></w:pPrDefault></w:docDefaults>";
        const string styles = defaults
            + "<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/></w:style>"
            + "<w:style w:type=\"paragraph\" w:styleId=\"LegacyDocBodyText\"><w:name w:val=\"Body Text\"/><w:basedOn w:val=\"Normal\"/>"
            + "<w:pPr><w:widowControl w:val=\"false\"/><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"22\"/></w:numPr>"
            + "<w:tabs><w:tab w:val=\"clear\" w:pos=\"1080\"/><w:tab w:val=\"center\" w:pos=\"4680\"/></w:tabs>"
            + "<w:suppressAutoHyphens w:val=\"true\"/><w:ind w:firstLine=\"0\"/></w:pPr></w:style>";
        var numbering = "<w:abstractNum w:abstractNumId=\"0\"><w:multiLevelType w:val=\"singleLevel\"/><w:lvl w:ilvl=\"0\">"
            + "<w:start w:val=\"1\"/><w:numFmt w:val=\"decimal\"/><w:pStyle w:val=\"LegacyDocBodyText\"/>" + suffix
            + "<w:lvlText w:val=\"%1.\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:tabs><w:tab w:val=\"num\" w:pos=\"1080\"/></w:tabs>"
            + "<w:ind w:left=\"0\" w:firstLine=\"720\"/></w:pPr></w:lvl></w:abstractNum>"
            + "<w:num w:numId=\"22\"><w:abstractNumId w:val=\"0\"/></w:num>";
        var numbered = "<w:p><w:pPr><w:pStyle w:val=\"LegacyDocBodyText\"/></w:pPr></w:p>"
            + "<w:p><w:pPr><w:pStyle w:val=\"LegacyDocBodyText\"/></w:pPr>" + Run("Petition for determination") + "</w:p>";
        var table = "<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/><w:tblInd w:w=\"-108\" w:type=\"dxa\"/>"
            + "<w:tblLayout w:type=\"fixed\"/></w:tblPr><w:tblGrid><w:gridCol w:w=\"4698\"/><w:gridCol w:w=\"630\"/>"
            + "<w:gridCol w:w=\"4248\"/></w:tblGrid><w:tr>"
            + Cell(numbered, "<w:tcW w:w=\"4698\" w:type=\"dxa\"/>")
            + Cell(Para("mid"), "<w:tcW w:w=\"630\" w:type=\"dxa\"/>")
            + Cell(Para("right"), "<w:tcW w:w=\"4248\" w:type=\"dxa\"/>")
            + "</w:tr></w:tbl>" + Para("after the table");
        const string letter = "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" "
            + "w:bottom=\"720\" w:left=\"1440\" w:header=\"720\" w:footer=\"720\" w:gutter=\"0\"/></w:sectPr>";
        return Package(table, styles, numbering, letter);
    }

    [Theory]
    [InlineData("<w:suff w:val=\"tab\"/>")]
    [InlineData("")]                      // an absent suffix means a tab
    public void AListMarkerTabPastTheCellEdgeRenders(string suffix)
    {
        var docx = ListMarkerPastTheCell(suffix);
        AssertRefusedWhereItReproduces(docx, MarkerRefusal);

        var repaired = ListMarkerSuffixSpace.Apply(docx);

        Assert.Null(RawRender.Refusal(repaired));
        var text = PdfProbe.ExtractText(DocxToPdfConverter.Convert(docx));
        Assert.Contains("Petition for determination", text);
        Assert.Contains("after the table", text);
    }

    [Fact]
    public void AnAbsentSuffixIsAddedWhereTheSchemaPutsIt()
    {
        var xml = Part(ListMarkerSuffixSpace.Apply(ListMarkerPastTheCell("")), "word/numbering.xml");

        Assert.Contains("<w:suff w:val=\"space\"/><w:lvlText", xml);
    }

    [Theory]
    [InlineData("<w:suff w:val=\"space\"/>")]
    [InlineData("<w:suff w:val=\"nothing\"/>")]
    public void AMarkerThatDoesNotUseATab_IsUntouched(string suffix)
    {
        var docx = ListMarkerPastTheCell(suffix);

        Assert.Same(docx, ListMarkerSuffixSpace.Apply(docx));
    }

    [Fact]
    public void NoNumberingPart_IsUntouched()
    {
        var docx = Package(Para("text"));

        Assert.Same(docx, ListMarkerSuffixSpace.Apply(docx));
    }

    // ---- the ladder --------------------------------------------------------------------------------

    [Fact]
    public void ADocumentNeedingTwoStepsRenders()
    {
        // An empty column (refused first, the table comes first) AND a too-wide indent (refused only
        // once the table is out of the way). Neither repair alone is enough.
        var docx = Package(
            NineColumnsWithAnEmptyOne() + IndentedParagraph("w:left=\"6480\" w:right=\"-576\" w:hanging=\"6480\""),
            section: TwoColumns);
        AssertRefusedWhereItReproduces(docx, PaddingRefusal);
        AssertRefused(EmptyCellMinWidth.Apply(docx), IndentRefusal);
        AssertRefusedWhereItReproduces(IndentFitToColumn.Apply(docx), PaddingRefusal);

        var text = PdfProbe.ExtractText(DocxToPdfConverter.Convert(docx));

        Assert.Contains("c70", text);
        Assert.Contains("indented paragraph text", text);
    }

    [Fact]
    public void ARefusalCanFollowARefusal_PaddingThenAPreferredWidth()
    {
        // html-000079: zeroing the padding only reaches the 1% table's own refusal.
        var docx = Package(FiftyNarrowColumns(SideMargins(108)) + ZeroWidthPercentTables());
        AssertRefused(docx, PaddingRefusal);
        AssertRefused(CellPaddingZero.Apply(docx), WidthRefusal);
        AssertRefused(TinyPercentWidthRepair.Apply(docx), PaddingRefusal);

        var text = PdfProbe.ExtractText(DocxToPdfConverter.Convert(docx));

        Assert.Contains("tiny percent table", text);
        Assert.Contains("end", text);
    }

    [Fact]
    public void ATopLevelTinyPercentTableIsNotChangedUpFront()
    {
        // 3.5 renders a top-level 1% table as it stands. TinyPercentWidthRepair used to run on every
        // conversion and rewrote it; it now waits for the refusal, so the output is the raw render.
        var docx = Package(
            "<w:tbl><w:tblPr><w:tblW w:w=\"50\" w:type=\"pct\"/></w:tblPr><w:tblGrid><w:gridCol/></w:tblGrid><w:tr>"
            + Cell(Para("sliver")) + "</w:tr></w:tbl>" + Para("end"));
        Assert.Null(RawRender.Refusal(docx));

        var raw = RawRender.Pdf(docx);

        // Every text run, in order, at the same place: the 1% cell wraps the words one to a line,
        // and the repair (automatic width) would put them on one.
        Assert.Equal(PdfProbe.TextRuns(raw), PdfProbe.TextRuns(DocxToPdfConverter.Convert(docx)));
    }

    [Fact]
    public void AnUnknownRefusalIsDiagnosedNotLooped()
    {
        var docx = Package(Para("text"));
        var unrelated = new ArgumentException("Something else entirely.");

        Assert.False(DocxPdfRefusalLadder.Knows(unrelated));
        Assert.Null(DocxPdfRefusalLadder.NextOrNull(unrelated, docx, new HashSet<int>()));
    }

    [Fact]
    public void ARefusalInsideAnotherException_IsKnown()
    {
        var wrapped = new InvalidOperationException("layout failed", new ArgumentException("Table " + PaddingRefusal + "."));

        Assert.True(DocxPdfRefusalLadder.Knows(wrapped));
    }

    [Fact]
    public void AStepThatChangesNothingIsSkipped()
    {
        // A padding refusal on a document with no empty cell goes straight to the zeroing step.
        var docx = Package(FiftyNarrowColumns(SideMargins(108)));
        var refusal = new ArgumentException("Table horizontal " + PaddingRefusal + ".");

        var next = DocxPdfRefusalLadder.NextOrNull(refusal, docx, new HashSet<int>());

        Assert.NotNull(next);
        Assert.Contains("<w:left w:w=\"0\"", Part(next, "word/document.xml"));
    }

    [Fact]
    public void AnEmptyCellIsGivenItsWidthBeforeAnyPaddingIsZeroed()
    {
        var docx = Package(NineColumnsWithAnEmptyOne());
        var refusal = new ArgumentException("Table horizontal " + PaddingRefusal + ".");

        var next = DocxPdfRefusalLadder.NextOrNull(refusal, docx, new HashSet<int>());

        Assert.NotNull(next);
        Assert.Contains("<w:left w:w=\"30\"", Part(next, "word/document.xml"));
        Assert.Contains("w:w=\"144\"", Part(next, "word/document.xml"));
    }

    [Fact]
    public void AStepRunsOnlyOnce()
    {
        var docx = Package(FiftyNarrowColumns(SideMargins(108)));
        var refusal = new ArgumentException("Table horizontal " + PaddingRefusal + ".");
        var used = new HashSet<int>();

        Assert.NotNull(DocxPdfRefusalLadder.NextOrNull(refusal, docx, used));
        Assert.Null(DocxPdfRefusalLadder.NextOrNull(refusal, docx, used));
    }

    [Fact]
    public void ARefusalNoStepCanChange_IsWrappedWithTheRenderersError()
    {
        // 60 automatic columns with the DEFAULT padding: refused with the padding message, but there is
        // no empty cell and no explicit margin to zero, so both padding steps find nothing.
        var docx = Package(AutoTable(60, "", (_, _) => Cell(Para("x")), rows: 1));
        AssertRefused(docx, PaddingRefusal);

        var failure = Assert.Throws<DocumentConversionException>(() => DocxToPdfConverter.Convert(docx));

        Assert.IsType<ArgumentException>(failure.InnerException);
        Assert.Contains(PaddingRefusal, failure.InnerException!.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ADocumentThatRendersIsNotRewritten()
    {
        // The ladder runs after a refusal and not before: nothing here may reach a repair.
        var docx = Package(AutoTable(3, SideMargins(108), (_, _) => Cell(Para("x"))));
        Assert.Null(RawRender.Refusal(docx));

        Assert.Contains("x", PdfProbe.ExtractText(DocxToPdfConverter.Convert(docx)));
    }

    // ---- a step that fails -------------------------------------------------------------------------

    /// <summary>
    /// Adds a styles part that no relationship points at, so the renderer never reads it, whose
    /// compressed data is invalid: only a repair that opens <c>word/styles.xml</c> meets it.
    /// </summary>
    private static byte[] WithCorruptUnreferencedStyles(byte[] docx)
    {
        using var ms = new MemoryStream();
        ms.Write(docx, 0, docx.Length);
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Update, leaveOpen: true))
        {
            using var writer = new StreamWriter(zip.CreateEntry("word/styles.xml").Open(), new UTF8Encoding(false));
            writer.Write("<w:styles " + Ns + ">" + string.Concat(Enumerable.Repeat("<w:style/>", 50)) + "</w:styles>");
        }

        var bytes = ms.ToArray();
        var name = "word/styles.xml"u8;
        // The entry's local header: 30 fixed bytes, then the name, then the extra field, then data.
        var at = bytes.AsSpan().IndexOf(name);
        var headerStart = at - 30;
        Assert.Equal(0x04034b50u, BitConverter.ToUInt32(bytes, headerStart));
        var dataStart = at + name.Length + BitConverter.ToUInt16(bytes, headerStart + 28);
        bytes[dataStart] = 0x07; // final block, block type 3: not a valid deflate block
        return bytes;
    }

    [Fact]
    public void AStepThatThrows_EndsTheLadderAsADocumentConversionException()
    {
        var docx = WithCorruptUnreferencedStyles(Package(FiftyNarrowColumns(SideMargins(108))));
        AssertRefused(docx, PaddingRefusal);
        // The empty-cell step has nothing to change; the next step opens the corrupt part and throws.
        Assert.Throws<InvalidDataException>(() => CellPaddingZero.Apply(docx));

        var failure = Assert.Throws<DocumentConversionException>(() => DocxToPdfConverter.Convert(docx));

        // The renderer's refusal is what is reported, not the step's own failure.
        Assert.IsType<ArgumentException>(failure.InnerException);
        Assert.Contains(PaddingRefusal, failure.InnerException!.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ExtremeMarginsDoNotOverflowTheColumnMeasure()
    {
        // Math.Abs(int.MinValue) throws; the margins are untrusted input.
        var section = "<w:sectPr><w:pgSz w:w=\"2147483647\" w:h=\"16838\"/><w:pgMar w:top=\"1440\" "
            + "w:right=\"-2147483648\" w:bottom=\"1440\" w:left=\"-2147483648\" w:header=\"720\" "
            + "w:footer=\"720\" w:gutter=\"0\"/><w:cols w:num=\"2147483647\" w:space=\"2147483647\"/></w:sectPr>";
        var docx = Package(IndentedParagraph("w:left=\"6480\""), section: section);

        Assert.Same(docx, IndentFitToColumn.Apply(docx));
    }
}
