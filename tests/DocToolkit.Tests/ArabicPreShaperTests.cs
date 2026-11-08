using System.IO.Compression;
using System.Text;
using System.Xml.Linq;

namespace DocToolkit.Tests;

/// <summary>
/// The Arabic pre-shaper (A124) turns logical Arabic-script letters into their contextual
/// presentation forms, still in logical order, one output string per input char.
/// </summary>
public class ArabicPreShaperTests
{
    private const string Ba = "\u0628";
    private const string Lam = "\u0644";
    private const string Alef = "\u0627";

    private static string[] Shape(string s) => ArabicPreShaper.Shape(s);

    [Fact]
    public void Lone_letter_takes_its_isolated_form()
        => Assert.Equal(new[] { "\uFE8F" }, Shape(Ba));

    [Fact]
    public void Two_joining_letters_take_initial_then_final()
        => Assert.Equal(new[] { "\uFE91", "\uFE90" }, Shape(Ba + Ba));

    [Fact]
    public void Three_joining_letters_take_initial_medial_final()
        => Assert.Equal(new[] { "\uFE91", "\uFE92", "\uFE90" }, Shape(Ba + Ba + Ba));

    [Fact]
    public void Lam_alef_is_one_ligature_and_leaves_the_alef_slot_empty()
        => Assert.Equal(new[] { "\uFEFB", "" }, Shape(Lam + Alef));

    [Fact]
    public void Lam_alef_after_a_joining_letter_takes_the_final_ligature()
        => Assert.Equal(new[] { "\uFE91", "\uFEFC", "" }, Shape(Ba + Lam + Alef));

    [Theory]
    [InlineData("\u0622", "\uFEF5")]
    [InlineData("\u0623", "\uFEF7")]
    [InlineData("\u0625", "\uFEF9")]
    [InlineData("\u0627", "\uFEFB")]
    public void Lam_alef_ligature_follows_the_alef_variant(string alef, string isolated)
    {
        var shaped = Shape(Lam + alef);
        Assert.Equal(isolated, shaped[0]);
        Assert.Equal("", shaped[1]);
        Assert.Equal(((char)(isolated[0] + 1)).ToString(), Shape(Ba + Lam + alef)[1]);
    }

    [Fact]
    public void A_right_joining_letter_breaks_the_chain()
        => Assert.Equal(new[] { "\uFEA9", "\uFE8F" }, Shape("\u062F" + Ba));

    [Fact]
    public void Persian_letters_take_their_presentation_forms()
    {
        Assert.Equal(new[] { "\uFB58", "\uFB57" }, Shape("\u067E\u067E"));   // pe
        Assert.Equal(new[] { "\uFB7C", "\uFB7B" }, Shape("\u0686\u0686"));   // che
        Assert.Equal(new[] { "\uFB8A" }, Shape("\u0698"));                    // zhe
        Assert.Equal(new[] { "\uFB94", "\uFB93" }, Shape("\u06AF\u06AF"));   // gaf
        Assert.Equal(new[] { "\uFBFE", "\uFBFD" }, Shape("\u06CC\u06CC"));   // farsi yeh
        Assert.Equal(new[] { "\uFB56" }, Shape("\u067E"));                    // lone pe is isolated
        Assert.Equal(new[] { "\uFBFC" }, Shape("\u06CC"));                    // lone yeh is isolated
        Assert.Equal(new[] { "\uFE91", "\uFB8B" }, Shape(Ba + "\u0698"));
    }

    [Fact]
    public void Harakat_are_transparent_to_joining()
    {
        // ba + fatha + ba: the first ba is initial and the second final, the fatha stays as is.
        Assert.Equal(new[] { "\uFE91", "\u064E", "\uFE90" }, Shape(Ba + "\u064E" + Ba));
        // dagger alef U+0670 is transparent too.
        Assert.Equal(new[] { "\uFE91", "\u0670", "\uFE90" }, Shape(Ba + "\u0670" + Ba));
    }

    [Fact]
    public void Lam_alef_joins_across_a_fatha()
        => Assert.Equal(new[] { "\uFEFB", "\u064E", "" }, Shape(Lam + "\u064E" + Alef));

    [Fact]
    public void Zero_width_non_joiner_breaks_joining()
        => Assert.Equal(new[] { "\uFE8F", "\u200C", "\uFE8F" }, Shape(Ba + "\u200C" + Ba));

    [Fact]
    public void Zero_width_non_joiner_breaks_a_lam_alef_ligature()
        => Assert.Equal(new[] { "\uFEDD", "\u200C", "\uFE8D" }, Shape(Lam + "\u200C" + Alef));

    [Fact]
    public void Tatweel_joins_and_stays_in_the_text()
        => Assert.Equal(new[] { "\uFE91", "\u0640", "\uFE90" }, Shape(Ba + "\u0640" + Ba));

    [Fact]
    public void Latin_and_digits_are_untouched_beside_Arabic()
    {
        const string text = "Invoice \u0641\u0627\u062A\u0648\u0631\u0629 12";
        var shaped = Shape(text);
        Assert.Equal("Invoice ", string.Concat(shaped.Take(8)));
        Assert.Equal(" 12", string.Concat(shaped.Skip(shaped.Length - 3)));
        // fa initial, alef final (breaks the chain), teh initial, waw final, reh isolated, teh marbuta isolated.
        Assert.Equal("\uFED3\uFE8E\uFE97\uFEEE\uFEAD\uFE93", string.Concat(shaped.Skip(8).Take(6)));
    }

    [Fact]
    public void Output_has_one_entry_per_input_char()
    {
        foreach (var s in new[] { "", "abc", Lam + Alef, Ba + Lam + Alef + "\u064E", "x \u0641\u0627\u062A\u0648\u0631\u0629 1" })
            Assert.Equal(s.Length, Shape(s).Length);
    }

    [Fact]
    public void A_word_joins_across_what_were_run_boundaries()
    {
        // kaf-alef split across two runs would be shaped as one container string by the caller.
        var shaped = Shape("\u06A9\u0627" + "\u0644\u0627");
        Assert.Equal(new[] { "\uFB90", "\uFE8E", "\uFEFB", "" }, shaped);
        // Same chars, ba + ba as two "runs": the join crosses the boundary.
        Assert.Equal(new[] { "\uFE91", "\uFE90" }, Shape(Ba + Ba));
    }

    // ---------------------------------------------------------------------------------------------
    // Apply: rewriting a package's text
    // ---------------------------------------------------------------------------------------------

    private const string WNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private const string SNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    private static byte[] Zip(params (string Name, string Xml)[] parts)
    {
        using var ms = new MemoryStream();
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var (name, xml) in parts)
            {
                using var w = new StreamWriter(zip.CreateEntry(name).Open(), new UTF8Encoding(false));
                w.Write(xml);
            }
        }
        return ms.ToArray();
    }

    private static string Part(byte[] package, string name)
    {
        using var zip = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        using var r = new StreamReader(zip.GetEntry(name)!.Open(), Encoding.UTF8);
        return r.ReadToEnd();
    }

    private static string Doc(string body)
        => $"<?xml version=\"1.0\" encoding=\"UTF-8\"?><w:document xmlns:w=\"{WNs}\"><w:body>{body}</w:body></w:document>";

    private static string[] WTexts(byte[] package, string part = "word/document.xml")
        => XDocument.Parse(Part(package, part), LoadOptions.PreserveWhitespace)
            .Descendants(XName.Get("t", WNs)).Select(t => t.Value).ToArray();

    private static bool IsPresentationForms(string s)
        => s.Length > 0 && s.All(c => c is >= '\uFB50' and <= '\uFEFC');

    [Fact]
    public void AWordSplitAcrossTwoRunsJoins()
    {
        // kaf + teh | ba: the join crosses the run boundary, so run 1 ends medial and run 2 is final.
        var docx = Zip(("word/document.xml", Doc("<w:p><w:r><w:t>\u0643\u062A</w:t></w:r><w:r><w:t>\u0628</w:t></w:r></w:p>")));
        Assert.Equal(new[] { "\uFEDB\uFE98", "\uFE90" }, WTexts(ArabicPreShaper.Apply(docx)));
    }

    [Fact]
    public void ALamAlefSplitAcrossRunsKeepsEachRunsOwnSlice()
    {
        // kaf alef | lam alef: the ligature takes the lam's slot, so run 2 shrinks by one char.
        var docx = Zip(("word/document.xml", Doc("<w:p><w:r><w:t>\u06A9\u0627</w:t></w:r><w:r><w:t>\u0644\u0627</w:t></w:r></w:p>")));
        Assert.Equal(new[] { "\uFB90\uFE8E", "\uFEFB" }, WTexts(ArabicPreShaper.Apply(docx)));
    }

    [Fact]
    public void ALamAlefAcrossTheRunBoundaryLeavesTheSecondRunEmpty()
    {
        var docx = Zip(("word/document.xml", Doc("<w:p><w:r><w:t>\u0644</w:t></w:r><w:r><w:t>\u0627</w:t></w:r></w:p>")));
        Assert.Equal(new[] { "\uFEFB", "" }, WTexts(ArabicPreShaper.Apply(docx)));
    }

    [Fact]
    public void ANestedTextBoxParagraphIsShapedOnItsOwn()
    {
        // The inner paragraph sits inside a run of the outer one. Merged, the two ba would join.
        var docx = Zip(("word/document.xml", Doc(
            "<w:p><w:r><w:t>\u0628</w:t></w:r><w:r><w:pict><w:txbxContent><w:p><w:r><w:t>\u0628</w:t></w:r></w:p></w:txbxContent></w:pict></w:r></w:p>")));
        Assert.Equal(new[] { "\uFE8F", "\uFE8F" }, WTexts(ArabicPreShaper.Apply(docx)));
    }

    [Fact]
    public void OnlyWTextIsShaped_AndSpacePreserveSurvives()
    {
        var docx = Zip(("word/document.xml", Doc(
            "<w:p><w:r><w:instrText>\u0628\u0628</w:instrText></w:r><w:del><w:r><w:delText>\u0628\u0628</w:delText></w:r></w:del>" +
            "<w:r><w:t xml:space=\"preserve\"> \u0628\u0628 </w:t></w:r></w:p>")));
        var result = ArabicPreShaper.Apply(docx);
        var xml = Part(result, "word/document.xml");
        Assert.Contains("<w:instrText>\u0628\u0628</w:instrText>", xml);
        Assert.Contains("<w:delText>\u0628\u0628</w:delText>", xml);
        Assert.Contains("xml:space=\"preserve\"", xml);
        Assert.Equal(new[] { " \uFE91\uFE90 " }, WTexts(result));
    }

    [Fact]
    public void MixedLatinAndArabicInOneRunKeepsLatin()
    {
        var docx = Zip(("word/document.xml", Doc("<w:p><w:r><w:t>Hi \u0628\u0628 ok</w:t></w:r></w:p>")));
        Assert.Equal(new[] { "Hi \uFE91\uFE90 ok" }, WTexts(ArabicPreShaper.Apply(docx)));
    }

    [Fact]
    public void HeadersFootnotesSheetsAndSlidesAreShaped()
    {
        const string ar = "\u0628\u0628";
        string header = $"<w:hdr xmlns:w=\"{WNs}\"><w:p><w:r><w:t>{ar}</w:t></w:r></w:p></w:hdr>";
        string footer = $"<w:ftr xmlns:w=\"{WNs}\"><w:p><w:r><w:t>{ar}</w:t></w:r></w:p></w:ftr>";
        string footnotes = $"<w:footnotes xmlns:w=\"{WNs}\"><w:footnote w:id=\"1\"><w:p><w:r><w:t>{ar}</w:t></w:r></w:p></w:footnote></w:footnotes>";
        string endnotes = $"<w:endnotes xmlns:w=\"{WNs}\"><w:endnote w:id=\"1\"><w:p><w:r><w:t>{ar}</w:t></w:r></w:p></w:endnote></w:endnotes>";
        string shared = $"<sst xmlns=\"{SNs}\"><si><t>{ar}</t></si></sst>";
        string sheet = $"<worksheet xmlns=\"{SNs}\"><sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>{ar}</t></is></c></row></sheetData></worksheet>";
        string slide = $"<p:sld xmlns:p=\"urn:p\" xmlns:a=\"{ANs}\"><a:p><a:r><a:t>{ar}</a:t></a:r></a:p></p:sld>";
        var pkg = Zip(
            ("word/document.xml", Doc("<w:p/>")),
            ("word/header1.xml", header), ("word/footer2.xml", footer),
            ("word/footnotes.xml", footnotes), ("word/endnotes.xml", endnotes),
            ("xl/sharedStrings.xml", shared), ("xl/worksheets/sheet1.xml", sheet),
            ("ppt/slides/slide1.xml", slide));

        var result = ArabicPreShaper.Apply(pkg);

        foreach (var name in new[] { "word/header1.xml", "word/footer2.xml", "word/footnotes.xml", "word/endnotes.xml",
                     "xl/sharedStrings.xml", "xl/worksheets/sheet1.xml", "ppt/slides/slide1.xml" })
        {
            var text = XDocument.Parse(Part(result, name)).Descendants().Single(e => e.Name.LocalName == "t").Value;
            Assert.True(IsPresentationForms(text), $"{name}: {text}");
        }
    }

    [Fact]
    public void PartsOutsideTheListAreLeftAlone()
    {
        string layout = $"<p:sldLayout xmlns:p=\"urn:p\" xmlns:a=\"{ANs}\"><a:p><a:r><a:t>\u0628\u0628</a:t></a:r></a:p></p:sldLayout>";
        string slide = $"<p:sld xmlns:p=\"urn:p\" xmlns:a=\"{ANs}\"><a:p><a:r><a:t>\u0628\u0628</a:t></a:r></a:p></p:sld>";
        var result = ArabicPreShaper.Apply(Zip(("ppt/slides/slide1.xml", slide), ("ppt/slideLayouts/slideLayout1.xml", layout)));
        Assert.Equal(layout, Part(result, "ppt/slideLayouts/slideLayout1.xml"));
        Assert.Contains("\uFE91\uFE90", Part(result, "ppt/slides/slide1.xml"));
    }

    [Fact]
    public void ASharedStringIsShapedOnceAndStaysShared()
    {
        string shared = $"<sst xmlns=\"{SNs}\" count=\"2\" uniqueCount=\"1\"><si><t>\u0628\u0628</t></si></sst>";
        string sheet = $"<worksheet xmlns=\"{SNs}\"><sheetData><row r=\"1\"><c r=\"A1\" t=\"s\"><v>0</v></c><c r=\"B1\" t=\"s\"><v>0</v></c></row></sheetData></worksheet>";
        var result = ArabicPreShaper.Apply(Zip(("xl/sharedStrings.xml", shared), ("xl/worksheets/sheet1.xml", sheet)));

        var si = XDocument.Parse(Part(result, "xl/sharedStrings.xml")).Descendants(XName.Get("si", SNs)).ToList();
        Assert.Single(si);
        Assert.Equal("\uFE91\uFE90", si[0].Value);
        var cells = XDocument.Parse(Part(result, "xl/worksheets/sheet1.xml")).Descendants(XName.Get("v", SNs)).Select(v => v.Value);
        Assert.Equal(new[] { "0", "0" }, cells);
    }

    [Fact]
    public void ASharedStringsRichRunsJoinAndPhoneticTextIsSkipped()
    {
        string shared = $"<sst xmlns=\"{SNs}\"><si><r><t>\u0643\u062A</t></r><r><t>\u0628</t></r><rPh sb=\"0\" eb=\"1\"><t>\u0628\u0628</t></rPh></si></sst>";
        var result = ArabicPreShaper.Apply(Zip(("xl/sharedStrings.xml", shared)));
        var ts = XDocument.Parse(Part(result, "xl/sharedStrings.xml")).Descendants(XName.Get("t", SNs)).Select(t => t.Value).ToArray();
        Assert.Equal(new[] { "\uFEDB\uFE98", "\uFE90", "\u0628\u0628" }, ts);
    }

    [Fact]
    public void NoArabic_ReturnsTheSameArray()
    {
        var docx = Zip(("word/document.xml", Doc("<w:p><w:r><w:t>Hello \u05E9\u05DC\u05D5\u05DD</w:t></w:r></w:p>")));
        Assert.Same(docx, ArabicPreShaper.Apply(docx));
    }

    [Fact]
    public void NotAZip_ReturnsTheSameArray()
    {
        var junk = new byte[] { 1, 2, 3 };
        Assert.Same(junk, ArabicPreShaper.Apply(junk));
    }

    [Theory]
    [InlineData(new byte[] { 0x41, 0x4B, 3, 4 })] // first byte is not 'P'
    [InlineData(new byte[] { 0x50, 0x58, 3, 4 })] // second byte is not 'K'
    public void FourBytesThatAreNotAZip_ReturnTheSameArray(byte[] junk)
        => Assert.Same(junk, ArabicPreShaper.Apply(junk));

    [Theory]
    [InlineData("<w:br/>")]
    [InlineData("<w:cr/>")]
    [InlineData("<w:tab/>")]
    [InlineData("<w:ptab w:relativeTo=\"margin\" w:alignment=\"left\" w:leader=\"none\"/>")]
    [InlineData("<w:noBreakHyphen/>")]
    [InlineData("<w:softHyphen/>")]
    [InlineData("<w:sym w:font=\"Wingdings\" w:char=\"F0FC\"/>")]
    [InlineData("<w:fldChar w:fldCharType=\"begin\"/>")]
    [InlineData("<w:drawing/>")]
    [InlineData("<w:object/>")]
    [InlineData("<w:pict/>")]
    public void LettersDoNotJoinAcrossBreaksTabsOrInlineObjects(string between)
    {
        var docx = Zip(("word/document.xml", Doc(
            $"<w:p><w:r><w:t>بب</w:t></w:r><w:r>{between}</w:r><w:r><w:t>بب</w:t></w:r></w:p>")));
        Assert.Equal(new[] { "ﺑﺐ", "ﺑﺐ" }, WTexts(ArabicPreShaper.Apply(docx)));
    }

    [Fact]
    public void LettersDoNotJoinAcrossADrawingMLLineBreak()
    {
        string slide = $"<p:sld xmlns:p=\"urn:p\" xmlns:a=\"{ANs}\"><a:p><a:r><a:t>بب</a:t></a:r><a:br/><a:r><a:t>بب</a:t></a:r></a:p></p:sld>";
        var result = ArabicPreShaper.Apply(Zip(("ppt/slides/slide1.xml", slide)));
        var ts = XDocument.Parse(Part(result, "ppt/slides/slide1.xml")).Descendants(XName.Get("t", ANs)).Select(t => t.Value);
        Assert.Equal(new[] { "ﺑﺐ", "ﺑﺐ" }, ts);
    }

    [Fact]
    public void ALatinOnlyNodeBesideAnArabicNodeIsLeftExactly()
    {
        var docx = Zip(("word/document.xml", Doc(
            "<w:p><w:r><w:t xml:space=\"preserve\">Hi </w:t></w:r><w:r><w:t>بب</w:t></w:r><w:r><w:t xml:space=\"preserve\">   </w:t></w:r></w:p>")));
        Assert.Equal(new[] { "Hi ", "ﺑﺐ", "   " }, WTexts(ArabicPreShaper.Apply(docx)));
    }

    [Fact]
    public void TextElementsOfAnotherNamespaceBelongToNoContainer()
    {
        // An s:t under a w:p, a w:t under an a:p and a w:t under an unrelated p are not that
        // container's text, so none is shaped. The w:p's own w:t is.
        string doc = $"<w:document xmlns:w=\"{WNs}\" xmlns:s=\"{SNs}\" xmlns:a=\"{ANs}\" xmlns:x=\"urn:x\"><w:body>" +
            "<w:p><w:r><s:t>بب</s:t></w:r><w:r><w:t>بب</w:t></w:r></w:p>" +
            "<a:p><w:t>بب</w:t></a:p>" +
            "<x:p><w:t>بب</w:t></x:p>" +
            "</w:body></w:document>";
        var xml = XDocument.Parse(Part(ArabicPreShaper.Apply(Zip(("word/document.xml", doc))), "word/document.xml"));
        var ts = xml.Descendants().Where(e => e.Name.LocalName == "t").Select(t => t.Value).ToArray();
        Assert.Equal(new[] { "بب", "ﺑﺐ", "بب", "بب" }, ts);
    }

    [Fact]
    public void AParagraphOfAnotherNamespaceIsNotAContainer()
    {
        var docx = Zip(("word/document.xml",
            $"<w:document xmlns:w=\"{WNs}\" xmlns:x=\"urn:x\"><w:body><x:p><w:r><w:t>بب</w:t></w:r></x:p></w:body></w:document>"));
        Assert.Same(docx, ArabicPreShaper.Apply(docx));
    }

    [Fact]
    public void PartNamesAreMatchedWithoutCase_AndRelsEntriesAreNotParts()
    {
        string header = $"<w:hdr xmlns:w=\"{WNs}\"><w:p><w:r><w:t>بب</w:t></w:r></w:p></w:hdr>";
        var result = ArabicPreShaper.Apply(Zip(
            ("word/document.xml", Doc("<w:p/>")),
            ("Word/Header1.xml", header),
            ("word/header1.xml.rels", header)));
        Assert.Contains("ﺑﺐ", Part(result, "Word/Header1.xml"));
        Assert.Equal(header, Part(result, "word/header1.xml.rels"));
    }

    [Fact]
    public void AlreadyShapedText_ReturnsTheSameArray()
    {
        var docx = Zip(("word/document.xml", Doc("<w:p><w:r><w:t>\uFE91\uFE90</w:t></w:r></w:p>")));
        Assert.Same(docx, ArabicPreShaper.Apply(docx));
    }
}
