using System.Text.RegularExpressions;

namespace DocToolkit.Tests;

/// <summary>
/// Fonts a caller supplies for characters the renderer cannot otherwise encode.
///
/// <b>Measured 2026-08-19 on the Cyrillic fixture: it is refused on this host with no font supplied,
/// and renders when one is.</b> The refusal names the fallbacks the machine offered — on Windows,
/// <c>Segoe UI Symbol</c> and <c>Segoe UI Emoji</c>, neither of which covers Cyrillic — which is the
/// host-dependence this option exists to remove.
///
/// <b>Almost everything here is host-independent, which took a rethink.</b> This project ships no
/// font, for the licence and size reasons recorded on <see cref="PdfFontOptions"/>, and the linux
/// container job runs on a bare SDK image with none installed — so an end-to-end test needing a real
/// font would be skipped there, and a skipped test is one that quietly does nothing.
///
/// <b>The proof that the bytes reach the renderer comes from INVALID ones instead.</b> The renderer
/// validates font data eagerly, so its refusal — in its own words, which nothing else in this
/// package could produce — is exactly the evidence a successful render would have given, and it
/// needs no font on the machine. The tests that do want a real font pass WITHOUT ASSERTING on a host
/// that has none - unless <c>DOCTOOLKIT_REQUIRE_FONT=1</c> is set, which makes a missing font a
/// failure. CI sets it on the runners that have fonts, so they can never quietly stop running there.
/// </summary>
public class PdfFontOptionsTests
{
    /// <summary>
    /// A font covering Cyrillic, Hebrew and Arabic, if this machine has one; a FAILURE instead when
    /// <c>DOCTOOLKIT_REQUIRE_FONT=1</c> says this host must have one.
    /// </summary>
    /// <remarks>
    /// <b>The switch is what keeps the real-font tests honest.</b> Without it a font-less host - the
    /// linux container job's image is one - runs every one of them to an early return with nothing
    /// asserted, and a change to a runner's image would switch them off with the suite still green.
    /// </remarks>
    internal static byte[]? FindFont()
    {
        var found = new[]
        {
            @"C:\Windows\Fonts\arial.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/Library/Fonts/Arial.ttf",
            "/System/Library/Fonts/Supplemental/Arial.ttf",
        }.FirstOrDefault(File.Exists);

        if (found is null && Environment.GetEnvironmentVariable("DOCTOOLKIT_REQUIRE_FONT") == "1")
            Assert.Fail("DOCTOOLKIT_REQUIRE_FONT=1, but none of the known font paths exists on this host.");

        return found is null ? null : File.ReadAllBytes(found);
    }

    private static byte[] Cyrillic() =>
        File.ReadAllBytes(Path.Join(AppContext.BaseDirectory, "assets", "word-cyrillic.docx"));

    // ---- the options type, tested everywhere -------------------------------------------------------

    [Fact]
    public void ItCarriesTheFontsItWasGiven()
    {
        var fonts = new PdfFontOptions("First", [1, 2, 3]).Add("Second", [4, 5, 6]);

        Assert.Equal(["First", "Second"], fonts.FontNames);
    }

    [Fact]
    public void AddReturnsANewInstance_SoOptionsCannotChangeUnderAConverter()
    {
        // The same reasoning that makes every converter here static and stateless: an options object
        // handed over must not mutate afterwards.
        var one = new PdfFontOptions("First", [1]);
        var two = one.Add("Second", [2]);

        Assert.Single(one.FontNames);
        Assert.Equal(2, two.FontNames.Count);
        Assert.NotSame(one, two);
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void ABlankFontNameIsRefused(string name)
    {
        Assert.Throws<ArgumentException>(() => new PdfFontOptions(name, [1, 2, 3]));
    }

    [Fact]
    public void EmptyFontDataIsRefused()
    {
        Assert.Throws<ArgumentException>(() => new PdfFontOptions("Name", []));
    }

    [Fact]
    public void NullFontDataIsRefused()
    {
        Assert.Throws<ArgumentNullException>(() => new PdfFontOptions("Name", null!));
    }

    [Fact]
    public void AddValidatesTheSameWayTheConstructorDoes()
    {
        // Two validation paths that disagree is the drift the *Core convention exists to prevent,
        // and this type has two entry points for the same thing.
        var fonts = new PdfFontOptions("First", [1]);

        Assert.Throws<ArgumentException>(() => fonts.Add("", [1]));
        Assert.Throws<ArgumentException>(() => fonts.Add("Name", []));
        Assert.Throws<ArgumentNullException>(() => fonts.Add("Name", null!));
    }

    // ---- the converters accept it, and it changes nothing when unused --------------------------------

    [Fact]
    public void TheNoFontsOverloadIsUnchanged()
    {
        // The old signature must keep behaving exactly as it did - this is the guarantee that lets
        // the new overload be additive.
        var docx = DocxEditor.Create([DocxBlock.Paragraph("Plain text.")]);

        Assert.True(PdfProbe.IsPdf(DocxToPdfConverter.Convert(docx)));
    }

    [Fact]
    public async Task NullFontsBehavesLikeTheOriginalOverload()
    {
        var docx = DocxEditor.Create([DocxBlock.Paragraph("Plain text.")]);

        Assert.True(PdfProbe.IsPdf(DocxToPdfConverter.Convert(docx, null)));

        using var source = new MemoryStream(docx);
        using var destination = new MemoryStream();
        await DocxToPdfConverter.ConvertAsync(source, destination, null);
        Assert.True(PdfProbe.IsPdf(destination.ToArray()));
    }

    [Fact]
    public async Task TheHtmlOverloadRefusesNulls()
    {
        await Assert.ThrowsAsync<ArgumentNullException>(
            () => HtmlToPdfConverter.ConvertAsync(null!, new PdfFontOptions("N", [1])));
        await Assert.ThrowsAsync<ArgumentNullException>(
            () => HtmlToPdfConverter.ConvertAsync("<p>x</p>", (PdfFontOptions)null!));
    }

    // ---- the bytes reach the renderer, proved without needing a font on the host --------------------

    /// <summary>
    /// Invalid font bytes fail with the RENDERER's own message.
    /// </summary>
    /// <remarks>
    /// <b>This is the end-to-end proof, and it is host-independent, which is why it is shaped like
    /// this.</b> A test that rendered successfully with a supplied font would need a real font file
    /// - and this project ships none, on purpose, while the linux container job runs on a bare SDK
    /// image with no fonts installed. Skipping it there would be a test that quietly does nothing.
    ///
    /// The renderer validates font data eagerly, so its refusal is proof the bytes were handed over:
    /// nothing else in this package could produce that message. It also documents that
    /// <see cref="PdfFontOptions"/> deliberately does not validate the format itself - the renderer
    /// already says it better than a second opinion here would.
    /// </remarks>
    [Theory]
    [InlineData(3)]      // too small to be a font at all
    [InlineData(1024)]   // large enough, but not a TrueType file
    public void InvalidFontBytesFailWithTheRenderersOwnMessage(int size)
    {
        var docx = DocxEditor.Create([DocxBlock.Paragraph("Plain Latin text.")]);

        var ex = Assert.Throws<DocumentConversionException>(
            () => DocxToPdfConverter.Convert(docx, new PdfFontOptions("Fake", new byte[size])));

        Assert.Contains("TrueType", ex.InnerException!.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void TheRefusalWarnsThatFontsREPLACETheHostsOwn()
    {
        // Measured over 99 real documents: one font took them from 71/99 to 63/99, fixing the four
        // needing Cyrillic and breaking twelve the host's own fallbacks had been covering. Somebody
        // reading the refusal is about to supply exactly one font, so the message has to say this.
        var described = DocxPdfFailureDiagnosis.Describe(
            new ArgumentException("Text contains character U+0421 that is not covered by any embedded font."));

        Assert.Contains("replace", described!, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("too few is worse than none", described, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void TheFallbackSetCarriesEveryFontInOrder()
    {
        // The conversion into the renderer's own type, checked directly - the last step before the
        // bytes leave this package.
        var set = new PdfFontOptions("First", [1, 2, 3]).Add("Second", [4, 5, 6]).ToFallbackSet();

        Assert.NotNull(set);
        Assert.Equal(["First", "Second"], set!.Candidates.Select(c => c.FontName));
    }

    /// <summary>
    /// A real font renders a document this host would otherwise refuse.
    /// </summary>
    /// <remarks>
    /// <b>Runs only where a suitable font exists</b>; see <see cref="FindFont"/> for the switch that
    /// makes a missing one a failure. The tests above carry the guarantee everywhere; this one adds the end-to-end
    /// confirmation where it can be had. Measured 2026-08-19 on Windows, where the Cyrillic fixture
    /// is otherwise refused.
    /// </remarks>
    [Fact]
    public void WithARealFont_ADocumentThisHostRefusesCanRender()
    {
        var font = FindFont();
        if (font is null)
            return;

        var pdf = DocxToPdfConverter.Convert(Cyrillic(), new PdfFontOptions("Supplied", font));

        Assert.True(PdfProbe.IsPdf(pdf));
    }

    // ---- A122: supplied fonts also turn on right-to-left shaping -------------------------------------

    [Fact]
    public void ToPdfOptions_CarriesTheFallbacksAndTurnsOnShaping()
    {
        var options = new PdfFontOptions("First", [1, 2, 3]).ToPdfOptions();

        Assert.Equal(["First"], options.EmbeddedFontFallbacks!.Candidates.Select(c => c.FontName));
        Assert.NotNull(options.TextShapingProvider);
    }

    /// <summary>The words of the Hebrew test line, in the order a reader meets them (right to left).</summary>
    internal static readonly string[] HebrewLine = ["חשבונית", "מס", "מספר", "12"];

    /// <summary>
    /// Asserts the PDF draws <see cref="HebrewLine"/> as a right-to-left line: the first word is the
    /// rightmost run, the last the leftmost. Judged by run x position, never by eye.
    /// </summary>
    internal static void AssertHebrewLineIsInReadingOrder(byte[] pdf) =>
        AssertRunsAreTheHebrewLine(PdfProbe.TextRuns(pdf));

    /// <summary>
    /// The exact run strings, by x descending. Measured 2026-11-08 (OfficeIMO 3.5.3, DejaVu): Hebrew
    /// needs no letter joining, so shaped and unshaped bytes are equal and "bytes differ" can no
    /// longer be the test. The line is drawn visually: one run per word, glyphs left to right, so a
    /// Hebrew word's run text is the word BACKWARDS (תינובשח, סמ, רפסמ) and digits stay as they are
    /// (12). The runs sit at x 141.5, 122.2, 88.7 and 72.0. The same words fed in reversed order land
    /// the other way round (12 at 170.1). A Hebrew word reading forwards in a run would have its first
    /// letter leftmost - the letters-reversed defect - so it is rejected, as is "21".
    /// </summary>
    internal static void AssertRunsAreTheHebrewLine(IEnumerable<(string Text, double X)> runs)
    {
        var drawn = runs
            .Where(r => r.Text.Trim().Length > 0)
            .OrderByDescending(r => r.X)
            .ToArray();
        var seen = string.Join(" | ", drawn.Select(r => $"{r.Text} @{r.X}"));
        var expected = HebrewLine
            .Select(w => w.All(char.IsAsciiDigit) ? w : new string(w.Reverse().ToArray()))
            .ToArray();

        Assert.True(drawn.Select(r => r.Text.Trim()).SequenceEqual(expected),
            $"runs by x descending should read {string.Join(" | ", expected)}, got: {seen}");
    }

    [Fact]
    public void HebrewLineAssertion_AcceptsTheMeasuredRuns()
    {
        AssertRunsAreTheHebrewLine([("תינובשח", 141.5), (" ", 138.2), ("סמ", 122.2), ("רפסמ", 88.7), ("12", 72.0)]);
    }

    [Fact]
    public void HebrewLineAssertion_RejectsWordsInReversedOrder()
    {
        // Measured: the reversed input lands 12 rightmost (170.1) and חשבונית leftmost (72.0).
        Assert.ThrowsAny<Exception>(() => AssertRunsAreTheHebrewLine(
            [("12", 170.1), ("רפסמ", 136.5), ("סמ", 117.3), ("תינובשח", 72.0)]));
    }

    [Fact]
    public void HebrewLineAssertion_RejectsLettersInLogicalOrder()
    {
        // Right word order, but each Hebrew word reads forwards: first letter leftmost.
        Assert.ThrowsAny<Exception>(() => AssertRunsAreTheHebrewLine(
            [("חשבונית", 141.5), ("מס", 122.2), ("מספר", 88.7), ("12", 72.0)]));
    }

    [Fact]
    public void HebrewLineAssertion_RejectsReversedDigits()
    {
        Assert.ThrowsAny<Exception>(() => AssertRunsAreTheHebrewLine(
            [("תינובשח", 141.5), ("סמ", 122.2), ("רפסמ", 88.7), ("21", 72.0)]));
    }

    /// <summary>Hebrew rendered with a supplied font reads right to left, in word order.</summary>
    /// <remarks>
    /// Same environment rule as the test above - see <see cref="FindFont"/>.
    /// </remarks>
    [Fact]
    public void WithARealFont_HebrewDocxIsInReadingOrder()
    {
        var font = FindFont();
        if (font is null)
            return;

        var docx = DocxEditor.Create([DocxBlock.Paragraph(string.Join(" ", HebrewLine))]);

        AssertHebrewLineIsInReadingOrder(DocxToPdfConverter.Convert(docx, new PdfFontOptions("Supplied", font)));
    }

    /// <summary>Shaping leaves Latin text byte-for-byte as it was - measured, and pinned here.</summary>
    [Fact]
    public void WithARealFont_LatinDocxBytesAreUnchangedByShaping()
    {
        var font = FindFont();
        if (font is null)
            return;

        var fonts = new PdfFontOptions("Supplied", font);
        var docx = DocxEditor.Create([DocxBlock.Paragraph("Invoice total, office fi fl ffi: 60,000")]);

        Assert.Equal(UnshapedRender.Docx(docx, fonts), DocxToPdfConverter.Convert(docx, fonts));
    }

    // ---- Arabic joins when fonts are supplied (A124) ------------------------------------------------

    internal const string ArabicLine = "فاتورة ضريبية";

    /// <summary>
    /// Asserts <paramref name="pdf"/> is the Arabic line drawn JOINED, given the same document drawn
    /// unshaped.
    /// </summary>
    /// <remarks>
    /// <b>Measured 2026-11-09 (OfficeIMO 3.5.3, DejaVu), and this is why the assertion is on glyph
    /// positions.</b> The decoded text is identical shaped or not: the ToUnicode map gives the base
    /// letters back for presentation forms, so a text comparison cannot tell. And the renderer emits
    /// no <c>/ActualText</c> at all, so there is no logical string to compare. What changes is the
    /// advance of each glyph, so where the run starts: 132.0 unshaped, 109.2 shaped, for the first
    /// DOCX word.
    /// </remarks>
    internal static void AssertArabicLineIsJoined(byte[] pdf, byte[] unshapedPdf)
    {
        var words = PdfProbe.TextRuns(pdf).Where(r => r.Text.Trim().Length > 0).Take(2).ToList();
        var unshaped = PdfProbe.TextRuns(unshapedPdf).Where(r => r.Text.Trim().Length > 0).Take(2).ToList();

        // Same two words, in the same code points (a word's run is its letters backwards)...
        Assert.Equal(unshaped.Select(r => r.Text), words.Select(r => r.Text));
        Assert.Equal(
            ArabicLine.Split(' ').Select(w => new string(w.Reverse().ToArray())).ToArray(),
            words.Select(r => r.Text).ToArray());
        // ...but the first starts somewhere else, because its glyphs are wider or narrower.
        Assert.True(
            unshaped[0].X != words[0].X,
            $"The first word starts at {words[0].X} shaped and unshaped alike, so pre-shaping changed no glyph."
            + $"\nShaped:   {TextEvidence(pdf)}\nUnshaped: {TextEvidence(unshapedPdf)}");
    }

    /// <summary>
    /// The fonts a PDF names and its first text-showing operators, raw, so a failure on a host this
    /// repository cannot run locally (macOS) shows which glyph codes each render drew.
    /// </summary>
    private static string TextEvidence(byte[] pdf)
    {
        var raw = PdfProbe.Raw(pdf);
        var fonts = Regex.Matches(raw, @"/BaseFont\s*/[^\s/>\]]+").Select(m => m.Value).Distinct();
        var shows = Regex.Matches(raw, @"(\[[^\]]{0,400}\]\s*TJ|<[0-9A-Fa-f]{1,400}>\s*Tj|\([^)]{0,200}\)\s*Tj)")
            .Take(6)
            .Select(m => m.Value);
        return string.Join(" ", fonts) + " | " + string.Join(" ; ", shows);
    }

    private static byte[] ArabicDocx() => DocxEditor.Create([DocxBlock.Paragraph(ArabicLine)]);

    [Fact]
    public void WithFonts_ArabicDocxIsPreShaped()
    {
        var font = ArabicFontTests.FindArabicFont();
        if (font is null)
            return;
        var fonts = new PdfFontOptions("Supplied", font);
        var docx = ArabicDocx();

        var pdf = DocxToPdfConverter.Convert(docx, fonts);

        AssertArabicLineIsJoined(pdf, UnshapedRender.Docx(docx, fonts));
        // And it is exactly the render of the pre-shaped package, so nothing else changed.
        Assert.Equal(UnshapedRender.Docx(ArabicPreShaper.Apply(docx), fonts), pdf);
    }

    /// <summary>The stream overload reaches the same call site as the byte[] one.</summary>
    [Fact]
    public async Task WithFonts_ArabicDocxThroughAStreamOverloadIsPreShaped()
    {
        var font = ArabicFontTests.FindArabicFont();
        if (font is null)
            return;
        var fonts = new PdfFontOptions("Supplied", font);
        var docx = ArabicDocx();
        using var destination = new MemoryStream();

        await DocxToPdfConverter.ConvertAsync(new MemoryStream(docx), destination, fonts);

        AssertArabicLineIsJoined(destination.ToArray(), UnshapedRender.Docx(docx, fonts));
    }

    /// <summary>HTML reaches the renderer through DOCX but not through Convert; it joins too.</summary>
    [Fact]
    public async Task WithFonts_ArabicHtmlIsPreShaped()
    {
        var font = ArabicFontTests.FindArabicFont();
        if (font is null)
            return;
        var fonts = new PdfFontOptions("Supplied", font);
        var html = $"<html><body><p>{ArabicLine}</p></body></html>";

        var pdf = await HtmlToPdfConverter.ConvertAsync(html, fonts);

        var unshapedDocx = await HtmlToDocxConverter.ConvertAsync(html, PageSetup.A4);
        AssertArabicLineIsJoined(pdf, UnshapedRender.Docx(unshapedDocx, fonts));
    }

    /// <summary>
    /// No fonts: the pre-shaper is never reached. Rendering the pre-shaped package gives a different
    /// outcome (here a refusal naming the first shaped letter, or a different PDF on a host whose own
    /// fonts cover Arabic), so a converter that shaped without fonts would make the two equal.
    /// </summary>
    [Fact]
    public void WithoutFonts_ArabicDocxIsNotPreShaped()
    {
        var docx = ArabicDocx();

        Assert.NotEqual(Outcome(ArabicPreShaper.Apply(docx)), Outcome(docx));
    }

    /// <summary>What the no-fonts render of this package does on this host: the PDF, or its refusal.</summary>
    private static string Outcome(byte[] docx)
    {
        try
        {
            return System.Convert.ToBase64String(DocxToPdfConverter.Convert(docx));
        }
        catch (DocumentConversionException ex)
        {
            return "refused: " + (ex.InnerException?.Message ?? ex.Message);
        }
    }
}
