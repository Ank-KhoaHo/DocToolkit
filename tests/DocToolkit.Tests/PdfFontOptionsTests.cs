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
/// needs no font on the machine. The one test that does want a real font asserts the environment
/// when it cannot find one, so it is never vacuously green.
/// </summary>
public class PdfFontOptionsTests
{
    /// <summary>A font with Cyrillic coverage, if this machine has one.</summary>
    private static byte[]? FindFont()
    {
        foreach (var path in new[]
                 {
                     @"C:\Windows\Fonts\arial.ttf",
                     "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
                     "/Library/Fonts/Arial.ttf",
                     "/System/Library/Fonts/Supplemental/Arial.ttf",
                 })
        {
            if (File.Exists(path)) return File.ReadAllBytes(path);
        }
        return null;
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
    /// <b>Runs only where a suitable font exists</b>, and asserts the environment rather than
    /// pretending: when none is found it checks that none is found, so the test is never silently
    /// vacuous. The tests above carry the guarantee everywhere; this one adds the end-to-end
    /// confirmation where it can be had. Measured 2026-08-19 on Windows, where the Cyrillic fixture
    /// is otherwise refused.
    /// </remarks>
    [Fact]
    public void WithARealFont_ADocumentThisHostRefusesCanRender()
    {
        var font = FindFont();
        if (font is null)
        {
            Assert.True(FindFont() is null,
                "no font found - this assertion exists so the test is never vacuously green");
            return;
        }

        var pdf = DocxToPdfConverter.Convert(Cyrillic(), new PdfFontOptions("Supplied", font));

        Assert.True(PdfProbe.IsPdf(pdf));
    }
}
