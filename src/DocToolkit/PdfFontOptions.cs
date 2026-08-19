namespace DocToolkit;

/// <summary>
/// Fonts the caller supplies for characters the renderer cannot otherwise encode.
/// </summary>
/// <remarks>
/// <b>Without this, whether a document renders depends on the machine it renders on.</b> A document
/// containing Cyrillic, Greek or CJK converts on one host and is refused on another, because the
/// renderer falls back to whatever fonts that machine happens to have — measured, a Windows box
/// offers <c>Segoe UI Symbol</c> and <c>Segoe UI Emoji</c>, neither of which covers Cyrillic. That
/// machine-dependence is the one thing a package whose whole premise is "runs everywhere .NET does"
/// should not have.
///
/// <b>Nothing is shipped inside this package to fix it, deliberately.</b> A font covering Cyrillic,
/// Greek and CJK is measured in megabytes against a package measured in tens of kilobytes, and every
/// consumer would pay it to serve the minority converting non-Latin text — plus a licence family
/// this repository has never audited. So the bytes come from the caller, who already licenses the
/// typeface their documents are written in and knows which one that is.
///
/// <b>This is opt-in and changes nothing for anybody who does not use it.</b> Two effects are worth
/// knowing before you do.
///
/// <b>1. The fonts you supply REPLACE the host's own fallbacks; they do not add to them.</b> That
/// makes supplying too few actively worse than supplying none, which is the opposite of what the
/// name suggests. Measured over 99 real documents:
///
/// <list type="bullet">
/// <item><description>no font supplied — <b>71/99</b></description></item>
/// <item><description>one font (Arial) — <b>63/99</b>: it fixed the 4 documents needing Cyrillic and
/// broke 12 that the host's own fallbacks had been covering</description></item>
/// <item><description>four fonts — <b>77/99</b></description></item>
/// </list>
///
/// So supply fonts covering <i>everything your documents use</i>, not just the script that failed.
/// The refusal names the character it could not encode, which tells you what is still missing.
///
/// <b>2. It changes how fonts are embedded generally.</b> Measured on an ordinary Latin document,
/// output went from 128,755 bytes to 1,306 — the same base-14 swing this project already records as
/// varying PDF size a hundredfold. Both render correctly; the smaller leans on the standard fonts
/// every reader has.
///
/// <b>There is deliberately no compiled example for this type</b>, which is worth explaining
/// because every other public type here has one. This project's examples are real tests, so they
/// run - and a runnable example of this would need a font file committed to the repository, which is
/// the exact thing the paragraph above says not to do. An example that used a system font by path
/// would then be a claim that only holds on the machine that wrote it.
/// </remarks>
public sealed class PdfFontOptions
{
    private readonly List<(string Name, byte[] Data)> _fonts;

    /// <summary>Creates options carrying one font.</summary>
    /// <param name="fontName">
    /// A name for the font. It labels the font inside the PDF and in diagnostics; it does not have
    /// to match a font installed on the machine, because the bytes are what is used.
    /// </param>
    /// <param name="trueTypeFont">The TrueType or OpenType font file's bytes.</param>
    /// <exception cref="ArgumentException"><paramref name="fontName"/> is blank, or <paramref name="trueTypeFont"/> is empty.</exception>
    /// <exception cref="ArgumentNullException"><paramref name="trueTypeFont"/> is null.</exception>
    public PdfFontOptions(string fontName, byte[] trueTypeFont)
        : this(new List<(string, byte[])>())
    {
        _fonts.Add(Checked(fontName, trueTypeFont));
    }

    private PdfFontOptions(List<(string Name, byte[] Data)> fonts) => _fonts = fonts;

    /// <summary>
    /// Returns options carrying this font and one more.
    /// </summary>
    /// <remarks>
    /// Returns a new instance rather than mutating, so an options object handed to a converter
    /// cannot change under it — the same reasoning that makes every converter here static and
    /// stateless.
    ///
    /// <b>Order matters.</b> Fonts are offered to the renderer in the order added, so put the one
    /// covering the most of your text first.
    /// </remarks>
    /// <param name="fontName">A name for the font.</param>
    /// <param name="trueTypeFont">The TrueType or OpenType font file's bytes.</param>
    /// <exception cref="ArgumentException"><paramref name="fontName"/> is blank, or <paramref name="trueTypeFont"/> is empty.</exception>
    /// <exception cref="ArgumentNullException"><paramref name="trueTypeFont"/> is null.</exception>
    public PdfFontOptions Add(string fontName, byte[] trueTypeFont) =>
        new(new List<(string, byte[])>(_fonts) { Checked(fontName, trueTypeFont) });

    /// <summary>The names of the fonts these options carry, in the order they will be offered.</summary>
    public IReadOnlyList<string> FontNames => _fonts.Select(f => f.Name).ToList();

    private static (string, byte[]) Checked(string fontName, byte[] trueTypeFont)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(fontName);
        ArgumentNullException.ThrowIfNull(trueTypeFont);
        if (trueTypeFont.Length == 0)
            throw new ArgumentException("Font data was empty.", nameof(trueTypeFont));

        // Deliberately NOT validated as a font here. Deciding what is a valid TrueType file is the
        // renderer's job and it already reports one clearly; a second opinion in this package would
        // be one more thing to keep true as font formats change.
        return (fontName, trueTypeFont);
    }

    /// <summary>The renderer's own representation, or <see langword="null"/> when no font was given.</summary>
    internal OfficeIMO.Pdf.PdfEmbeddedFontFallbackSet? ToFallbackSet() =>
        _fonts.Count == 0
            ? null
            : new OfficeIMO.Pdf.PdfEmbeddedFontFallbackSet(
                _fonts.Select(f => new OfficeIMO.Pdf.PdfEmbeddedFontFallbackCandidate(f.Name, f.Data)));
}
