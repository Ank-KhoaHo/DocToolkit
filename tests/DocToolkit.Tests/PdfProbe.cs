using System.Globalization;
using System.Text;
using System.Linq;
using System.Text.RegularExpressions;

namespace DocToolkit.Tests;

/// <summary>
/// Reads facts out of a generated PDF for assertions.
///
/// IMPORTANT: OfficeIMO writes content streams UNCOMPRESSED (no /Filter) and emits text as
/// hex-string operators, e.g. "&lt;41636D65&gt; Tj" == "Acme". So neither inflating the streams
/// nor substring-searching the raw bytes finds any text - both return nothing and look exactly
/// like a broken converter. Always go through this helper.
/// </summary>
public static class PdfProbe
{
    private static readonly Regex HexText = new(@"<([0-9A-Fa-f]+)>\s*Tj", RegexOptions.Compiled);

    // /ToUnicode CMap sections. Singleline so the bodies can span lines, which they always do.
    private static readonly Regex BfCharSection =
        new(@"beginbfchar(.*?)endbfchar", RegexOptions.Compiled | RegexOptions.Singleline);

    private static readonly Regex BfCharPair =
        new(@"<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>", RegexOptions.Compiled);

    private static readonly Regex BfRangeSection =
        new(@"beginbfrange(.*?)endbfrange", RegexOptions.Compiled | RegexOptions.Singleline);

    private static readonly Regex BfRangeEntry =
        new(@"<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>", RegexOptions.Compiled);

    // A dictionary that may contain up to one level of nested "<< ... >>" (e.g. a Catalog's
    // /Names dictionary, or a Page's /Resources dictionary), and/or single-angle-bracket hex
    // string tokens (e.g. a trailer's "/ID [<HEX> <HEX>]"). This is not a full PDF parser - it
    // doesn't handle arbitrarily deep nesting - but it's enough to keep a match from spilling
    // out of the dictionary it started in and into an unrelated neighboring object, which a
    // plain non-greedy ".*?" (as the old implementation used) does not guard against.
    //
    // The "<[0-9A-Fa-f]*>" branch matters in practice: every real trailer PdfProbe has been
    // pointed at (OfficeIMO's writer, at minimum) emits "/ID [<HEX> <HEX>]" in the trailer
    // dictionary. Without this branch, TrailerDict below fails to match ANY real trailer, so
    // TryPageCountViaCatalog silently returns null on every real PDF and PageCount always falls
    // through to the max-/Count scan - the "preferred" catalog-resolution path documented below
    // was dead code on real input until this branch was added.
    private const string DictBody = @"<<(?:[^<>]|<<[^<>]*>>|<[0-9A-Fa-f]*>)*>>";
    private static readonly Regex TrailerDict = new(@"trailer\s*(" + DictBody + ")",
                                                     RegexOptions.Compiled | RegexOptions.Singleline);
    private static readonly Regex RootRef = new(@"/Root\s+(\d+)\s+\d+\s+R", RegexOptions.Compiled);
    private static readonly Regex PagesRef = new(@"/Pages\s+(\d+)\s+\d+\s+R", RegexOptions.Compiled);
    private static readonly Regex TypeIsPages = new(@"/Type\s*/Pages\b", RegexOptions.Compiled);
    private static readonly Regex CountField = new(@"/Count\s+(\d+)", RegexOptions.Compiled);
    private static readonly Regex AnyDict = new(DictBody, RegexOptions.Compiled | RegexOptions.Singleline);
    // Matches every "<objNum> <gen> obj << ... >>" in the document in one pass; FindObjectDict
    // below filters the pre-computed matches by object number instead of building a fresh,
    // uncompiled Regex per call via string interpolation (as the old implementation did), so
    // it's consistent with every other pattern in this class being a precompiled static field.
    private static readonly Regex ObjectDict = new(@"\b(\d+)\s+\d+\s+obj\s*(" + DictBody + ")",
                                                     RegexOptions.Compiled | RegexOptions.Singleline);
    // Matches the identity-scale text matrix "a b c d e f Tm" where a=d=1, b=c=0 - i.e. plain
    // translation, no rotation/scale - tolerating both OfficeIMO's integer form ("1 0 0 1 ...")
    // and a decimal form ("1.000000 0.000000 0.000000 1.000000 ..."). Deliberately does NOT
    // match rotated/scaled matrices (different a/b/c/d), so it stays as selective as before.
    // The leading "(?<![-\d.])" anchors the match so it can't start partway through a larger
    // number - without it, "21 0 0 1 x y Tm" would false-match the substring "1 0 0 1 x y Tm"
    // (starting at the second digit of "21") and report a bogus Y position.
    private static readonly Regex TextMatrix =
        new(@"(?<![-\d.])1(?:\.0+)? 0(?:\.0+)? 0(?:\.0+)? 1(?:\.0+)? [-\d.]+ ([-\d.]+) Tm",
            RegexOptions.Compiled);

    // The PDF simple fonts OfficeIMO declares here use /Encoding /WinAnsiEncoding (WinAnsi is
    // ~= Windows-1252). For byte values 0x20-0x7E and 0xA0-0xFF, WinAnsi and Latin-1 agree, so
    // treating the byte value as the Unicode code point is fine there. They diverge for
    // 0x80-0x9F: Latin-1 maps that range to the C1 control codes, while WinAnsi maps it to
    // typographic characters (smart quotes, dashes, ellipsis, trademark, etc.) that appear
    // constantly in Word-authored text. This table holds the 32 WinAnsi code points for
    // 0x80-0x9F; slots WinAnsi leaves undefined (0x81, 0x8D, 0x8F, 0x90, 0x9D) map to U+FFFD
    // (replacement character) rather than silently producing a wrong character.
    private static readonly char[] WinAnsiHighRange =
    {
        '€', '�', '‚', 'ƒ', '„', '…', '†', '‡', // 80-87
        'ˆ', '‰', 'Š', '‹', 'Œ', '�', 'Ž', '�', // 88-8F
        '�', '‘', '’', '“', '”', '•', '–', '—', // 90-97
        '˜', '™', 'š', '›', 'œ', '�', 'ž', 'Ÿ', // 98-9F
    };

    private static char DecodeWinAnsiByte(int b) =>
        b is >= 0x80 and <= 0x9F ? WinAnsiHighRange[b - 0x80] : (char)b;

    private static string Raw(byte[] pdf) => Encoding.Latin1.GetString(pdf);

    public static bool IsPdf(byte[] pdf) =>
        pdf.Length >= 5 && Encoding.ASCII.GetString(pdf, 0, 5) == "%PDF-";

    /// <summary>
    /// All visible text, in content-stream order.
    ///
    /// TWO ENCODINGS, decided per document. OfficeIMO up to 3.0.3 emitted single-byte WinAnsi hex
    /// strings. From 3.1.0 it embeds a subset TrueType font as a Type0/Identity-H composite, so the
    /// same <c>&lt;...&gt; Tj</c> operator now holds <b>two-byte glyph IDs</b>, not character codes:
    /// <c>&lt;002C0051&gt;</c> is glyph 0x2C then 0x51, which happen to be "In". Decoding those as
    /// WinAnsi bytes yields NUL-separated rubbish and every text assertion fails while the PDF is
    /// perfectly good — which is exactly the silent-looking failure this class exists to prevent.
    ///
    /// Glyph IDs are meaningless without the font, so the mapping back to characters comes from the
    /// document's own <c>/ToUnicode</c> CMap. That is a required part of a well-formed Type0 font
    /// and is what makes the PDF extractable by any conforming reader — its absence would be a real
    /// defect rather than a probe problem.
    /// </summary>
    public static string ExtractText(byte[] pdf)
    {
        var raw = Raw(pdf);

        // /Identity-H is the signal that codes are two-byte glyph IDs. Keyed on that rather than on
        // "a ToUnicode exists", so a simple-encoded PDF that happens to carry one still decodes the
        // old way.
        var glyphs = raw.Contains("/Identity-H", StringComparison.Ordinal)
            ? ToUnicodeMap(raw)
            : null;

        var sb = new StringBuilder();
        foreach (var hex in HexText.Matches(raw).Select(m => m.Groups[1].Value))
        {
            if (glyphs is not null)
            {
                if (hex.Length % 4 != 0) continue;
                for (var i = 0; i < hex.Length; i += 4)
                {
                    var code = Convert.ToInt32(hex.Substring(i, 4), 16);
                    if (glyphs.TryGetValue(code, out var text)) sb.Append(text);
                }
                continue;
            }

            if (hex.Length % 2 != 0) continue;
            for (var i = 0; i < hex.Length; i += 2)
                sb.Append(DecodeWinAnsiByte(Convert.ToInt32(hex.Substring(i, 2), 16)));
        }
        return sb.ToString();
    }

    /// <summary>
    /// Glyph id to text, read from the <c>/ToUnicode</c> CMap's <c>bfchar</c> and <c>bfrange</c>
    /// sections.
    ///
    /// The CMap is emitted uncompressed, so it is parsed straight out of the raw bytes — the same
    /// reason the content streams are readable at all. A destination may be several UTF-16 code
    /// units, because one glyph can stand for a ligature ("fi" from a single glyph), so the value is
    /// a string rather than a char.
    /// </summary>
    private static Dictionary<int, string> ToUnicodeMap(string raw)
    {
        var map = new Dictionary<int, string>();

        foreach (Match section in BfCharSection.Matches(raw))
            foreach (Match pair in BfCharPair.Matches(section.Groups[1].Value))
                map[Convert.ToInt32(pair.Groups[1].Value, 16)] = FromUtf16Hex(pair.Groups[2].Value);

        foreach (Match section in BfRangeSection.Matches(raw))
        {
            foreach (Match entry in BfRangeEntry.Matches(section.Groups[1].Value))
            {
                var lo = Convert.ToInt32(entry.Groups[1].Value, 16);
                var hi = Convert.ToInt32(entry.Groups[2].Value, 16);
                var dst = Convert.ToInt32(entry.Groups[3].Value, 16);

                // Guarded: a malformed range must not hang the suite on a 2^32 loop.
                if (hi < lo || hi - lo > 0xFFFF) continue;
                for (var g = lo; g <= hi; g++) map[g] = char.ConvertFromUtf32(dst + (g - lo));
            }
        }

        return map;
    }

    /// <summary>A CMap destination is UTF-16BE hex, so surrogate pairs arrive as four bytes.</summary>
    private static string FromUtf16Hex(string hex)
    {
        var sb = new StringBuilder();
        for (var i = 0; i + 4 <= hex.Length; i += 4)
            sb.Append((char)Convert.ToInt32(hex.Substring(i, 4), 16));
        return sb.ToString();
    }

    /// <summary>Page count taken from the /Pages tree node.</summary>
    public static int PageCount(byte[] pdf)
    {
        var raw = Raw(pdf);

        // Preferred: resolve the actual page-tree root via trailer -> /Root -> Catalog ->
        // /Pages, so the total comes from the one dictionary that's actually authoritative.
        var viaCatalog = TryPageCountViaCatalog(raw);
        if (viaCatalog is int fromCatalog) return fromCatalog;

        // Fallback (no trailer/xref present, e.g. a hand-built test fixture, or the catalog
        // chain didn't resolve): scan every "<< ... >>" dictionary in the document, keep the
        // ones whose /Type is /Pages, and take the MAXIMUM /Count among them. A page tree's
        // root node's /Count is always >= any intermediate node's /Count, so max is a safe
        // stand-in for "the document total" when we can't walk the tree explicitly.
        var max = 0;
        var found = false;
        foreach (Match dict in AnyDict.Matches(raw).Where(d => TypeIsPages.IsMatch(d.Value)))
        {
            var count = CountField.Match(dict.Value);
            if (!count.Success) continue;
            found = true;
            var value = int.Parse(count.Groups[1].Value, CultureInfo.InvariantCulture);
            if (value > max) max = value;
        }
        return found ? max : 0;
    }

    private static int? TryPageCountViaCatalog(string raw)
    {
        var trailer = TrailerDict.Match(raw);
        if (!trailer.Success) return null;

        var rootRef = RootRef.Match(trailer.Groups[1].Value);
        if (!rootRef.Success) return null;

        var catalogDict = FindObjectDict(raw, rootRef.Groups[1].Value);
        if (catalogDict is null) return null;

        var pagesRef = PagesRef.Match(catalogDict);
        if (!pagesRef.Success) return null;

        var pagesDict = FindObjectDict(raw, pagesRef.Groups[1].Value);
        if (pagesDict is null || !TypeIsPages.IsMatch(pagesDict)) return null;

        var count = CountField.Match(pagesDict);
        return count.Success ? int.Parse(count.Groups[1].Value, CultureInfo.InvariantCulture) : null;
    }

    /// <summary>Finds "&lt;objNum&gt; &lt;gen&gt; obj &lt;&lt; ... &gt;&gt;" and returns the dictionary body.</summary>
    private static string? FindObjectDict(string raw, string objNum)
    {
        foreach (Match m in ObjectDict.Matches(raw).Where(m => m.Groups[1].Value == objNum))
            return m.Groups[2].Value;
        return null;
    }

    /// <summary>Y coordinate of every text-positioning operator. Negative values are drawn off-page.</summary>
    public static IReadOnlyList<double> TextYPositions(byte[] pdf) =>
        TextMatrix.Matches(Raw(pdf))
                  .Select(m => double.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture))
                  .ToList();

    // One /MediaBox [x0 y0 x1 y1] per page object. A regex rather than a real PDF parse for the
    // same reason the rest of this file is one: OfficeIMO writes uncompressed, so the array is
    // present in the bytes as text, and a general parser here would be more code for no more
    // certainty.
    private static readonly Regex MediaBox = new(
        @"/MediaBox\s*\[\s*(-?[\d.]+)\s+(-?[\d.]+)\s+(-?[\d.]+)\s+(-?[\d.]+)\s*\]",
        RegexOptions.Compiled);

    /// <summary>
    /// Every page's <c>/MediaBox</c>, as (width, height) in points, in document order.
    ///
    /// This is the only assertion that proves page setup survives the DOCX → PDF render. Every
    /// other page-setup test reads the .docx, which OfficeIMO could stop honouring without any of
    /// them noticing.
    /// </summary>
    public static IReadOnlyList<(double Width, double Height)> MediaBoxes(byte[] pdf)
    {
        ArgumentNullException.ThrowIfNull(pdf);

        static double N(Group g) => double.Parse(g.Value, CultureInfo.InvariantCulture);

        return MediaBox.Matches(Raw(pdf))
                       .Select(m => (Width: N(m.Groups[3]) - N(m.Groups[1]),
                                     Height: N(m.Groups[4]) - N(m.Groups[2])))
                       .ToList();
    }

    private static readonly Regex Rotate = new(@"/Rotate\s+(-?\d+)", RegexOptions.Compiled);

    /// <summary>
    /// Every <b>explicit</b> <c>/Rotate</c> value in the file, in the order they appear.
    ///
    /// Deliberately NOT one entry per page, unlike <see cref="MediaBoxes"/>. Every page carries a
    /// <c>/MediaBox</c>, but <c>/Rotate</c> is optional and defaults to 0, so an unrotated page has
    /// no entry at all. Reporting a padded zero per page would be a guess dressed as a measurement —
    /// this returns what the file actually says and lets the test assert on that.
    ///
    /// Which makes the count load-bearing: rotating one page of three must produce exactly one
    /// entry, and that is what distinguishes a working rotation from one that silently touched
    /// everything.
    /// </summary>
    public static IReadOnlyList<int> PageRotations(byte[] pdf)
    {
        ArgumentNullException.ThrowIfNull(pdf);

        return Rotate.Matches(Raw(pdf))
                     .Select(m => int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture))
                     .ToList();
    }
}
