using System.Globalization;
using System.Text;
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

    // A dictionary that may contain up to one level of nested "<< ... >>" (e.g. a Catalog's
    // /Names dictionary, or a Page's /Resources dictionary). This is not a full PDF parser -
    // it doesn't handle arbitrarily deep nesting - but it's enough to keep a match from
    // spilling out of the dictionary it started in and into an unrelated neighboring object,
    // which a plain non-greedy ".*?" (as the old implementation used) does not guard against.
    private const string DictBody = @"<<(?:[^<>]|<<[^<>]*>>)*>>";
    private static readonly Regex TrailerDict = new(@"trailer\s*(" + DictBody + ")",
                                                     RegexOptions.Compiled | RegexOptions.Singleline);
    private static readonly Regex RootRef = new(@"/Root\s+(\d+)\s+\d+\s+R", RegexOptions.Compiled);
    private static readonly Regex PagesRef = new(@"/Pages\s+(\d+)\s+\d+\s+R", RegexOptions.Compiled);
    private static readonly Regex TypeIsPages = new(@"/Type\s*/Pages\b", RegexOptions.Compiled);
    private static readonly Regex CountField = new(@"/Count\s+(\d+)", RegexOptions.Compiled);
    private static readonly Regex AnyDict = new(DictBody, RegexOptions.Compiled | RegexOptions.Singleline);
    // Matches the identity-scale text matrix "a b c d e f Tm" where a=d=1, b=c=0 - i.e. plain
    // translation, no rotation/scale - tolerating both OfficeIMO's integer form ("1 0 0 1 ...")
    // and a decimal form ("1.000000 0.000000 0.000000 1.000000 ..."). Deliberately does NOT
    // match rotated/scaled matrices (different a/b/c/d), so it stays as selective as before.
    private static readonly Regex TextMatrix =
        new(@"1(?:\.0+)? 0(?:\.0+)? 0(?:\.0+)? 1(?:\.0+)? [-\d.]+ ([-\d.]+) Tm", RegexOptions.Compiled);

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

    /// <summary>All visible text, in content-stream order.</summary>
    public static string ExtractText(byte[] pdf)
    {
        var sb = new StringBuilder();
        foreach (Match m in HexText.Matches(Raw(pdf)))
        {
            var hex = m.Groups[1].Value;
            if (hex.Length % 2 != 0) continue;
            for (var i = 0; i < hex.Length; i += 2)
                sb.Append(DecodeWinAnsiByte(Convert.ToInt32(hex.Substring(i, 2), 16)));
        }
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
        foreach (Match dict in AnyDict.Matches(raw))
        {
            if (!TypeIsPages.IsMatch(dict.Value)) continue;
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
        var m = Regex.Match(raw, @"\b" + objNum + @"\s+\d+\s+obj\s*(" + DictBody + ")",
                             RegexOptions.Singleline);
        return m.Success ? m.Groups[1].Value : null;
    }

    /// <summary>Y coordinate of every text-positioning operator. Negative values are drawn off-page.</summary>
    public static IReadOnlyList<double> TextYPositions(byte[] pdf) =>
        TextMatrix.Matches(Raw(pdf))
                  .Select(m => double.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture))
                  .ToList();
}
