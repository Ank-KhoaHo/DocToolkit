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
    private static readonly Regex PageTree = new(@"/Type\s*/Pages.*?/Count\s+(\d+)",
                                                 RegexOptions.Compiled | RegexOptions.Singleline);
    private static readonly Regex TextMatrix = new(@"1 0 0 1 [-\d.]+ ([-\d.]+) Tm", RegexOptions.Compiled);

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
                sb.Append((char)Convert.ToInt32(hex.Substring(i, 2), 16));
        }
        return sb.ToString();
    }

    /// <summary>Page count taken from the /Pages tree node.</summary>
    public static int PageCount(byte[] pdf)
    {
        var m = PageTree.Match(Raw(pdf));
        return m.Success ? int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture) : 0;
    }

    /// <summary>Y coordinate of every text-positioning operator. Negative values are drawn off-page.</summary>
    public static IReadOnlyList<double> TextYPositions(byte[] pdf) =>
        TextMatrix.Matches(Raw(pdf))
                  .Select(m => double.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture))
                  .ToList();
}
