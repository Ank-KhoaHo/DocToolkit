using System.Globalization;
using Xunit;

namespace DocToolkit.Tests;

/// <summary>
/// XLSX → CSV and XLSX → HTML (A21).
///
/// The culture tests are the reason this class exists. `ReadSheet` renders through
/// <see cref="CultureInfo.CurrentCulture"/>, so a workbook reads back as <c>1234,5</c> on a German
/// machine — a decimal comma inside a comma-delimited file. Measured across en-US, de-DE and
/// fr-FR before either exporter was written.
/// </summary>
public class XlsxExportTests
{
    private static byte[] Workbook() => WorkbookEditor.Create("Data", new[]
    {
        new object?[] { "Region", "Total", "When" },
        new object?[] { "North", 1234.5, new DateTime(2026, 8, 13) },
        new object?[] { "South", 980, new DateTime(2026, 8, 14, 9, 30, 0) },
    });

    /// <summary>Runs <paramref name="action"/> with the thread pinned to <paramref name="culture"/>.</summary>
    private static void InCulture(string culture, Action action)
    {
        var prior = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = new CultureInfo(culture);
            action();
        }
        finally
        {
            CultureInfo.CurrentCulture = prior;
        }
    }

    // =====================================================================================
    // CSV
    // =====================================================================================

    [Fact]
    public void CsvHoldsTheLiteralGrid()
    {
        var csv = XlsxToCsvConverter.Convert(Workbook(), "Data");

        Assert.Equal(
            "Region,Total,When\r\n" +
            "North,1234.5,2026-08-13\r\n" +
            "South,980,2026-08-14 09:30:00\r\n",
            csv);
    }

    /// <summary>
    /// <b>The test this class exists for.</b> The same workbook produces byte-identical CSV on
    /// three cultures whose number and date formats all differ. Without this, a German machine
    /// emits <c>1234,5</c> and every downstream reader sees an extra column.
    /// </summary>
    [Fact]
    public void CsvIsIdenticalAcrossCultures()
    {
        var xlsx = Workbook();
        var expected = XlsxToCsvConverter.Convert(xlsx, "Data");

        foreach (var culture in new[] { "en-US", "de-DE", "fr-FR" })
        {
            InCulture(culture, () =>
                Assert.Equal(expected, XlsxToCsvConverter.Convert(xlsx, "Data")));
        }

        // ...and the anchor: identical-but-wrong would satisfy the loop above on its own.
        Assert.Contains("1234.5", expected, StringComparison.Ordinal);
        Assert.DoesNotContain("1234,5", expected, StringComparison.Ordinal);
    }

    /// <summary>
    /// It is the EXPORTER that is invariant, not the whole library: `ReadSheet` still follows the
    /// caller's culture, which its own documentation promises. Pinning both sides here means a
    /// future change cannot quietly make them agree by breaking the documented one.
    /// </summary>
    [Fact]
    public void ReadSheetStillFollowsTheCallersCulture()
    {
        var xlsx = Workbook();

        InCulture("de-DE", () =>
        {
            Assert.Equal("1234,5", WorkbookEditor.ReadSheet(xlsx, "Data")[1][1]);
            Assert.Contains("1234.5", XlsxToCsvConverter.Convert(xlsx, "Data"), StringComparison.Ordinal);
        });
    }

    [Fact]
    public void CsvQuotesOnlyFieldsThatNeedIt_AndDoublesEmbeddedQuotes()
    {
        var xlsx = WorkbookEditor.Create("S", new[]
        {
            new object?[] { "plain", "has,comma", "has\"quote", "has\nnewline" },
        });

        var csv = XlsxToCsvConverter.Convert(xlsx, "S");

        Assert.Equal("plain,\"has,comma\",\"has\"\"quote\",\"has\nnewline\"\r\n", csv);
    }

    [Fact]
    public void AFormulaExportsItsComputedValue()
    {
        var xlsx = WorkbookEditor.Create("S", new[]
        {
            new object?[] { 2, 3, XlsxFormula.From("=A1+B1") },
        });

        Assert.Equal("2,3,5\r\n", XlsxToCsvConverter.Convert(xlsx, "S"));
    }

    // =====================================================================================
    // HTML
    // =====================================================================================

    [Fact]
    public void HtmlIsATableFragmentWithAHeaderRow()
    {
        var html = XlsxToHtmlConverter.Convert(Workbook(), "Data");

        Assert.Equal(
            "<table>\n" +
            "  <thead>\n" +
            "    <tr><th>Region</th><th>Total</th><th>When</th></tr>\n" +
            "  </thead>\n" +
            "  <tbody>\n" +
            "    <tr><td>North</td><td>1234.5</td><td>2026-08-13</td></tr>\n" +
            "    <tr><td>South</td><td>980</td><td>2026-08-14 09:30:00</td></tr>\n" +
            "  </tbody>\n" +
            "</table>",
            html);

        // A fragment, deliberately unlike DocxToHtmlConverter.
        Assert.DoesNotContain("<html", html, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// A workbook is untrusted input. A cell holding markup must arrive as text.
    /// </summary>
    [Fact]
    public void HtmlEscapesCellContent()
    {
        var xlsx = WorkbookEditor.Create("S", new[]
        {
            new object?[] { "<script>alert('x')</script>", "a & b", "\"q\"" },
        });

        var html = XlsxToHtmlConverter.Convert(xlsx, "S");

        Assert.DoesNotContain("<script>", html, StringComparison.Ordinal);
        Assert.Contains("&lt;script&gt;alert(&#39;x&#39;)&lt;/script&gt;", html, StringComparison.Ordinal);
        Assert.Contains("a &amp; b", html, StringComparison.Ordinal);
        Assert.Contains("&quot;q&quot;", html, StringComparison.Ordinal);
    }

    [Fact]
    public void TheTwoExportersAgreeOnWhatACellSays()
    {
        var xlsx = Workbook();

        var csv = XlsxToCsvConverter.Convert(xlsx, "Data");
        var html = XlsxToHtmlConverter.Convert(xlsx, "Data");

        // Both render the same invariant text; neither can drift because both read one grid.
        foreach (var literal in new[] { "1234.5", "2026-08-13", "2026-08-14 09:30:00", "980" })
        {
            Assert.Contains(literal, csv, StringComparison.Ordinal);
            Assert.Contains(literal, html, StringComparison.Ordinal);
        }
    }

    // =====================================================================================
    // Guards and Stream overloads
    // =====================================================================================

    [Fact]
    public async Task StreamOverloadsMatchTheByteArrayForm_AndLeaveTheSourceOpen()
    {
        var xlsx = Workbook();

        using var forCsv = new MemoryStream(xlsx, writable: false);
        Assert.Equal(
            XlsxToCsvConverter.Convert(xlsx, "Data"),
            await XlsxToCsvConverter.ConvertAsync(forCsv, "Data"));
        Assert.True(forCsv.CanRead, "ConvertAsync disposed a source it does not own");

        using var forHtml = new MemoryStream(xlsx, writable: false);
        Assert.Equal(
            XlsxToHtmlConverter.Convert(xlsx, "Data"),
            await XlsxToHtmlConverter.ConvertAsync(forHtml, "Data"));
    }

    [Fact]
    public async Task RejectsBadInput()
    {
        Assert.Throws<ArgumentNullException>(() => XlsxToCsvConverter.Convert(null!, "Data"));
        Assert.Throws<ArgumentNullException>(() => XlsxToHtmlConverter.Convert(null!, "Data"));

        Assert.Throws<ArgumentException>(() => XlsxToCsvConverter.Convert(Array.Empty<byte>(), "Data"));
        Assert.Throws<ArgumentException>(() => XlsxToCsvConverter.Convert(Workbook(), " "));

        // A sheet that is not there is a conversion failure, not an argument error - the caller's
        // arguments were well formed; the workbook simply does not hold that sheet.
        var missing = Assert.Throws<DocumentConversionException>(
            () => XlsxToCsvConverter.Convert(Workbook(), "Nope"));
        Assert.Contains("Nope", missing.Message, StringComparison.Ordinal);

        using var empty = new MemoryStream();
        var ex = await Assert.ThrowsAsync<ArgumentException>(
            () => XlsxToCsvConverter.ConvertAsync(empty, "Data"));
        Assert.Equal("source", ex.ParamName);
    }

    [Fact]
    public void AnEmptySheetProducesEmptyOutput()
    {
        var xlsx = WorkbookEditor.Create("Blank", Array.Empty<object?[]>());

        Assert.Equal(string.Empty, XlsxToCsvConverter.Convert(xlsx, "Blank"));
        Assert.Equal("<table>\n</table>", XlsxToHtmlConverter.Convert(xlsx, "Blank"));
    }
}
