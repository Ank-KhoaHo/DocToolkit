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

    /// <summary>
    /// A field whose special character is at <b>index 0</b>.
    ///
    /// <c>CsvQuotesOnlyFieldsThatNeedIt</c> above never puts one there — its commas and quotes are
    /// all mid-field — so it passes identically against <c>IndexOfAny(...) &gt; 0</c>, which fails
    /// to quote exactly these fields and emits a CSV with a row that silently gained a column.
    /// Found by mutation (B14): that <c>&gt;= 0 → &gt; 0</c> mutant survived the whole suite.
    /// </summary>
    [Fact]
    public void CsvQuotesAFieldWhoseSpecialCharacterIsFirst()
    {
        var xlsx = WorkbookEditor.Create("S", new[]
        {
            new object?[] { ",leading comma", "\"leading quote", "\nleading newline" },
        });

        var csv = XlsxToCsvConverter.Convert(xlsx, "S");

        Assert.Equal("\",leading comma\",\"\"\"leading quote\",\"\nleading newline\"\r\n", csv);
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

    /// <summary>
    /// A sheet holding <b>exactly one</b> row is a header and no body, so there must be no
    /// <c>&lt;tbody&gt;</c> at all — not an unopened closing tag.
    ///
    /// The two neighbouring cases do not catch this. The three-row workbook has a body, and the
    /// empty sheet has zero rows, so <c>rows.Count &gt; 1 → &gt;= 1</c> is false in both. One row
    /// is the only count that separates them, and mutation (B14) found that mutant surviving.
    /// </summary>
    [Fact]
    public void ASingleRowSheetEmitsNoTbodyAtAll()
    {
        var xlsx = WorkbookEditor.Create("One", new[]
        {
            new object?[] { "Region", "Total" },
        });

        var html = XlsxToHtmlConverter.Convert(xlsx, "One");

        Assert.Equal(
            "<table>\n" +
            "  <thead>\n" +
            "    <tr><th>Region</th><th>Total</th></tr>\n" +
            "  </thead>\n" +
            "</table>",
            html);

        // Stated separately from the literal above, because this is the property that matters:
        // a closing tag with no opening one is malformed markup every parser handles differently.
        Assert.DoesNotContain("</tbody>", html, StringComparison.Ordinal);
    }

    /// <summary>
    /// The <c>Stream</c> overloads refuse a blank sheet name and a cancelled token
    /// <b>before reading a single byte</b>.
    ///
    /// <b>Asserting the exception alone proves nothing here, and that is the whole point of this
    /// test.</b> <c>ConvertAsync</c> drains the source and then calls the <c>byte[]</c> overload,
    /// which carries the same <c>ThrowIfNullOrWhiteSpace</c> — and <c>DrainAsync</c> observes the
    /// token itself. So deleting either guard at the top leaves the identical
    /// <c>ArgumentException</c> and <c>OperationCanceledException</c> arriving from one layer
    /// down, after the entire stream has been read. Measured: an exception-only version of this
    /// test passed against all four mutants.
    ///
    /// The read count is what separates them, and it is a real contract rather than a trick to
    /// kill a mutant: a caller handing over a 200 MB upload and a mistyped sheet name should not
    /// pay for the transfer first. Same failure shape as the seven <c>PdfEditor</c> overloads that
    /// passed the cancellation suite only because <c>destination.WriteAsync</c> refused at the
    /// end.
    /// </summary>
    [Fact]
    public async Task StreamOverloadsRefuseABlankSheetNameAndACancelledTokenBeforeReading()
    {
        foreach (var convert in new Func<Stream, string, CancellationToken, Task>[]
        {
            (s, n, ct) => XlsxToCsvConverter.ConvertAsync(s, n, ct),
            (s, n, ct) => XlsxToHtmlConverter.ConvertAsync(s, n, ct),
        })
        {
            using var inner = new MemoryStream(Workbook());
            using var source = new TrackingStream(inner);

            var blank = await Assert.ThrowsAsync<ArgumentException>(
                () => convert(source, "  ", CancellationToken.None));
            Assert.Equal("sheetName", blank.ParamName);
            Assert.Equal(0, source.SyncReads + source.AsyncReads);

            using var cancelled = new CancellationTokenSource();
            await cancelled.CancelAsync();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(
                () => convert(source, "Data", cancelled.Token));
            Assert.Equal(0, source.SyncReads + source.AsyncReads);
        }
    }
}
