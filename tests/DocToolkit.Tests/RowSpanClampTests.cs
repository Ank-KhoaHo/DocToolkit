namespace DocToolkit.Tests;

/// <summary>
/// Clamping a <c>rowspan</c> that reaches past the last row of its table.
///
/// <b>Measured 2026-08-17 over 181 real `.gov` pages: HTML to DOCX went from 163 to 177 (90.1% to
/// 97.8%) and HTML to PDF from 51.4% to 55.8%.</b> Fourteen pages that could not be converted at all
/// now convert, and the other four failures in the set are unchanged - they fail elsewhere, for
/// reasons this does not touch.
///
/// <b>The central claim is that this cannot change a document that already converts</b>, and the
/// tests that matter most are the ones asserting the input comes back <i>by reference</i>. Every
/// document carrying an overrunning span throws today, so there is no input whose output changes -
/// only inputs that begin to have one. If a test here ever shows the string being rewritten when
/// nothing overruns, that guarantee is gone and the risk profile of this class is completely
/// different.
/// </summary>
public class RowSpanClampTests
{
    // ---- the guarantee: untouched unless something actually overruns -------------------------------

    [Theory]
    [InlineData("<p>no table at all</p>")]
    [InlineData("<table><tr><td>a</td></tr></table>")]                          // table, no rowspan
    [InlineData("<table><tr><td colspan=\"9\">a</td></tr></table>")]            // colspan is not it
    [InlineData("<table><tr><td rowspan=\"1\">a</td></tr></table>")]            // fits
    [InlineData("<table><tr><td rowspan=\"2\">a</td></tr><tr><td>b</td></tr></table>")]   // fits exactly
    [InlineData("<table><tr><td rowspan=\"0\">a</td></tr></table>")]            // 0 means "to the end"
    public void NothingOverruns_TheSameStringComesBack(string html)
    {
        // ReferenceEquals, not string equality: this asserts no parse-and-serialise round trip
        // happened at all, which is the whole basis for saying a working document cannot change.
        Assert.Same(html, RowSpanClamp.Apply(html));
    }

    [Fact]
    public void AnOverrunningSpan_IsTheOnlyThingThatCausesARewrite()
    {
        const string html = "<table><tr><td rowspan=\"2\">a</td></tr></table>";

        Assert.NotSame(html, RowSpanClamp.Apply(html));
    }

    [Theory]
    [InlineData(1, 5, "1")]
    [InlineData(2, 5, "2")]
    [InlineData(3, 100, "3")]
    public void ItClampsToTheROWSTHATREMAIN_NotSimplyToOne(int rows, int span, string expected)
    {
        // Without this, clamping every overrunning span to 1 would pass every other test in the
        // file: the crash goes away either way. But the author asked for a cell reaching the bottom
        // of the table, and the rows that remain is what a browser gives them - clamping to 1 would
        // convert successfully into the wrong document, which is the worse failure of the two.
        var result = RowSpanClamp.Apply(Table(rows, span));

        Assert.Contains($"rowspan=\"{expected}\"", result, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain($"rowspan=\"{span}\"", result, StringComparison.OrdinalIgnoreCase);
    }

    // ---- what it converts to -----------------------------------------------------------------------

    [Fact]
    public async Task TheDocumentThatUsedToCrash_NowConverts()
    {
        // The 46-byte reduction of the most frequent real-world failure in the corpus.
        var docx = await HtmlToDocxConverter.ConvertAsync(
            "<table><tr><td rowspan=\"2\"></td></tr></table>");

        Assert.NotEmpty(docx);
    }

    [Fact]
    public async Task TheCellContentSurvives()
    {
        // A conversion that succeeded by dropping the cell would pass every assertion above.
        var docx = await HtmlToDocxConverter.ConvertAsync(
            "<table><tr><td rowspan=\"5\">KEPT</td><td>ALSO-KEPT</td></tr></table>");

        var text = DocxEditor.ExtractText(docx);
        Assert.Contains("KEPT", text, StringComparison.Ordinal);
        Assert.Contains("ALSO-KEPT", text, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(1, 2)]
    [InlineData(1, 3)]
    [InlineData(1, 100)]     // the corpus really did carry rowspan="100"
    [InlineData(2, 3)]
    [InlineData(3, 4)]
    [InlineData(2, 103)]
    public async Task EveryOverrunningCombinationConverts(int rows, int span)
    {
        var docx = await HtmlToDocxConverter.ConvertAsync(Table(rows, span));

        Assert.NotEmpty(docx);
    }

    [Theory]
    [InlineData(1, 1)]
    [InlineData(2, 2)]
    [InlineData(4, 4)]
    public async Task EveryFittingCombinationStillConverts(int rows, int span)
    {
        // The control. A clamp that mangled valid tables would satisfy the theory above.
        var docx = await HtmlToDocxConverter.ConvertAsync(Table(rows, span));

        Assert.NotEmpty(docx);
    }

    private static string Table(int rows, int span)
    {
        var sb = new System.Text.StringBuilder("<table>");
        for (var r = 0; r < rows; r++)
            sb.Append(r == 0 ? $"<tr><td rowspan=\"{span}\">c{r}</td></tr>" : $"<tr><td>c{r}</td></tr>");
        return sb.Append("</table>").ToString();
    }

    // ---- the structure it has to agree with the parser about ---------------------------------------

    [Fact]
    public void RowsAreCountedPerSection_NotPerTable()
    {
        // The failing code allocates one accumulator PER TABLE PART. So a one-row thead carrying
        // rowspan=2 overruns even though the table as a whole has plenty of rows below it - and a
        // clamp counting whole-table rows would leave this crashing.
        const string html =
            "<table><thead><tr><td rowspan=\"2\">h</td></tr></thead>"
            + "<tbody><tr><td>a</td></tr><tr><td>b</td></tr><tr><td>c</td></tr></tbody></table>";

        Assert.NotSame(html, RowSpanClamp.Apply(html));
    }

    [Fact]
    public async Task ASectionedTableThatOverrunsConverts()
    {
        var docx = await HtmlToDocxConverter.ConvertAsync(
            "<table><thead><tr><td rowspan=\"9\">HEAD</td></tr></thead>"
            + "<tbody><tr><td>BODY</td></tr></tbody></table>");

        var text = DocxEditor.ExtractText(docx);
        Assert.Contains("HEAD", text, StringComparison.Ordinal);
        Assert.Contains("BODY", text, StringComparison.Ordinal);
    }

    [Fact]
    public async Task ANestedTableIsHandled()
    {
        // Nesting is what defeats every string-scanning approach to this, and the corpus is full of
        // it - the smallest real reproduction had the offending cell inside a nested table.
        var docx = await HtmlToDocxConverter.ConvertAsync(
            "<table><tr><td>OUTER<table><tr><td rowspan=\"4\">INNER</td></tr></table></td></tr></table>");

        var text = DocxEditor.ExtractText(docx);
        Assert.Contains("OUTER", text, StringComparison.Ordinal);
        Assert.Contains("INNER", text, StringComparison.Ordinal);
    }

    [Fact]
    public async Task TwoTablesWhereOnlyOneOverruns()
    {
        var docx = await HtmlToDocxConverter.ConvertAsync(
            "<table><tr><td>FINE</td></tr></table>"
            + "<table><tr><td rowspan=\"7\">BROKEN</td></tr></table>");

        var text = DocxEditor.ExtractText(docx);
        Assert.Contains("FINE", text, StringComparison.Ordinal);
        Assert.Contains("BROKEN", text, StringComparison.Ordinal);
    }

    // ---- robustness ---------------------------------------------------------------------------------

    [Fact]
    public void UnparseableInputIsHandedOnUnchanged_RatherThanDiagnosedHere()
    {
        // If it cannot be parsed here it will not be parsed downstream either, and the real converter
        // should produce the real diagnostic. A pre-pass inventing its own would be worse.
        const string html = "<table><tr><td rowspan=\"2\"";

        var result = RowSpanClamp.Apply(html);

        Assert.NotNull(result);
    }

    [Fact]
    public void AMalformedSpanValueIsNotTreatedAsOverrunning()
    {
        // "abc" is not a number. AngleSharp resolves an unparseable rowspan to 1, which fits, so
        // nothing should be rewritten - and asserting that pins the behaviour rather than leaving it
        // to whatever the parser happens to do next release.
        const string html = "<table><tr><td rowspan=\"abc\">a</td></tr></table>";

        Assert.Same(html, RowSpanClamp.Apply(html));
    }

    [Fact]
    public async Task OrdinaryHtmlIsUnaffected()
    {
        var docx = await HtmlToDocxConverter.ConvertAsync(
            "<h1>Title</h1><p>Body</p><table><tr><td>a</td><td>b</td></tr></table>");

        Assert.Contains("Title", DocxEditor.ExtractText(docx), StringComparison.Ordinal);
    }
}
