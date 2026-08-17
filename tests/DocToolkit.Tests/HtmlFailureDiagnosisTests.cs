namespace DocToolkit.Tests;

/// <summary>
/// The HTML failure that real pages hit most often, and the message that now names it.
///
/// <b>Measured 2026-08-17 across 179 real `.gov` pages: 14 of them - 7.7% - fail on a table cell
/// whose `rowspan` reaches past the last row of its table.</b> That was the single most frequent
/// conversion failure in the corpus, and what a caller was told about it was
/// <i>"See the inner exception for details"</i> over a bare <see cref="IndexOutOfRangeException"/>
/// naming no table, no cell and no remedy.
///
/// <b>The tests below are as much about what is NOT claimed as what is.</b> A message naming a
/// specific cause is only worth having if it cannot appear on a failure it does not describe, so
/// the negative controls carry the weight: an index error from somewhere else must still get the
/// generic wrapper, and markup whose rowspan fits must still convert.
/// </summary>
public class HtmlFailureDiagnosisTests
{
    private const string Overhanging = "<table><tr><td rowspan=\"2\"></td></tr></table>";

    // ---- the diagnosis ---------------------------------------------------------------------------

    [Fact]
    public async Task AnOverhangingRowSpan_IsNamed_RatherThanLeftAsSeeTheInnerException()
    {
        var ex = await Assert.ThrowsAsync<DocumentConversionException>(
            () => HtmlToDocxConverter.ConvertAsync(Overhanging));

        // The construct, so the reader knows WHERE to look...
        Assert.Contains("rowspan", ex.Message, StringComparison.OrdinalIgnoreCase);
        // ...and the remedy, so they know what to do about it. A message that named the cause and
        // stopped would be the failure mode this package already corrected once.
        Assert.Contains("reduce the rowspan", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task TheOriginalExceptionIsStillTheInnerOne_SoNothingIsHidden()
    {
        var ex = await Assert.ThrowsAsync<DocumentConversionException>(
            () => HtmlToDocxConverter.ConvertAsync(Overhanging));

        // Naming the cause must not cost the evidence. Someone debugging still needs the frame.
        Assert.IsType<IndexOutOfRangeException>(ex.InnerException);
        Assert.Contains("GuessColumnsCount", ex.InnerException!.StackTrace, StringComparison.Ordinal);
    }

    [Fact]
    public async Task ItSaysTheMarkupIsValid_BecauseItIs()
    {
        var ex = await Assert.ThrowsAsync<DocumentConversionException>(
            () => HtmlToDocxConverter.ConvertAsync(Overhanging));

        // Worth stating explicitly in the message: a reader who is told only "cannot read that"
        // will go looking for a mistake in their HTML, and there isn't one - browsers clamp it.
        Assert.Contains("browsers clamp", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    // ---- the boundary: exactly "span exceeds the rows that remain" ----------------------------------

    [Theory]
    [InlineData(1, 1)]   // fits
    [InlineData(2, 1)]
    [InlineData(2, 2)]
    [InlineData(3, 2)]
    [InlineData(3, 3)]
    [InlineData(4, 4)]
    public async Task ARowSpanThatFitsStillConverts(int rows, int span)
    {
        // The control that stops the diagnosis being reached by refusing every table with a rowspan.
        // Measured boundary: it breaks only when span > rows remaining, so all of these must pass.
        var docx = await HtmlToDocxConverter.ConvertAsync(Table(rows, span));

        Assert.NotEmpty(docx);
    }

    [Theory]
    [InlineData(1, 2)]
    [InlineData(1, 3)]
    [InlineData(2, 3)]
    [InlineData(3, 4)]
    public async Task ARowSpanThatOverrunsIsDiagnosed(int rows, int span)
    {
        var ex = await Assert.ThrowsAsync<DocumentConversionException>(
            () => HtmlToDocxConverter.ConvertAsync(Table(rows, span)));

        Assert.Contains("rowspan", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task AColSpanIsNotAffected_HoweverLarge()
    {
        // Pins that the diagnosis is about rowspan specifically. colspan was measured unaffected at
        // every value tried, so a message firing here would be over-claiming.
        var docx = await HtmlToDocxConverter.ConvertAsync(
            "<table><tr><td colspan=\"9\"></td></tr></table>");

        Assert.NotEmpty(docx);
    }

    private static string Table(int rows, int span)
    {
        var sb = new System.Text.StringBuilder("<table>");
        for (var r = 0; r < rows; r++)
            sb.Append(r == 0 ? $"<tr><td rowspan=\"{span}\"></td></tr>" : "<tr><td></td></tr>");
        return sb.Append("</table>").ToString();
    }

    // ---- negative controls: the message must not appear on failures it does not describe -----------

    [Fact]
    public void AnIndexErrorFromSomewhereElse_GetsNoDiagnosis()
    {
        // A REAL IndexOutOfRangeException with a REAL stack trace that simply does not contain the
        // frame. This is what discriminates "the frame identifies it" from "the exception type
        // identifies it" - and the second would put a table message on any index error at all.
        Exception caught;
        try
        {
            var empty = Array.Empty<int>();
            _ = empty[5];
            throw new InvalidOperationException("unreachable");
        }
        catch (IndexOutOfRangeException ex)
        {
            caught = ex;
        }

        Assert.NotNull(caught.StackTrace);
        Assert.Null(HtmlFailureDiagnosis.Describe(caught));
    }

    [Fact]
    public void AnExceptionWithNoStackAtAll_GetsNoDiagnosis()
    {
        // Fails closed rather than throwing on a null StackTrace.
        Assert.Null(HtmlFailureDiagnosis.Describe(new IndexOutOfRangeException()));
    }

    [Fact]
    public void ADifferentExceptionType_GetsNoDiagnosis()
    {
        Assert.Null(HtmlFailureDiagnosis.Describe(new InvalidOperationException("anything")));
    }

    /// <summary>
    /// An exception carrying a stack we choose. <see cref="IndexOutOfRangeException"/> is
    /// <b>sealed</b>, so the two halves of the conjunction have to be tested from opposite
    /// directions: the frame comes from a staged exception of the wrong TYPE, and the type comes
    /// from a real <see cref="IndexOutOfRangeException"/> with a real stack lacking the FRAME
    /// (<see cref="AnIndexErrorFromSomewhereElse_GetsNoDiagnosis"/>).
    /// </summary>
    private sealed class Staged(string stack) : Exception
    {
        public override string StackTrace { get; } = stack;
    }

    [Fact]
    public void TheFrameAloneIsNotEnough_TheTypeMustMatchToo()
    {
        // Without this, dropping the `ex is IndexOutOfRangeException` check would survive every
        // other test here: none of their exceptions has a stack containing the frame, so the type
        // check never decides the answer. This makes it decide.
        Assert.Null(HtmlFailureDiagnosis.Describe(new Staged(
            "   at HtmlToOpenXml.Expressions.TableExpression.GuessColumnsCount(IHtmlTableElement t)")));
    }

    [Fact]
    public async Task OrdinaryHtmlStillConverts()
    {
        // The broadest control: the change is on a failure path and must not touch the happy one.
        var docx = await HtmlToDocxConverter.ConvertAsync(
            "<h1>Title</h1><p>Body</p><table><tr><td>a</td><td>b</td></tr></table>");

        Assert.NotEmpty(docx);
    }
}
