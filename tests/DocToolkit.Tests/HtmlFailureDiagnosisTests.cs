using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlToOpenXml;

namespace DocToolkit.Tests;

/// <summary>
/// The message that names an overrunning <c>rowspan</c>, and — more importantly — the failures it
/// must NOT appear on.
///
/// <b>This is a safety net now, not the primary answer.</b> <see cref="RowSpanClamp"/> repairs the
/// markup before the parser sees it, so no ordinary document reaches this message any more. It is
/// kept because the clamp is a heuristic over a parser's view of a document and can be wrong in ways
/// that are hard to enumerate: input AngleSharp cannot parse is handed on untouched, and a table
/// shape where its section view disagrees with the converter's would slip through. If either
/// happens, the caller gets a message naming the construct rather than a bare
/// <see cref="IndexOutOfRangeException"/>.
///
/// <b>The underlying parser defect is unchanged</b> — the clamp stops callers reaching it, it does
/// not fix it. So the genuine exception is still obtainable by driving the parser directly, which is
/// how the message below is tested rather than being asserted against a copy of itself.
/// </summary>
public class HtmlFailureDiagnosisTests
{
    private const string Overrunning = "<table><tr><td rowspan=\"2\"></td></tr></table>";

    /// <summary>
    /// The real exception, from the real parser, bypassing the clamp.
    /// </summary>
    /// <remarks>
    /// <b>This test goes red on its own if the parser is ever fixed upstream</b>, which is the
    /// intent: the day <c>ParseBody</c> stops throwing here, the diagnosis has nothing left to
    /// describe and both it and <see cref="RowSpanClamp"/> should be reconsidered. A test that
    /// quietly passed in that world would hide the news.
    /// </remarks>
    private static Exception RealParserFailure()
    {
        try
        {
            using var stream = new MemoryStream();
            using var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(new Body());
            new HtmlConverter(main).ParseBody(Overrunning).GetAwaiter().GetResult();
        }
        catch (IndexOutOfRangeException ex)
        {
            return ex;
        }

        throw new InvalidOperationException(
            "The parser no longer throws on an overrunning rowspan. If it was fixed upstream, "
            + "RowSpanClamp and HtmlFailureDiagnosis both have nothing left to do.");
    }

    // ---- what the message says --------------------------------------------------------------------

    [Fact]
    public void ItNamesTheConstructAndTheRemedy()
    {
        var message = HtmlFailureDiagnosis.Describe(RealParserFailure());

        Assert.NotNull(message);
        Assert.Contains("rowspan", message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("reduce the rowspan", message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ItSaysTheMarkupIsValid_BecauseItIs()
    {
        // A reader told only "cannot read that" goes looking for a mistake they did not make.
        var message = HtmlFailureDiagnosis.Describe(RealParserFailure());

        Assert.Contains("browsers clamp", message!, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void TheFrameIsWhatIdentifiesIt()
    {
        // Guards the premise of every negative control below: the real exception really does carry
        // the frame the check looks for. If the parser's internals were renamed, the diagnosis would
        // silently stop firing and only this assertion would notice.
        Assert.Contains("GuessColumnsCount", RealParserFailure().StackTrace!, StringComparison.Ordinal);
    }

    // ---- the boundaries: what must NOT get this message --------------------------------------------

    [Fact]
    public void AnIndexErrorFromSomewhereElse_GetsNoDiagnosis()
    {
        // A REAL IndexOutOfRangeException with a REAL stack that does not contain the frame. This
        // discriminates "the frame identifies it" from "the exception type identifies it" — and the
        // second would put a table message on any index error anywhere in the conversion.
        Exception caught;
        try
        {
            var empty = Array.Empty<int>();
            _ = empty[3];
            throw new InvalidOperationException("unreachable");
        }
        catch (IndexOutOfRangeException ex)
        {
            caught = ex;
        }

        Assert.NotNull(caught.StackTrace);
        Assert.Null(HtmlFailureDiagnosis.Describe(caught));
    }

    /// <summary>An exception carrying a stack we choose, to test the type half of the conjunction.</summary>
    private sealed class Staged(string stack) : Exception
    {
        public override string StackTrace { get; } = stack;
    }

    [Fact]
    public void TheFrameAloneIsNotEnough_TheTypeMustMatchToo()
    {
        // The mirror of the test above. Without it, dropping the `ex is IndexOutOfRangeException`
        // check would survive every other test here, because nothing else carries the frame.
        // IndexOutOfRangeException is sealed, so the two halves have to be probed from opposite
        // directions: the frame from a staged exception of the wrong type, the type from a real one.
        Assert.Null(HtmlFailureDiagnosis.Describe(new Staged(
            "   at HtmlToOpenXml.Expressions.TableExpression.GuessColumnsCount(IHtmlTableElement t)")));
    }

    [Fact]
    public void AnExceptionWithNoStackAtAll_GetsNoDiagnosis()
    {
        Assert.Null(HtmlFailureDiagnosis.Describe(new IndexOutOfRangeException()));
    }

    [Fact]
    public void ADifferentExceptionType_GetsNoDiagnosis()
    {
        Assert.Null(HtmlFailureDiagnosis.Describe(new InvalidOperationException("anything")));
    }

    // ---- the call site still falls back, and still preserves the evidence ---------------------------

    [Fact]
    public async Task AnUnrecognisedFailureKeepsTheGenericWrapper()
    {
        // A vertical tab is not valid in XML, so this fails inside the writer — a real conversion
        // failure of a completely different kind. It proves the `?? generic` fallback is live: a
        // change that put the rowspan message on every failure would fail here.
        var ex = await Assert.ThrowsAsync<DocumentConversionException>(
            () => HtmlToDocxConverter.ConvertAsync("<p>ab</p>"));

        Assert.Contains("See the inner exception", ex.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("rowspan", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task AnUnrecognisedFailureStillCarriesItsInnerException()
    {
        // Naming a cause must never cost the evidence, on either branch.
        var ex = await Assert.ThrowsAsync<DocumentConversionException>(
            () => HtmlToDocxConverter.ConvertAsync("<p>ab</p>"));

        Assert.NotNull(ex.InnerException);
    }

    [Fact]
    public async Task OrdinaryHtmlStillConverts()
    {
        var docx = await HtmlToDocxConverter.ConvertAsync(
            "<h1>Title</h1><p>Body</p><table><tr><td>a</td><td>b</td></tr></table>");

        Assert.NotEmpty(docx);
    }
}
