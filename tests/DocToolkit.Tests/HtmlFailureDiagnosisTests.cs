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
        // U+FFFE is a NONCHARACTER: legal in a C# string, refused by the XML writer with
        // "String contains invalid Unicode code points" - a real conversion failure of a third
        // kind. It proves the `?? generic` fallback is live: a change that put any diagnosis on
        // every failure would fail here.
        //
        // This used to be a vertical tab, and that stopped working the moment the invalid-character
        // diagnosis shipped: a control character is now RECOGNISED, so it can no longer play the
        // part of an unrecognised failure. Replacing it rather than deleting these two tests is
        // the point - the fallback still needs proving, and it needs a REAL failure to prove it
        // with, not a staged exception.
        var ex = await Assert.ThrowsAsync<DocumentConversionException>(
            () => HtmlToDocxConverter.ConvertAsync("<p>a\uFFFEb</p>"));

        // "for details" is the GENERIC wrapper's wording, and the discrimination is deliberate:
        // both diagnoses also end with "See the inner exception", so asserting that phrase alone
        // would pass whether or not one of them had fired. This assertion has to be able to fail.
        Assert.Contains("See the inner exception for details", ex.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("rowspan", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("control character", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task AnUnrecognisedFailureStillCarriesItsInnerException()
    {
        // Naming a cause must never cost the evidence, on either branch.
        var ex = await Assert.ThrowsAsync<DocumentConversionException>(
            () => HtmlToDocxConverter.ConvertAsync("<p>a\uFFFEb</p>"));

        Assert.NotNull(ex.InnerException);
    }

    [Fact]
    public async Task OrdinaryHtmlStillConverts()
    {
        var docx = await HtmlToDocxConverter.ConvertAsync(
            "<h1>Title</h1><p>Body</p><table><tr><td>a</td><td>b</td></tr></table>");

        Assert.NotEmpty(docx);
    }

    // ---- the control character: what a caller gets for passing something that is not HTML ----------
    //
    // Measured 2026-08-20 over govdocs1: 8 of 8 JPEGs and 1 of 12 .txt files fail this way, and the
    // caller was told only "See the inner exception" over "hexadecimal value 0x10 is an invalid
    // character" - a message about a character nobody typed.

    [Theory]
    [InlineData("\u000B")]  // vertical tab
    [InlineData("\u0010")]  // what a JPEG's bytes produce first
    [InlineData("\u0001")]  // what a PDF's bytes produce first
    [InlineData("\u001F")]  // the top of the C0 range
    public async Task AControlCharacterIsNamedRatherThanLeftToTheInnerException(string control)
    {
        var ex = await Assert.ThrowsAsync<DocumentConversionException>(
            () => HtmlToDocxConverter.ConvertAsync($"<p>before{control}after</p>"));

        Assert.Contains("control character", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("rowspan", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ItQuotesTheCharacterBack()
    {
        // So the reader does not have to open the inner exception to learn which one it was.
        var ex = await Assert.ThrowsAsync<DocumentConversionException>(
            () => HtmlToDocxConverter.ConvertAsync("<p>a\u0010b</p>"));

        Assert.Contains("0x10", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public async Task ItNamesBothCausesAndAssertsNeither()
    {
        // The exception CANNOT tell binary content from a stray control byte in real HTML - both
        // arrive here byte-identical. Claiming either as fact is the mistake this class's remarks
        // already record about the old timeout message, so the test pins that it does not.
        var ex = await Assert.ThrowsAsync<DocumentConversionException>(
            () => HtmlToDocxConverter.ConvertAsync("<p>a\u0010b</p>"));

        Assert.Contains("not HTML at all", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("genuine HTML", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("cannot tell them apart", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task RealBinaryContentGetsIt()
    {
        // The case that motivated the row, exercised end to end rather than through a staged string:
        // a real PNG's bytes, read as text and handed to the HTML converter.
        var png = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==");

        var ex = await Assert.ThrowsAsync<DocumentConversionException>(
            () => HtmlToDocxConverter.ConvertAsync(System.Text.Encoding.Latin1.GetString(png)));

        Assert.Contains("control character", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    // ---- the boundaries, which is where this earns its place -----------------------------------------

    [Theory]
    [InlineData("<p>a\tb</p>")]                       // tab
    [InlineData("<p>a\r\nb</p>")]                     // CR LF
    [InlineData("<p>&amp; &lt; &#233; &#10;</p>")]    // entities, including a numeric line feed
    [InlineData("<p>café naïve</p>")]       // accented Latin
    [InlineData("<p>漢字</p>")]               // CJK
    [InlineData("<p>hi \U0001F600</p>")]              // astral plane
    public async Task OrdinaryTextIsUnaffected(string html)
    {
        // The message claims all of these convert. Without this, that claim is prose nothing checks
        // - and a matcher widened to any ArgumentException would still pass every test above.
        Assert.NotEmpty(await HtmlToDocxConverter.ConvertAsync(html));
    }

    [Fact]
    public async Task ADifferentArgumentException_GetsNoDiagnosis()
    {
        // The type half is not sufficient, proved with a REAL failure rather than a staged one:
        // U+FFFE is a noncharacter, refused by the same writer, same exception type, different
        // message - "String contains invalid Unicode code points". A matcher testing only
        // `ex is ArgumentException` would put a control-character message on it.
        var thrown = await Assert.ThrowsAsync<DocumentConversionException>(
            () => HtmlToDocxConverter.ConvertAsync("<p>a\uFFFEb</p>"));
        var caught = thrown.InnerException;

        Assert.IsType<ArgumentException>(caught);
        Assert.DoesNotContain("hexadecimal value 0x", caught!.Message, StringComparison.Ordinal);
        Assert.Null(HtmlFailureDiagnosis.Describe(caught));
    }

    [Fact]
    public void TheMessageAloneIsNotEnough_TheTypeMustMatchToo()
    {
        // The mirror. Without it, dropping `ex is ArgumentException` survives every test here.
        Assert.Null(HtmlFailureDiagnosis.Describe(
            new InvalidOperationException("'x', hexadecimal value 0x10, is an invalid character.")));
    }
}
