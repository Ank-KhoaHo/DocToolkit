namespace DocToolkit.Tests;

/// <summary>
/// Filling a text-less table cell so the renderer's own auto-layout does not squeeze it to a width
/// it then refuses.
///
/// <b>Measured 2026-08-17 over 181 real `.gov` pages: HTML to PDF went from 131 to 143 (72.4% to
/// 79.0%), and the "positive text width" group from 16 pages to 4.</b>
///
/// <b>The message names padding, and that reading is wrong.</b> Reduced from a real page to 316
/// bytes, the trigger is a whitespace-only cell beside a cell of long text - with <b>no width
/// specified anywhere</b>. Automatic layout gives the empty cell a near-zero width and the renderer
/// then rejects the layout it just computed. Measured: <c>width="20"</c> on the spacer,
/// <c>width="1"</c>, <c>style="padding:0"</c>, <c>cellpadding="0"</c> and a table
/// <c>width="100%"</c> all still fail. A non-breaking space is what works, which is exactly what
/// authors of that era wrote in spacer cells.
///
/// <b>This one is applied only AFTER a failure, and that is the load-bearing difference from the
/// other repairs.</b> A table cell with no text is completely ordinary and usually renders perfectly
/// well, so filling every one up front would edit documents that were fine. The tests that matter
/// most here are the ones pinning that it stays out of the up-front path.
/// </summary>
public class EmptyTableCellRepairTests
{
    private const string LongText =
        "This project will assess the effects of off-road vehicle traffic on the beach and dunes";

    private static string Failing =>
        $"<table><tr><td> </td><td>{LongText}</td></tr></table>";

    // ---- it is NOT applied up front -----------------------------------------------------------------

    [Fact]
    public void PrepareLeavesTextLessCellsAlone()
    {
        // The whole safety argument: a document that renders on the first attempt is never repaired.
        // If this repair ever moves into Prepare, every document with an empty cell gets edited -
        // including the overwhelming majority that convert perfectly well.
        const string html = "<table><tr><td></td><td>text</td></tr></table>";

        Assert.Same(html, HtmlForPdf.Prepare(html));
    }

    [Fact]
    public void ADocumentWithNoTablesIsNeverParsed()
    {
        const string html = "<p>no tables here</p>";

        Assert.Same(html, EmptyTableCellRepair.Apply(html));
    }

    [Fact]
    public void ATableWhoseCellsAllHaveTextIsUnchanged()
    {
        const string html = "<table><tr><td>a</td><td>b</td></tr></table>";

        Assert.Same(html, EmptyTableCellRepair.Apply(html));
    }

    // ---- what it does when it does run ---------------------------------------------------------------

    [Fact]
    public async Task TheFailingShapeNowRenders()
    {
        var pdf = await HtmlToPdfConverter.ConvertAsync(Failing);

        Assert.True(PdfProbe.IsPdf(pdf));
    }

    [Fact]
    public async Task TheRealTextSurvives()
    {
        // A repair that succeeded by dropping the long cell would satisfy the assertion above.
        var pdf = await HtmlToPdfConverter.ConvertAsync(Failing);

        Assert.Contains("off-road vehicle", PdfProbe.ExtractText(pdf), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("<td></td>")]
    [InlineData("<td> </td>")]
    [InlineData("<td><br></td>")]
    [InlineData("<td><p><br></p></td>")]
    [InlineData("<td><img alt=\"\"></td>")]
    public async Task EveryShapeOfTextLessCellIsHandled(string cell)
    {
        // The real pages carry <br>, <p><br></p> and <img> in these cells - all of which render to
        // nothing. Restricting the rule to structurally EMPTY cells fixed one page out of seventeen,
        // which is how this theory came to exist.
        var pdf = await HtmlToPdfConverter.ConvertAsync($"<table><tr>{cell}<td>{LongText}</td></tr></table>");

        Assert.True(PdfProbe.IsPdf(pdf));
    }

    [Fact]
    public void TheCellsExISTINGCONTENTIsKept_NotReplaced()
    {
        // Appended, not assigned. Replacing the content would delete an image the caller may well
        // have asked to be embedded.
        var result = EmptyTableCellRepair.Apply(
            $"<table><tr><td><img alt=\"\" src=\"logo.gif\"></td><td>{LongText}</td></tr></table>");

        Assert.Contains("logo.gif", result, StringComparison.Ordinal);
    }

    // ---- the retry fires for the right reason and no other -------------------------------------------

    [Fact]
    public void WouldHelpMatchesTheRenderersOwnMessage()
    {
        var ex = new DocumentConversionException(
            "wrapper", new InvalidOperationException(
                "Table horizontal cell padding must leave a positive text width."));

        Assert.True(EmptyTableCellRepair.WouldHelp(ex));
    }

    [Theory]
    [InlineData("PDF bookmark link target 'content' was not found.")]
    [InlineData("Parameter 'linkContents' cannot be empty or whitespace.")]
    [InlineData("Something else entirely.")]
    public void WouldHelpRefusesEveryOtherFailure(string message)
    {
        // Matched on the message, not the exception type, because the type says nothing: every
        // conversion failure arrives as a DocumentConversionException. A retry that fired on all of
        // them would double the cost of every failure and fix none of the others.
        var ex = new DocumentConversionException("wrapper", new InvalidOperationException(message));

        Assert.False(EmptyTableCellRepair.WouldHelp(ex));
    }

    [Fact]
    public void WouldHelpRefusesAnExceptionWithNoInner()
    {
        Assert.False(EmptyTableCellRepair.WouldHelp(new DocumentConversionException("bare")));
    }

    [Fact]
    public async Task AFailureNoRepairClaimsIsRethrownUnchanged()
    {
        // The retry must not swallow or relabel a failure it has no answer for.
        //
        // This used to use an empty link in a table cell, which ImageLinkRepair now fixes - so the
        // test went green-by-success and had to be re-aimed. A vertical tab is not valid in XML, so
        // this fails in the DOCX stage before any PDF repair could apply, and no repair claims it.
        // U+000B is written as an ESCAPE, not as a literal character. An invisible control
        // character in source is exactly the kind of thing an editor or a patch silently drops,
        // and this test passes vacuously the moment it does - which happened twice while writing
        // it, in both directions.
        var ex = await Assert.ThrowsAsync<DocumentConversionException>(
            () => HtmlToPdfConverter.ConvertAsync("<p>a\u000Bb</p>"));

        Assert.NotNull(ex.InnerException);
        Assert.DoesNotContain("positive text width", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task OrdinaryTablesStillRender()
    {
        var pdf = await HtmlToPdfConverter.ConvertAsync(
            "<table><tr><td>CELLA</td><td>CELLB</td></tr></table>");

        var text = PdfProbe.ExtractText(pdf);
        Assert.Contains("CELLA", text, StringComparison.Ordinal);
        Assert.Contains("CELLB", text, StringComparison.Ordinal);
    }
}
