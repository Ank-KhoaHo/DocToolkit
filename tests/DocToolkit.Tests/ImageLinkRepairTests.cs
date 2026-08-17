namespace DocToolkit.Tests;

/// <summary>
/// Labelling a link that wraps only an image, so a table cell holding one can be rendered.
///
/// <b>Measured 2026-08-17 over 181 real `.gov` pages: HTML to PDF went from 143 to 159 (79.0% to
/// 87.8%)</b>, closing both <c>linkContents</c> groups outright.
///
/// <b>This is the first repair here with no browser behaviour to appeal to</b>, which is why it was
/// a maintainer decision rather than a measurement. Unwrapping keeps the image and loses the
/// navigation; using the image's <c>alt</c> keeps the navigation and replaces the image with words.
/// The maintainer chose the alt text, and the default path makes that close to free - this package
/// does not fetch remote images, so on the default path the image is not in the output anyway and
/// unwrapping would leave an empty cell.
/// </summary>
public class ImageLinkRepairTests
{
    private const string Cell = "<table><tr><td>{0}</td><td>other</td></tr></table>";

    private static string Table(string cellContent) => string.Format(Cell, cellContent);

    // ---- what it does -------------------------------------------------------------------------------

    [Fact]
    public void TheAltTextBecomesTheLinkText()
    {
        var result = ImageLinkRepair.Apply(
            Table("<a href=\"https://e.com\"><img alt=\"NOAA logo\" src=\"l.gif\"></a>"));

        // `Contains("NOAA logo")` alone passes even when the link is unwrapped and no text is added
        // at all - the words are still there in the alt ATTRIBUTE. Mutation testing caught that: a
        // mutant that never uses the alt was killed by three other tests and not by this one, which
        // is supposed to be the test for exactly it. Asserting the position pins what was meant.
        Assert.Contains("NOAA logo</a>", result, StringComparison.Ordinal);
    }

    [Fact]
    public void TheImageSurvives_BecauseTheTextIsAppendedNotSubstituted()
    {
        // A caller who DID ask for images to be embedded must still get the image. Replacing the
        // link's content would delete it.
        var result = ImageLinkRepair.Apply(
            Table("<a href=\"https://e.com\"><img alt=\"NOAA logo\" src=\"l.gif\"></a>"));

        Assert.Contains("l.gif", result, StringComparison.Ordinal);
    }

    [Fact]
    public void TheLinkSurvives_WhichIsTheWholePointOfChoosingAltOverUnwrapping()
    {
        var result = ImageLinkRepair.Apply(
            Table("<a href=\"https://e.com\"><img alt=\"Home\" src=\"l.gif\"></a>"));

        Assert.Contains("https://e.com", result, StringComparison.Ordinal);
    }

    [Fact]
    public void WithNoUsableAltTheLinkIsUnwrapped_NotGivenAPlaceholder()
    {
        // There is nothing to label it with, and inventing text would put words in somebody's
        // document that they never wrote. The image stays exactly where it was.
        var result = ImageLinkRepair.Apply(
            Table("<a href=\"https://e.com\"><img alt=\"\" src=\"spacer.gif\"></a>"));

        Assert.Contains("spacer.gif", result, StringComparison.Ordinal);
        Assert.DoesNotContain("https://e.com", result, StringComparison.Ordinal);
    }

    [Fact]
    public async Task TheFailingShapeNowRenders()
    {
        var pdf = await HtmlToPdfConverter.ConvertAsync(
            Table("<a href=\"https://e.com\"><img alt=\"Logo\" src=\"l.gif\"></a>"));

        Assert.True(PdfProbe.IsPdf(pdf));
    }

    [Fact]
    public async Task TheAltTextReachesThePdf()
    {
        var pdf = await HtmlToPdfConverter.ConvertAsync(
            Table("<a href=\"https://e.com\"><img alt=\"ALTSENTINEL\" src=\"l.gif\"></a>"));

        Assert.Contains("ALTSENTINEL", PdfProbe.ExtractText(pdf), StringComparison.Ordinal);
    }

    // ---- what it leaves alone -----------------------------------------------------------------------

    [Fact]
    public void ALinkOUTSIDEATableIsUntouched()
    {
        // Measured: the same link renders perfectly well outside a table. Rewriting it would be an
        // edit with no purpose.
        //
        // The document MUST contain a table, or this proves nothing: without one the cheap
        // pre-check returns early and the selector is never consulted, so a widened selector would
        // survive. Mutation testing found exactly that.
        const string html =
            "<p><a href=\"https://e.com\"><img alt=\"Logo\"></a></p>"
            + "<table><tr><td>has text</td></tr></table>";

        Assert.Same(html, ImageLinkRepair.Apply(html));
    }

    [Fact]
    public void ALinkWithTextIsUntouched()
    {
        var html = Table("<a href=\"https://e.com\">Read more</a>");

        Assert.Same(html, ImageLinkRepair.Apply(html));
    }

    [Fact]
    public void ADocumentWithNoTablesIsNeverParsed()
    {
        const string html = "<p>nothing here</p>";

        Assert.Same(html, ImageLinkRepair.Apply(html));
    }

    [Fact]
    public void ItIsNotAppliedUpFront()
    {
        // Same discipline as EmptyTableCellRepair: only a document that has already failed is edited.
        var html = Table("<a href=\"https://e.com\"><img alt=\"Logo\"></a>");

        Assert.Same(html, HtmlForPdf.Prepare(html));
    }

    // ---- the retry selects on the failure, not on the type -------------------------------------------

    [Fact]
    public void WouldHelpMatchesTheRenderersOwnMessage()
    {
        var ex = new DocumentConversionException("wrapper",
            new ArgumentException("Parameter 'linkContents' cannot be empty or whitespace."));

        Assert.True(ImageLinkRepair.WouldHelp(ex));
    }

    [Theory]
    [InlineData("Table horizontal cell padding must leave a positive text width.")]
    [InlineData("PDF bookmark link target 'content' was not found.")]
    [InlineData("Parameter 'text' cannot be empty or whitespace.")]
    public void WouldHelpRefusesEveryOtherFailure(string message)
    {
        // The third case matters: it is a DIFFERENT empty-parameter failure from the same renderer,
        // and a looser match would send this repair after a problem it cannot fix.
        var ex = new DocumentConversionException("wrapper", new ArgumentException(message));

        Assert.False(ImageLinkRepair.WouldHelp(ex));
    }

    // ---- a page can need more than one repair ---------------------------------------------------------

    [Fact]
    public async Task ADocumentNeedingBOTHRepairsGetsBoth()
    {
        // Real pages do this: filling the empty cells reveals that the image links are also
        // unlabelled. A single retry would fix one and report the other, so the render loops until
        // no known repair matches - and this is the test that pins it.
        const string html =
            "<table><tr>"
            + "<td></td>"
            + "<td><a href=\"https://e.com\"><img alt=\"LOGOALT\" src=\"l.gif\"></a></td>"
            + "<td>A long sentence of ordinary text that squeezes the empty cell to nothing at all</td>"
            + "</tr></table>";

        var pdf = await HtmlToPdfConverter.ConvertAsync(html);

        var text = PdfProbe.ExtractText(pdf);
        Assert.Contains("LOGOALT", text, StringComparison.Ordinal);
        Assert.Contains("long sentence", text, StringComparison.Ordinal);
    }

    [Fact]
    public async Task OrdinaryTablesStillRender()
    {
        var pdf = await HtmlToPdfConverter.ConvertAsync(
            "<table><tr><td><a href=\"https://e.com\">LINKTEXT</a></td><td>CELLB</td></tr></table>");

        var text = PdfProbe.ExtractText(pdf);
        Assert.Contains("LINKTEXT", text, StringComparison.Ordinal);
        Assert.Contains("CELLB", text, StringComparison.Ordinal);
    }
}
