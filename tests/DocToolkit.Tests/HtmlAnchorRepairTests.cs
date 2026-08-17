namespace DocToolkit.Tests;

/// <summary>
/// Repairing internal links so a page with old-style anchors can be rendered to PDF.
///
/// <b>Measured 2026-08-17 over 181 real `.gov` pages: HTML to PDF went from 106 to 115 (58.6% to
/// 63.5%), and the "bookmark link target was not found" group from 27 pages to 13.</b> Nothing that
/// converted before changed - verified by comparing the failure sets, and true by construction: the
/// input string is returned by reference unless a link is found that does not resolve, and such a
/// document fails today.
///
/// <b>The interesting test in here is <see cref="TheIdMustLandOnABlock_NotOnTheAnchorItself"/>.</b>
/// The obvious repair - give the <c>&lt;a name="x"&gt;</c> a matching <c>id</c> - is what this class
/// did first, and it fixed exactly nothing: a bookmark is created from an <c>id</c> on a BLOCK, and
/// an <c>id</c> on the anchor itself is ignored. Two rounds of corpus measurement went into finding
/// that, and a test that did not pin it would let the same wrong fix come back.
/// </summary>
public class HtmlAnchorRepairTests
{
    private const string Link = "<p><a href=\"#c\">Jump</a></p>";

    // ---- the guarantee: untouched unless a link does not resolve -----------------------------------

    [Theory]
    [InlineData("<p>no links at all</p>")]
    [InlineData("<p><a href=\"https://example.com\">external</a></p>")]
    [InlineData("<p style=\"color:#fff\">a hash that is not a link</p>")]
    [InlineData("<p><a href=\"#c\">J</a></p><p id=\"c\">target</p>")]      // already resolves
    public void NothingUnresolvable_TheSameStringComesBack(string html)
    {
        // ReferenceEquals, not equality: this asserts no parse-and-serialise round trip happened,
        // which is the basis for saying a document that converts today cannot change.
        Assert.Same(html, HtmlAnchorRepair.Apply(html));
    }

    [Fact]
    public void AnUnresolvableLink_IsTheOnlyThingThatCausesARewrite()
    {
        var html = Link + "<p><a name=\"c\">target</a></p>";

        Assert.NotSame(html, HtmlAnchorRepair.Apply(html));
    }

    // ---- the finding that cost two rounds of measurement --------------------------------------------

    [Fact]
    public void TheIdMustLandOnABlock_NotOnTheAnchorItself()
    {
        // Measured: an id on <p>, <div> or <h2> produces a bookmark; an id on <a>, <span> or <td>
        // does not. The first version of this class put the id on the <a> and changed nothing at all
        // across 181 pages, which is why this is pinned directly rather than only end to end.
        var result = HtmlAnchorRepair.Apply(Link + "<p><a name=\"c\">target</a></p>");

        Assert.Contains("<p id=\"c\">", result, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("<a name=\"c\" id=\"c\"", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task AnOldStyleAnchorNowRenders()
    {
        var pdf = await HtmlToPdfConverter.ConvertAsync(Link + "<p><a name=\"c\">target</a></p>");

        Assert.True(PdfProbe.IsPdf(pdf));
    }

    [Fact]
    public async Task TheLinkTextAndTheTargetTextBothSurvive()
    {
        // A repair that succeeded by deleting the link or the target would satisfy every assertion
        // about conversion. The document still has to say what it said.
        var pdf = await HtmlToPdfConverter.ConvertAsync(
            "<p><a href=\"#c\">JUMPTEXT</a></p><p><a name=\"c\">TARGETTEXT</a></p>");

        var text = PdfProbe.ExtractText(pdf);
        Assert.Contains("JUMPTEXT", text, StringComparison.Ordinal);
        Assert.Contains("TARGETTEXT", text, StringComparison.Ordinal);
    }

    [Fact]
    public void AnIdOnANonBlockIsNotTreatedAsResolved()
    {
        // The second rule that was wrong: "does any element carry this id" reads as resolved while
        // no bookmark is ever created, so the link was left in place and the document still failed.
        var html = Link + "<p><span id=\"c\">target</span></p>";

        var result = HtmlAnchorRepair.Apply(html);

        Assert.NotSame(html, result);
        Assert.Contains("<p id=\"c\">", result, StringComparison.OrdinalIgnoreCase);
    }

    // ---- genuinely absent targets ------------------------------------------------------------------

    [Fact]
    public async Task ALinkToNothingIsDropped_AndTheTextKept()
    {
        var pdf = await HtmlToPdfConverter.ConvertAsync("<p><a href=\"#gone\">KEEPME</a></p>");

        Assert.Contains("KEEPME", PdfProbe.ExtractText(pdf), StringComparison.Ordinal);
    }

    [Fact]
    public void ALinkToNothingLosesOnlyItsHref()
    {
        var result = HtmlAnchorRepair.Apply("<p><a href=\"#gone\">KEEPME</a></p>");

        Assert.DoesNotContain("href=\"#gone\"", result, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("KEEPME", result, StringComparison.Ordinal);
    }

    [Fact]
    public void AnExternalLinkIsNeverTouched()
    {
        // The repair is about internal navigation. Removing an http link would be a different and
        // much larger liberty, and nothing here has any reason to.
        var result = HtmlAnchorRepair.Apply(
            "<p><a href=\"#gone\">x</a><a href=\"https://example.com\">out</a></p>");

        Assert.Contains("https://example.com", result, StringComparison.Ordinal);
    }

    [Fact]
    public void ABlockThatAlreadyHasAnIdIsNotRelabelled()
    {
        // An id must be unique. Fixing a link by creating a duplicate would trade one malformed
        // document for another, so the link is dropped instead.
        var result = HtmlAnchorRepair.Apply(
            Link + "<p id=\"taken\"><a name=\"c\">target</a></p>");

        Assert.Contains("id=\"taken\"", result, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("href=\"#c\"", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void OnlyTargetsSomethingLinksToAreRelabelled()
    {
        // Relabelling a block nobody points at would be an edit with no purpose, and every edit here
        // is a liberty taken with somebody's document.
        var result = HtmlAnchorRepair.Apply(Link + "<p><a name=\"c\">t</a></p><p><a name=\"unused\">u</a></p>");

        Assert.DoesNotContain("id=\"unused\"", result, StringComparison.OrdinalIgnoreCase);
    }

    // ---- three rules that were wrong on the first pass, each measured -------------------------------

    [Fact]
    public async Task AnEmptyBlockIsNeverLabelled_BecauseThatTradesOneFailureForAWorseOne()
    {
        // Measured: a block with no text of its own does not merely fail to make a bookmark, it makes
        // the render throw from ThrowNoElementsException - an opaque crash rather than a legible
        // "target not found". The first version of this class did exactly that to EIGHT corpus pages.
        var pdf = await HtmlToPdfConverter.ConvertAsync(
            "<p><a href=\"#c\">Jump</a></p><div><a name=\"c\"></a></div>");

        Assert.True(PdfProbe.IsPdf(pdf));
    }

    [Fact]
    public void ABlockHoldingOnlyATableIsNotLabelled()
    {
        // Measured: a block whose only content is a table produces no paragraph of its own, so no
        // bookmark - labelling it leaves the document broken while looking repaired.
        var result = HtmlAnchorRepair.Apply(
            Link + "<div><table><tr><td><a name=\"c\">t</a></td></tr></table></div>");

        Assert.DoesNotContain("<div id=\"c\"", result, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("href=\"#c\"", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void AnEmptyBlockGetsNoStrayId()
    {
        // The sibling test above cannot tell whether Promote skipped the block or labelled it
        // pointlessly, because Satisfied refuses it either way and the link is dropped either way -
        // mutation testing showed exactly that. An id nobody can use is still an edit to somebody's
        // document, so it is asserted directly.
        var result = HtmlAnchorRepair.Apply(Link + "<div><a name=\"c\"></a></div>");

        Assert.DoesNotContain("<div id=\"c\"", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ABlockAlreadyCarryingTheIdIsNotTrusted_IfItCannotYieldABookmark()
    {
        // FOUR corpus pages are exactly this shape: the block already has the right id, so a naive
        // "is the id present" test says the link resolves - and it does not, because a block holding
        // only a table produces no bookmark. The link has to be dropped or the render fails.
        //
        // Nothing else in this file discriminates it: everywhere else the id is absent to begin with,
        // so both the right and the wrong rule reach the same answer. Mutation testing found that.
        const string html =
            "<p><a href=\"#c\">Jump</a></p><div id=\"c\"><table><tr><td>T</td></tr></table></div>";

        var result = HtmlAnchorRepair.Apply(html);

        Assert.NotSame(html, result);
        Assert.DoesNotContain("href=\"#c\"", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ABlockAlreadyCarryingTheIdButHoldingOnlyATableStillRenders()
    {
        var pdf = await HtmlToPdfConverter.ConvertAsync(
            "<p><a href=\"#c\">JUMPTEXT</a></p><div id=\"c\"><table><tr><td>CELL</td></tr></table></div>");

        var text = PdfProbe.ExtractText(pdf);
        Assert.Contains("JUMPTEXT", text, StringComparison.Ordinal);
        Assert.Contains("CELL", text, StringComparison.Ordinal);
    }

    [Fact]
    public async Task APreIsNotBookmarkable()
    {
        // `pre` was in the block list on the first pass and is measurably wrong: an id on it yields
        // no bookmark. The comment above that list says every entry must be measured; this is the
        // test that makes the claim falsifiable.
        var pdf = await HtmlToPdfConverter.ConvertAsync(
            "<p><a href=\"#c\">Jump</a></p><pre><a name=\"c\">t</a></pre>");

        Assert.True(PdfProbe.IsPdf(pdf));
    }

    [Fact]
    public async Task ARelativeUrlCarryingAFragmentIsInternalToo()
    {
        // Measured: `page.html#privacy` becomes an internal bookmark link, while
        // `https://host/page.html#privacy` stays an ordinary external one. Assuming only a bare
        // `#name` counted left four corpus pages failing with no visible cause.
        var pdf = await HtmlToPdfConverter.ConvertAsync("<p><a href=\"page.html#gone\">TEXT</a></p>");

        Assert.Contains("TEXT", PdfProbe.ExtractText(pdf), StringComparison.Ordinal);
    }

    [Fact]
    public void AnAbsoluteUrlWithAFragmentIsLeftAlone()
    {
        // The other half of that boundary. Stripping the href here would break a working external
        // link to fix a problem that does not exist.
        const string html = "<p><a href=\"https://example.com/p.html#frag\">x</a></p>";

        Assert.Same(html, HtmlAnchorRepair.Apply(html));
    }

    // ---- every entry point on the PDF path gets it ---------------------------------------------------

    /// <summary>
    /// <b>The repair is applied at four separate call sites, and nothing structural keeps them in
    /// step.</b> That is the same hand-maintained-inventory hazard that let all eight
    /// <c>PdfEditor</c> stream overloads go unguarded, so every public HTML-to-PDF entry point is
    /// exercised here rather than trusting the four to stay complete.
    /// </summary>
    public static TheoryData<string, Func<string, Task<byte[]>>> EveryPdfEntryPoint() => new()
    {
        { "ConvertAsync(html)", h => HtmlToPdfConverter.ConvertAsync(h) },
        { "ConvertAsync(html, page)", h => HtmlToPdfConverter.ConvertAsync(h, PageSetup.A4) },
        { "ConvertAsync(html, bool)", h => HtmlToPdfConverter.ConvertAsync(h, false) },
        { "ConvertAsync(html, options)", h => HtmlToPdfConverter.ConvertAsync(h, new RemoteImageOptions()) },
        { "ConvertAsync(html, page, options)", h => HtmlToPdfConverter.ConvertAsync(h, PageSetup.A4, new RemoteImageOptions()) },
    };

    [Theory]
    [MemberData(nameof(EveryPdfEntryPoint))]
    public async Task EveryEntryPointRepairsTheDocument(string name, Func<string, Task<byte[]>> convert)
    {
        var pdf = await convert(Link + "<p><a name=\"c\">target</a></p>");

        Assert.True(PdfProbe.IsPdf(pdf), $"{name} did not apply the repair");
    }

    [Fact]
    public async Task OrdinaryHtmlStillRenders()
    {
        var pdf = await HtmlToPdfConverter.ConvertAsync("<h1>Title</h1><p>Body</p>");

        Assert.Contains("Title", PdfProbe.ExtractText(pdf), StringComparison.Ordinal);
    }
}
