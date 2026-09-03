using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace DocToolkit.Extensions.DependencyInjection.Tests;

/// <summary>
/// <see cref="DocxCompareService"/> — the dependency-injection mirror of
/// <see cref="DocToolkit.DocxCompare"/>, which landed one release behind it because the extensions
/// package builds against the PUBLISHED core.
///
/// <b>This service is pure delegation, so the risk is not that the comparison is wrong — it is
/// that a member delegates to the WRONG thing, or silently does nothing.</b> The DI package is held
/// at 100% line and branch coverage precisely because an uncovered member here is a member nobody
/// checked was wired to anything.
///
/// So every assertion below is one a passthrough would fail:
///
/// <list type="bullet">
/// <item><c>Compare</c> is asserted to produce a document that CARRIES revisions. A member
/// returning either input unchanged fails, and so does one returning an empty document.</item>
/// <item>The negative control asserts a document compared with ITSELF carries none — which the
/// same stub would also fail, in the opposite direction. Neither test alone distinguishes a
/// working comparison from a member that marks everything.</item>
/// <item><c>CompareWithReport</c> is asserted on its WARNINGS, which is the only thing separating
/// it from <c>Compare</c> at the call site. A mirror wired to the wrong one loses them.</item>
/// <item>Every guard is asserted to throw, which a member doing nothing would not.</item>
/// </list>
/// </summary>
public class DocxCompareServiceTests
{
    private const string Author = "Reviewer";

    private static IDocxCompare Service() => new DocxCompareService();

    private static byte[] Doc(params string[] paragraphs) =>
        DocxEditor.Create([.. paragraphs.Select(DocxBlock.Paragraph)]);

    /// <summary>How many tracked insertions and deletions a document carries.</summary>
    /// <remarks>
    /// Counted from the OOXML rather than through <c>DocxReview</c>, so this test does not pass or
    /// fail on account of a DIFFERENT capability's behaviour. OpenXml arrives transitively through
    /// <c>Ank.DocToolkit</c>, so this adds no dependency.
    /// </remarks>
    private static int RevisionCount(byte[] docx)
    {
        using var ms = new MemoryStream(docx);
        using var document = WordprocessingDocument.Open(ms, false);

        var body = document.MainDocumentPart!.Document!.Body!;
        return body.Descendants<InsertedRun>().Count() + body.Descendants<DeletedRun>().Count();
    }

    // ---- Compare ---------------------------------------------------------------------------

    [Fact]
    public void Compare_MarksTheDifferenceRatherThanReturningEitherInput()
    {
        var original = Doc("the quick brown fox");
        var revised = Doc("the quick red fox");

        var compared = Service().Compare(original, revised, Author);

        // A member returning `revised` unchanged carries no revisions and fails here; so does one
        // returning `original`, and so does one returning an empty document.
        Assert.True(RevisionCount(compared) > 0);
        Assert.Equal(0, RevisionCount(revised));
    }

    [Fact]
    public void Compare_NegativeControl_ADocumentComparedWithItselfCarriesNoRevisions()
    {
        var docx = Doc("nothing has changed here");

        var compared = Service().Compare(docx, docx, Author);

        // Without this, a member that marked EVERYTHING changed would satisfy the test above.
        Assert.Equal(0, RevisionCount(compared));
    }

    [Fact]
    public void Compare_RecordsTheAuthorItWasGiven()
    {
        var compared = Service().Compare(Doc("before"), Doc("after"), "Someone Specific");

        using var ms = new MemoryStream(compared);
        using var document = WordprocessingDocument.Open(ms, false);
        var body = document.MainDocumentPart!.Document!.Body!;

        // A member dropping the author argument on the floor still produces revisions.
        Assert.Contains(
            body.Descendants<InsertedRun>().Select(i => i.Author?.Value),
            a => a == "Someone Specific");
    }

    // ---- CompareWithReport -----------------------------------------------------------------

    [Fact]
    public void CompareWithReport_CarriesTheWarningsThatCompareDiscards()
    {
        var result = Service().CompareWithReport(Doc("one"), Doc("two"), Author);

        // The warnings are the ONLY thing distinguishing this member from Compare at the call
        // site. A mirror wired to the wrong one returns a result with none.
        Assert.NotEmpty(result.Warnings);
        Assert.True(result.HasLoss);
    }

    [Fact]
    public void CompareWithReport_ReturnsTheSameMarkedUpDocumentCompareDoes()
    {
        var original = Doc("the quick brown fox");
        var revised = Doc("the quick red fox");
        var service = Service();

        var withReport = service.CompareWithReport(original, revised, Author);

        // Not a byte comparison: each call stamps its revisions with DateTime.UtcNow, so two runs
        // differ in the timestamp alone. What must agree is the comparison itself.
        Assert.Equal(RevisionCount(service.Compare(original, revised, Author)), RevisionCount(withReport.Value));
    }

    // ---- guards ----------------------------------------------------------------------------

    [Fact]
    public void Compare_RefusesNullAndEmptyAndBlank()
    {
        var docx = Doc("content");
        var service = Service();

        Assert.Throws<ArgumentNullException>(() => service.Compare(null!, docx, Author));
        Assert.Throws<ArgumentNullException>(() => service.Compare(docx, null!, Author));
        Assert.Throws<ArgumentNullException>(() => service.Compare(docx, docx, null!));
        Assert.Throws<ArgumentException>(() => service.Compare([], docx, Author));
        Assert.Throws<ArgumentException>(() => service.Compare(docx, [], Author));
        Assert.Throws<ArgumentException>(() => service.Compare(docx, docx, "   "));
    }

    [Fact]
    public void CompareWithReport_RefusesNullAndEmptyAndBlank()
    {
        var docx = Doc("content");
        var service = Service();

        Assert.Throws<ArgumentNullException>(() => service.CompareWithReport(null!, docx, Author));
        Assert.Throws<ArgumentNullException>(() => service.CompareWithReport(docx, null!, Author));
        Assert.Throws<ArgumentNullException>(() => service.CompareWithReport(docx, docx, null!));
        Assert.Throws<ArgumentException>(() => service.CompareWithReport([], docx, Author));
        Assert.Throws<ArgumentException>(() => service.CompareWithReport(docx, [], Author));
        Assert.Throws<ArgumentException>(() => service.CompareWithReport(docx, docx, "   "));
    }
}
