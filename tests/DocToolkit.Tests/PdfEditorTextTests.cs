using Xunit;

namespace DocToolkit.Tests;

/// <summary>
/// Reading text back out of a PDF (A18).
///
/// PdfProbe is the ORACLE here, not the thing under test. It is a hand-rolled parser that has
/// verified every PDF this library produces since before PdfPig existed, and it is itself pinned
/// against the literal "AcmeCorp" plus the WinAnsi em-dash case. Two independently written
/// extractors agreeing is real evidence - but only because one of them is anchored, so every test
/// below also asserts a literal.
/// </summary>
public class PdfEditorTextTests
{
    private static Task<byte[]> PdfAsync(string html) => HtmlToPdfConverter.ConvertAsync(html);

    [Fact]
    public async Task ReadsTheTextOfASinglePage()
    {
        var pdf = await PdfAsync("<h1>Acme Corporation</h1><p>Invoice 42</p>");

        var pages = PdfEditor.ExtractText(pdf);

        Assert.Single(pages);
        Assert.Contains("Acme Corporation", pages[0], StringComparison.Ordinal);
        Assert.Contains("Invoice 42", pages[0], StringComparison.Ordinal);
    }

    [Fact]
    public async Task SeparatesBlocksRatherThanRunningThemTogether()
    {
        // THE test of this feature. PdfPig's obvious entry point, page.Text, returns
        // "Acme CorporationInvoice 42" with no separator - which is exactly the A26 defect
        // ("TitleBody text.") fixed in DocxEditor.ExtractText two days before this was written.
        // Asserting Contains for each piece would pass against that. Only this fails.
        var pdf = await PdfAsync("<h1>Acme Corporation</h1><p>Invoice 42</p>");

        var text = PdfEditor.ExtractText(pdf)[0];

        Assert.DoesNotContain("CorporationInvoice", text, StringComparison.Ordinal);
    }

    [Fact]
    public async Task ReturnsOneEntryPerPageInDocumentOrder()
    {
        var merged = PdfEditor.Merge(
            [await PdfAsync("<h1>First page</h1>"), await PdfAsync("<h1>Second page</h1>")]);

        var pages = PdfEditor.ExtractText(merged);

        // Asserting WHICH page holds what, not just the count: a count passes against an
        // implementation that returned the same page twice.
        Assert.Equal(2, pages.Count);
        Assert.Contains("First page", pages[0], StringComparison.Ordinal);
        Assert.Contains("Second page", pages[1], StringComparison.Ordinal);
        Assert.DoesNotContain("Second page", pages[0], StringComparison.Ordinal);
    }

    [Fact]
    public async Task AgreesWithThePdfProbeOracle()
    {
        // Two independently written extractors on the same document. Sound only because PdfProbe
        // is itself pinned - see the class comment.
        var pdf = await PdfAsync("<h1>Acme Corporation</h1>");

        var mine = string.Concat(PdfEditor.ExtractText(pdf)).Replace("\n", string.Empty);
        var oracle = PdfProbe.ExtractText(pdf).Replace("\n", string.Empty);

        Assert.Equal(oracle, mine);
        Assert.Contains("Acme", mine, StringComparison.Ordinal);
    }

    [Fact]
    public void SomethingThatIsNotAPdfIsRejected()
    {
        var notAPdf = System.Text.Encoding.UTF8.GetBytes("This is not a PDF.");

        Assert.Throws<DocumentConversionException>(() => PdfEditor.ExtractText(notAPdf));
    }

    [Fact]
    public void MissingInputIsRejectedBeforeAnyWork()
    {
        Assert.Throws<ArgumentNullException>(() => PdfEditor.ExtractText(null!));
        Assert.Throws<ArgumentException>(() => PdfEditor.ExtractText([]));
    }
}
