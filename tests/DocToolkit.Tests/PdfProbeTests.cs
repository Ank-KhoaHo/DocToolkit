using System.IO;
using Xunit;

namespace DocToolkit.Tests;

public class PdfProbeTests
{
    private static readonly string ResultPdfPath = Path.Combine(AppContext.BaseDirectory, "Fixtures", "result.pdf");
    private static readonly string BigPdfPath = Path.Combine(AppContext.BaseDirectory, "Fixtures", "big.pdf");

    [Fact]
    public void ExtractText_DecodesHexStringTextOperators()
    {
        // "Acme" == 41 63 6D 65 ; "Corp" == 43 6F 72 70
        var pdf = System.Text.Encoding.Latin1.GetBytes(
            "%PDF-1.4\n5 0 obj\n<< /Length 40 >>\nstream\nBT\n<41636D65> Tj\n<436F7270> Tj\nET\nendstream\nendobj\n");

        Assert.Equal("AcmeCorp", PdfProbe.ExtractText(pdf));
    }

    [Fact]
    public void ExtractText_Decodes0x80To0x9FAsWinAnsiNotLatin1()
    {
        // WinAnsiEncoding (roughly Windows-1252) maps 0x80-0x9F to typographic characters;
        // Latin-1 maps that same range to C1 control codes. Word-authored content is full of
        // these: em-dash, smart quotes, ellipsis, trademark. Byte 0x97 is what OfficeIMO's
        // real result.pdf actually contains for an em-dash - see the test below for that.
        var pdf = System.Text.Encoding.Latin1.GetBytes(
            "%PDF-1.4\nBT\n<9791928599> Tj\nET\n");

        // byte 0x97=em-dash, 0x91/0x92=left/right single quote, 0x85=ellipsis, 0x99=trademark
        var expected = "—‘’…™";
        Assert.Equal(expected, PdfProbe.ExtractText(pdf));
    }

    [Fact]
    public void ExtractText_OnRealPdf_DecodesTheEmDashByteCorrectly()
    {
        // result.pdf genuinely contains byte 0x97 as a Tj operator (established by
        // inspection). Under the old Latin-1 decode this produced U+0097 (a C1 control
        // character); the correct WinAnsi decoding is U+2014, an em-dash.
        var text = PdfProbe.ExtractText(File.ReadAllBytes(ResultPdfPath));

        Assert.Contains('—', text);
        Assert.DoesNotContain('\u0097', text);
    }

    [Fact]
    public void PageCount_ReadsTheCountFromThePageTree()
    {
        var pdf = System.Text.Encoding.Latin1.GetBytes(
            "%PDF-1.4\n1 0 obj\n<< /Type /Pages /Count 7 /Kids [ 7 0 R ] >>\nendobj\n");

        Assert.Equal(7, PdfProbe.PageCount(pdf));
    }

    [Fact]
    public void PageCount_OnRealMultiPagePdf_ReturnsTheDocumentTotal()
    {
        // big.pdf is a genuine 5-page OfficeIMO-generated PDF. This is the assertion that
        // proves the /Root -> /Catalog -> /Pages resolution (and/or max-Count fallback)
        // against a real multi-page document, not just a synthetic single-node fixture.
        Assert.Equal(5, PdfProbe.PageCount(File.ReadAllBytes(BigPdfPath)));
    }

    [Fact]
    public void PageCount_ReturnsTheDocumentTotalNotTheFirstPagesNodeEncountered()
    {
        // A page tree can have an intermediate /Type /Pages node (a partial-count subtree)
        // that appears earlier in the byte stream than the root /Pages node referenced by the
        // catalog. The unscoped/unanchored old regex just grabbed the first "/Count" it met
        // after the first "/Type /Pages" in the whole document - here that would be the
        // intermediate node's count of 2, not the document's real total of 4.
        var pdf = System.Text.Encoding.Latin1.GetBytes(
            "%PDF-1.4\n" +
            "5 0 obj\n<< /Type /Pages /Kids [ 6 0 R 7 0 R ] /Count 2 >>\nendobj\n" +
            "1 0 obj\n<< /Type /Pages /Kids [ 5 0 R 8 0 R 9 0 R 10 0 R ] /Count 4 >>\nendobj\n" +
            "17 0 obj\n<< /Type /Catalog /Pages 1 0 R >>\nendobj\n" +
            "trailer\n<< /Size 18 /Root 17 0 R >>\n");

        Assert.Equal(4, PdfProbe.PageCount(pdf));
    }

    [Fact]
    public void PageCount_DoesNotAssumeTypePrecedesCountInTheDictionary()
    {
        // PDF dictionary key order is unspecified. A /Pages dictionary with /Count written
        // before /Type is just as valid as one with /Type first.
        var pdf = System.Text.Encoding.Latin1.GetBytes(
            "%PDF-1.4\n1 0 obj\n<< /Count 7 /Type /Pages /Kids [ 7 0 R ] >>\nendobj\n");

        Assert.Equal(7, PdfProbe.PageCount(pdf));
    }

    [Fact]
    public void PageCount_OnRealSinglePagePdf_ReturnsTheDocumentTotal()
    {
        // result.pdf is a genuine 1-page OfficeIMO-generated PDF (ground truth). This is the
        // Finding-D coverage gap: nothing previously asserted PageCount against the single-page
        // fixture, only the 5-page one.
        Assert.Equal(1, PdfProbe.PageCount(File.ReadAllBytes(ResultPdfPath)));
    }

    [Fact]
    public void PageCount_ViaCatalog_WinsOverMaxCountFallback_OnRealisticIdBearingTrailer()
    {
        // Finding A: every real fixture PDF (result.pdf, big.pdf) has a trailer of the shape
        // "trailer << /Size N /Root R 0 R /Info I 0 R /ID [<HEX> <HEX>] >>" - a single-angle-
        // bracket hex-string token that the old DictBody alternation (only "[^<>]" or one level
        // of nested "<<...>>") cannot consume, so TrailerDict never matches on real PDFs and
        // TryPageCountViaCatalog always returns null there, silently falling back to the
        // max-/Count scan for every real document today.
        //
        // Both real fixtures happen to have exactly one /Type /Pages node, so the fallback
        // gives the same (correct) answer as the catalog chain would - that alone doesn't prove
        // the catalog path runs. This fixture instead has an /ID-bearing trailer (like the real
        // ones) AND two /Type /Pages nodes with DIFFERENT /Count values: the one the catalog
        // actually points to (/Count 3) and an unrelated orphan node with a higher count
        // (/Count 99) that no live Catalog references. The fallback max-scan cannot tell these
        // apart and would report 99; only genuinely resolving trailer -> /Root -> Catalog ->
        // /Pages yields the correct 3. A test that passed under both mechanisms would prove
        // nothing here.
        var pdf = System.Text.Encoding.Latin1.GetBytes(
            "%PDF-1.4\n" +
            "1 0 obj\n<< /Type /Pages /Kids [ 2 0 R 3 0 R 4 0 R ] /Count 3 >>\nendobj\n" +
            "5 0 obj\n<< /Type /Pages /Kids [ 6 0 R ] /Count 99 >>\nendobj\n" +
            "17 0 obj\n<< /Type /Catalog /Pages 1 0 R >>\nendobj\n" +
            "trailer\n<< /Size 18 /Root 17 0 R /ID [<BD81DB1E0BDC5C195FB28E64E5D337E9> " +
            "<BD81DB1E0BDC5C195FB28E64E5D337E9>] >>\n");

        Assert.Equal(3, PdfProbe.PageCount(pdf));
    }

    [Fact]
    public void IsPdf_ChecksTheHeaderMagic()
    {
        Assert.True(PdfProbe.IsPdf(System.Text.Encoding.Latin1.GetBytes("%PDF-1.4\n")));
        Assert.False(PdfProbe.IsPdf(new byte[] { 0x50, 0x4B, 0x03, 0x04 }));
    }

    [Fact]
    public void TextYPositions_ToleratesBothIntegerAndDecimalIdentityMatrixForms()
    {
        // Integer form is OfficeIMO's current output; decimal form is what other PDF writers
        // (or a future OfficeIMO version) could plausibly emit for the same identity scale.
        // Neither should be confused with an unrelated Tm operator (e.g. a rotated/scaled matrix).
        var pdf = System.Text.Encoding.Latin1.GetBytes(
            "%PDF-1.4\nBT\n1 0 0 1 72 685.64 Tm\n<41> Tj\n" +
            "1.000000 0.000000 0.000000 1.000000 72.00 700.00 Tm\n<42> Tj\n" +
            "2 0 0 2 72 500 Tm\n<43> Tj\nET\n");

        var positions = PdfProbe.TextYPositions(pdf);

        Assert.Equal(new[] { 685.64, 700.00 }, positions);
    }

    [Fact]
    public void TextYPositions_DoesNotFalseMatchInsideALargerLeadingNumber()
    {
        // Finding B: TextMatrix was unanchored, so "1 0 0 1 x y Tm" could match starting
        // partway through an unrelated larger number - e.g. inside "21 0 0 1 72 685.64 Tm" it
        // would match the substring "1 0 0 1 72 685.64 Tm" (starting at the second digit of
        // "21") and report a bogus Y position of 685.64 that doesn't correspond to any real
        // identity-scale text-positioning operator. TextYPositions is load-bearing for later
        // "nothing drawn off-page" assertions, so a spurious value here could mask a real defect.
        var pdf = System.Text.Encoding.Latin1.GetBytes(
            "%PDF-1.4\nBT\n21 0 0 1 72 685.64 Tm\n<41> Tj\n1 0 0 1 72 700 Tm\n<42> Tj\nET\n");

        var positions = PdfProbe.TextYPositions(pdf);

        Assert.Equal(new[] { 700.0 }, positions);
    }

    [Fact]
    public void TextYPositions_OnRealPdf_ReturnsNonEmptyListOfPositiveYCoordinates()
    {
        var positions = PdfProbe.TextYPositions(File.ReadAllBytes(ResultPdfPath));

        Assert.NotEmpty(positions);
        Assert.DoesNotContain(positions, y => y < 0);
    }
}
