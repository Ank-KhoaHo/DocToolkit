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
    public void TextYPositions_OnRealPdf_ReturnsNonEmptyListOfPositiveYCoordinates()
    {
        var positions = PdfProbe.TextYPositions(File.ReadAllBytes(ResultPdfPath));

        Assert.NotEmpty(positions);
        Assert.DoesNotContain(positions, y => y < 0);
    }
}
