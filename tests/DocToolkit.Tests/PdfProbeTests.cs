using Xunit;

namespace DocToolkit.Tests;

public class PdfProbeTests
{
    [Fact]
    public void ExtractText_DecodesHexStringTextOperators()
    {
        // "Acme" == 41 63 6D 65 ; "Corp" == 43 6F 72 70
        var pdf = System.Text.Encoding.Latin1.GetBytes(
            "%PDF-1.4\n5 0 obj\n<< /Length 40 >>\nstream\nBT\n<41636D65> Tj\n<436F7270> Tj\nET\nendstream\nendobj\n");

        Assert.Equal("AcmeCorp", PdfProbe.ExtractText(pdf));
    }

    [Fact]
    public void PageCount_ReadsTheCountFromThePageTree()
    {
        var pdf = System.Text.Encoding.Latin1.GetBytes(
            "%PDF-1.4\n1 0 obj\n<< /Type /Pages /Count 7 /Kids [ 7 0 R ] >>\nendobj\n");

        Assert.Equal(7, PdfProbe.PageCount(pdf));
    }

    [Fact]
    public void IsPdf_ChecksTheHeaderMagic()
    {
        Assert.True(PdfProbe.IsPdf(System.Text.Encoding.Latin1.GetBytes("%PDF-1.4\n")));
        Assert.False(PdfProbe.IsPdf(new byte[] { 0x50, 0x4B, 0x03, 0x04 }));
    }
}
