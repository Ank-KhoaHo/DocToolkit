namespace DocToolkit.Tests;

/// <summary>
/// Legacy binary Office input — the pre-2007 formats — on the PDF converters.
///
/// <b>The two formats are deliberately treated differently, and the difference is measured rather
/// than assumed.</b> Both were reaching the PDF converters by accident: the renderer underneath
/// reads binary Office files, so `.ppt` and `.xls` "worked" on those paths while every editor
/// refused them, documented nowhere.
///
/// <list type="bullet">
/// <item><description>
/// <b>`.ppt` is claimed.</b> Measured 2026-08-17: a 16.8 MB deck converted in 858 ms, and the
/// slowest of eight real files was 1.7 s. That is an ordinary cost for an ordinary capability.
/// </description></item>
/// <item><description>
/// <b>`.xls` is refused.</b> A 101 KB workbook took 10.9 s, a 2.3 MB one did not finish in ten
/// minutes, and a 7.7 MB one spent 161 s before failing anyway — while the supported `.xlsx` path
/// renders 20,000 rows in 3.7 s. Unbounded work on caller-chosen input, through a path that never
/// claimed the format.
/// </description></item>
/// </list>
/// </summary>
public class LegacyBinaryOfficeTests
{
    private static byte[] Fixture(string name) =>
        File.ReadAllBytes(Path.Join(AppContext.BaseDirectory, "assets", name));

    private static bool IsCompoundFile(byte[] b) =>
        b.Length > 8 && b[0] == 0xD0 && b[1] == 0xCF && b[2] == 0x11 && b[3] == 0xE0;

    [Theory]
    [InlineData("legacy.ppt")]
    [InlineData("legacy.xls")]
    public void TheFixtures_AreRealBinaryOfficeFiles(string name)
    {
        // Guards the premise: both tests below are about the compound-file format specifically, and
        // a fixture regenerated as OOXML would make them test nothing.
        Assert.True(IsCompoundFile(Fixture(name)));
    }

    // ---- .ppt: claimed ---------------------------------------------------------------------------

    [Fact]
    public void ALegacyPpt_RendersToPdf_WithItsTextIntact()
    {
        var pdf = PptxToPdfConverter.Convert(Fixture("legacy.ppt"));

        Assert.True(PdfProbe.IsPdf(pdf));
        // The text, not merely that bytes came back: a renderer that produced an empty deck would
        // still return a plausible PDF.
        Assert.Contains("PPT-SENTINEL", PdfProbe.ExtractText(pdf), StringComparison.Ordinal);
    }

    [Fact]
    public void ALegacyPpt_IsStillRefusedByThePresentationEditor()
    {
        // The asymmetry is intentional and worth pinning, because it looks like an inconsistency
        // until you know why: the editors are OOXML-only, and the PDF path goes through a renderer
        // that reads the binary format. Claiming .ppt on the converter does NOT claim it everywhere.
        Assert.Throws<DocumentConversionException>(
            () => PresentationEditor.SlideCount(Fixture("legacy.ppt")));
    }

    // ---- .xls: refused, and quickly ---------------------------------------------------------------

    [Fact]
    public void ALegacyXls_IsRefusedByTheePdfConverter_AndToldWhatToDo()
    {
        var ex = Assert.Throws<DocumentConversionException>(
            () => XlsxToPdfConverter.Convert(Fixture("legacy.xls")));

        Assert.Contains("legacy Excel 97-2003", ex.Message, StringComparison.Ordinal);
        Assert.Contains("save it as .xlsx", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ALegacyXls_IsRefusedIMMEDIATELY_NotAfterMinutesOfWork()
    {
        // The point of the refusal is the COST, so the test asserts the cost. Before this check the
        // same call spent 10.9 s on a 101 KB file and over ten minutes on a 2.3 MB one; the fixture
        // here is 25 KB, so any figure in seconds would mean the signature check was not reached.
        //
        // A generous ceiling on purpose: this runs on shared CI runners where wall-clock assertions
        // have flapped before, and it only needs to separate "immediate" from "minutes".
        var xls = Fixture("legacy.xls");
        var sw = System.Diagnostics.Stopwatch.StartNew();
        Assert.Throws<DocumentConversionException>(() => XlsxToPdfConverter.Convert(xls));
        sw.Stop();

        Assert.True(sw.ElapsedMilliseconds < 2000,
            $"refusal took {sw.ElapsedMilliseconds} ms - the signature check is not being reached, "
            + "and the converter is doing real work on input it does not support.");
    }

    [Fact]
    public void ARealXlsx_StillRenders()
    {
        // The control. A check that refused every workbook would satisfy every assertion above.
        var xlsx = WorkbookEditor.Create("Sales", new[] { new object?[] { "XLSX-SENTINEL", 1 } });

        Assert.True(PdfProbe.IsPdf(XlsxToPdfConverter.Convert(xlsx)));
    }

    [Fact]
    public void ARealPptx_StillRenders()
    {
        var pptx = PresentationEditor.Create(new[] { PptxSlide.Titled("PPTX-SENTINEL") });

        Assert.True(PdfProbe.IsPdf(PptxToPdfConverter.Convert(pptx)));
    }
}
