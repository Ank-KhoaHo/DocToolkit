using System.Text;
using DocToolkit;

namespace DocToolkit.Tests;

/// <summary>
/// Legacy .doc import.
///
/// <b>The two fixtures are a matched pair and the suite is only honest because of it.</b>
/// <c>legacy.doc</c> holds a table, which makes Word emit a binary Data stream that a .docx cannot
/// carry, so it is the case that must be REFUSED without an opt-in. <c>legacy-lossless.doc</c> is
/// one line of text with no such stream, so it is the case that must be ALLOWED without one.
/// Either assertion alone is vacuous: a converter that refused everything would pass the first, and
/// one that refused nothing would pass the second.
/// </summary>
public class DocToDocxConverterTests
{
    private static readonly string BlockingPath =
        Path.Join(AppContext.BaseDirectory, "assets", "legacy.doc");

    private static readonly string LosslessPath =
        Path.Join(AppContext.BaseDirectory, "assets", "legacy-lossless.doc");

    private static byte[] Blocking() => File.ReadAllBytes(BlockingPath);
    private static byte[] Lossless() => File.ReadAllBytes(LosslessPath);

    private static readonly LegacyDocOptions Allow = new() { AllowContentLoss = true };

    // ---- the fixtures are what they claim to be -------------------------------------------

    [Fact]
    public void TheFixtures_AreRealCompoundBinaryFiles_NotRenamedDocxPackages()
    {
        // Guards the suite itself: if someone regenerates a fixture as a .docx with a .doc name,
        // every test below would be exercising the wrong format while still passing.
        foreach (var bytes in new[] { Blocking(), Lossless() })
        {
            Assert.Equal(new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 }, bytes.Take(8));
            Assert.NotEqual(new byte[] { 0x50, 0x4B }, bytes.Take(2));   // not a ZIP, i.e. not .docx
        }
    }

    // ---- reading -------------------------------------------------------------------------

    [Fact]
    public void ExtractText_ReadsBodyText_FromARealLegacyDocument()
    {
        var text = DocToDocxConverter.ExtractText(Blocking());

        // Literals, not "is not empty": an assertion that only checks for content passes against
        // any document at all.
        Assert.Contains("Legacy Heading", text, StringComparison.Ordinal);
        Assert.Contains("First paragraph of a Word 97-2003 binary file.", text, StringComparison.Ordinal);
        Assert.Contains("Second paragraph SENTINEL-TWO.", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ExtractText_ReadsTableCells_WhichAreNotInTheParagraphList()
    {
        // Table cells live outside WordDocument.Paragraphs. Reading only paragraphs would silently
        // drop every table in the document and still look like it worked.
        var text = DocToDocxConverter.ExtractText(Blocking());

        Assert.Contains("R1C1", text, StringComparison.Ordinal);
        Assert.Contains("R1C2", text, StringComparison.Ordinal);
        Assert.Contains("R2C1", text, StringComparison.Ordinal);
        Assert.Contains("R2C2", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ExtractText_SeparatesBlocks_SoAdjacentTextDoesNotFuse()
    {
        var text = DocToDocxConverter.ExtractText(Blocking());

        // The defect this repository shipped for eight releases was exactly this fusion, so the
        // assertion is on the literal separator rather than on the words being present.
        Assert.DoesNotContain("fileSecond", text, StringComparison.Ordinal);
        Assert.Contains("R1C1\tR1C2", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ExtractText_DoesNotRequireTheOptIn_EvenForADocumentThatCannotConvert()
    {
        // The whole justification for ExtractText taking no options: reading is not conversion.
        // If this ever starts throwing, the reading path has picked up the save path's policy.
        var text = DocToDocxConverter.ExtractText(Blocking());
        Assert.Contains("SENTINEL-TWO", text, StringComparison.Ordinal);
    }

    // ---- the loss policy: both directions --------------------------------------------------

    [Fact]
    public void Convert_Refuses_WhenTheSourceHoldsContentADocxCannotCarry()
    {
        var ex = Assert.Throws<DocumentConversionException>(() => DocToDocxConverter.Convert(Blocking()));

        // The message must name DocToolkit's own option. Naming WordSaveOptions.LossPolicy would
        // be advice the caller cannot act on - the D18 failure shape.
        Assert.Contains("AllowContentLoss", ex.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("LossPolicy", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Convert_Succeeds_WithoutAnyOptIn_WhenThereIsNothingToLose()
    {
        // The positive control for the test above. Without this, "it refused" would pass against a
        // converter that refuses unconditionally.
        var docx = DocToDocxConverter.Convert(Lossless());

        Assert.NotEmpty(docx);
        Assert.Equal(new byte[] { 0x50, 0x4B }, docx.Take(2));   // a real ZIP, i.e. a .docx package
        Assert.Contains("PLAIN-SENTINEL", DocxEditor.ExtractText(docx), StringComparison.Ordinal);
    }

    [Fact]
    public void Convert_Succeeds_WhenTheLossIsAcceptedDeliberately()
    {
        var docx = DocToDocxConverter.Convert(Blocking(), Allow);

        Assert.Equal(new byte[] { 0x50, 0x4B }, docx.Take(2));
    }

    [Fact]
    public void Convert_KeepsTextTablesAndFormatting_WhenTheLossIsAccepted()
    {
        // "It saved" is not the assertion. Opting into loss must lose only the binary payload -
        // if it quietly produced an empty or text-only document, every other test here would still
        // pass.
        var docx = DocToDocxConverter.Convert(Blocking(), Allow);
        var text = DocxEditor.ExtractText(docx);

        Assert.Contains("Legacy Heading", text, StringComparison.Ordinal);
        Assert.Contains("SENTINEL-TWO", text, StringComparison.Ordinal);
        Assert.Contains("R1C1", text, StringComparison.Ordinal);
        Assert.Contains("R2C2", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Convert_WithAllowContentLossFalse_IsTheSameAsNotPassingOptionsAtAll()
    {
        // The default must be the safe one however it is spelled.
        Assert.Throws<DocumentConversionException>(
            () => DocToDocxConverter.Convert(Blocking(), new LegacyDocOptions { AllowContentLoss = false }));
        Assert.Throws<DocumentConversionException>(
            () => DocToDocxConverter.Convert(Blocking(), options: null));
    }

    // ---- the report ------------------------------------------------------------------------

    [Fact]
    public void ConvertWithReport_NamesTheBinaryPayloadThatWasDropped()
    {
        var result = DocToDocxConverter.ConvertWithReport(Blocking(), Allow);

        var omission = Assert.Single(result.Warnings, w => w.Kind == ConversionLossKind.Omission);
        Assert.Equal("DOC-BINARY-DATA-STREAM-PRESENT", omission.Code);
        Assert.Contains("bytes", omission.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ConvertWithReport_ReturnsTheSameBytesAsConvert()
    {
        // Two entry points, one conversion. If they diverge, a caller who switched overloads to
        // read the report would silently get a different document.
        var plain = DocToDocxConverter.Convert(Blocking(), Allow);
        var reported = DocToDocxConverter.ConvertWithReport(Blocking(), Allow);

        Assert.Equal(DocxEditor.ExtractText(plain), DocxEditor.ExtractText(reported.Value));
    }

    [Fact]
    public void ConvertWithReport_ReportsNoOmission_ForADocumentThatLosesNothing()
    {
        // Discriminates: without this, the omission assertion above could be satisfied by a
        // converter that reports the same warning for every input.
        var result = DocToDocxConverter.ConvertWithReport(Lossless());

        Assert.DoesNotContain(result.Warnings, w => w.Kind == ConversionLossKind.Omission);
    }

    // ---- refusing what is not a .doc --------------------------------------------------------

    [Fact]
    public void Convert_RefusesADocxPackage_AndSaysWhichApiToUseInstead()
    {
        // Handing a .docx to the legacy converter is the likely mistake, so the message has to
        // answer it rather than say "invalid".
        var docx = DocxEditor.Create(new[] { DocxBlock.Paragraph("not a legacy document") });

        var ex = Assert.Throws<DocumentConversionException>(() => DocToDocxConverter.Convert(docx));
        Assert.Contains("DocxEditor", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Convert_RefusesArbitraryBytes()
    {
        var junk = Encoding.UTF8.GetBytes(new string('x', 4096));
        Assert.Throws<DocumentConversionException>(() => DocToDocxConverter.Convert(junk));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void EveryByteArrayEntryPoint_RejectsNullAndEmpty(bool extract)
    {
        if (extract)
        {
            Assert.Throws<ArgumentNullException>(() => DocToDocxConverter.ExtractText(null!));
            Assert.Throws<ArgumentException>(() => DocToDocxConverter.ExtractText(Array.Empty<byte>()));
        }
        else
        {
            Assert.Throws<ArgumentNullException>(() => DocToDocxConverter.Convert(null!));
            Assert.Throws<ArgumentException>(() => DocToDocxConverter.Convert(Array.Empty<byte>()));
        }
    }

    // ---- the Stream overloads ---------------------------------------------------------------

    [Fact]
    public async Task ConvertAsync_WritesTheSameDocumentTheByteArrayOverloadProduces()
    {
        using var source = new MemoryStream(Blocking(), writable: false);
        using var destination = new MemoryStream();

        await DocToDocxConverter.ConvertAsync(source, destination, Allow);

        Assert.Equal(
            DocxEditor.ExtractText(DocToDocxConverter.Convert(Blocking(), Allow)),
            DocxEditor.ExtractText(destination.ToArray()));
    }

    [Fact]
    public async Task ConvertAsync_Refuses_WithoutTheOptIn()
    {
        using var source = new MemoryStream(Blocking(), writable: false);
        using var destination = new MemoryStream();

        await Assert.ThrowsAsync<DocumentConversionException>(
            () => DocToDocxConverter.ConvertAsync(source, destination));

        // Nothing may be written on the refusing path - a half-written destination is worse than
        // an exception, because it looks like a document.
        Assert.Equal(0, destination.Length);
    }

    [Fact]
    public async Task ExtractTextAsync_ReadsAForwardOnlySource()
    {
        using var source = new ForwardOnlySource(Blocking());

        var text = await DocToDocxConverter.ExtractTextAsync(source);

        Assert.Contains("SENTINEL-TWO", text, StringComparison.Ordinal);
    }

    [Fact]
    public async Task EveryStreamEntryPoint_ThrowsForAnAlreadyCancelledToken()
    {
        using var cts = new CancellationTokenSource();
        await cts.CancelAsync();

        using var source = new MemoryStream(Blocking(), writable: false);
        using var destination = new MemoryStream();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(
            () => DocToDocxConverter.ConvertAsync(source, destination, null, cts.Token));
        await Assert.ThrowsAnyAsync<OperationCanceledException>(
            () => DocToDocxConverter.ExtractTextAsync(source, cts.Token));

        // The token must be observed BEFORE the work, not noticed by a later write. Seven
        // PdfEditor overloads passed a cancellation test for that wrong reason.
        Assert.Equal(0, destination.Length);
    }

    [Fact]
    public async Task StreamOverloads_RejectUnusableStreams()
    {
        using var readable = new MemoryStream(Lossless(), writable: false);
        using var writable = new MemoryStream();

        await Assert.ThrowsAsync<ArgumentNullException>(
            () => DocToDocxConverter.ConvertAsync(null!, writable));
        await Assert.ThrowsAsync<ArgumentNullException>(
            () => DocToDocxConverter.ConvertAsync(readable, null!));
        using var empty = new MemoryStream();
        await Assert.ThrowsAsync<ArgumentException>(
            () => DocToDocxConverter.ConvertAsync(empty, writable));   // empty source
    }
}
