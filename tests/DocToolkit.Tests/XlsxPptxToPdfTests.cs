using Xunit;

namespace DocToolkit.Tests;

/// <summary>
/// XLSX → PDF and PPTX → PDF.
///
/// Every input here is produced by <b>DocToolkit's own writers</b> — `WorkbookEditor` writes XLSX
/// through ClosedXML and `PptxDocumentWriter` writes PPTX itself, neither of which is OfficeIMO. A
/// renderer that only handled documents its own library produced would be useless in this package,
/// and nothing but this would notice.
///
/// Assertions go through <see cref="PdfProbe.MediaBoxes"/> rather than stopping at "the bytes start
/// with %PDF-". A renderer that emitted one blank page would pass the weaker check, which is exactly
/// the silent-success shape this repository keeps finding.
/// </summary>
public class XlsxPptxToPdfTests
{
    private static byte[] Workbook() => WorkbookEditor.Create("Sales", new object?[][]
    {
        new object?[] { "Region", "Total" },
        new object?[] { "North", 1200 },
        new object?[] { "South", 950 },
    });

    private static byte[] Deck(int slides) => PresentationEditor.Create(
        Enumerable.Range(1, slides).Select(i => PptxSlide.Titled($"Slide {i}", $"Body {i}")));

    // =====================================================================================
    // XLSX
    // =====================================================================================

    [Fact]
    public void XlsxConvert_ProducesARealPdfAtThePrintSize()
    {
        byte[] pdf = XlsxToPdfConverter.Convert(Workbook());

        Assert.True(PdfProbe.IsPdf(pdf));
        var boxes = PdfProbe.MediaBoxes(pdf);
        Assert.NotEmpty(boxes);
        Assert.All(boxes, box =>
        {
            Assert.Equal(612, box.Width, 0);
            Assert.Equal(792, box.Height, 0);
        });
    }

    [Fact]
    public void XlsxConvert_CarriesTheCellText()
    {
        string text = PdfProbe.ExtractText(XlsxToPdfConverter.Convert(Workbook()));

        Assert.Contains("Region", text, StringComparison.Ordinal);
        Assert.Contains("North", text, StringComparison.Ordinal);
    }

    [Fact]
    public void XlsxConvert_RejectsNullAndEmptyUnwrapped()
    {
        Assert.Throws<ArgumentNullException>(() => XlsxToPdfConverter.Convert(null!));
        Assert.Throws<ArgumentException>(() => XlsxToPdfConverter.Convert(Array.Empty<byte>()));
    }

    [Fact]
    public void XlsxConvert_WrapsAFailureInDocumentConversionException()
    {
        Assert.Throws<DocumentConversionException>(
            () => XlsxToPdfConverter.Convert(new byte[] { 1, 2, 3, 4 }));
    }

    [Fact]
    public async Task XlsxConvertAsync_WritesTheSamePdfToADestination()
    {
        using var source = new MemoryStream(Workbook());
        using var destination = new MemoryStream();

        await XlsxToPdfConverter.ConvertAsync(source, destination);

        Assert.True(PdfProbe.IsPdf(destination.ToArray()));
        Assert.NotEmpty(PdfProbe.MediaBoxes(destination.ToArray()));
    }

    [Fact]
    public void XlsxConvertFile_WritesThePdf()
    {
        string input = Path.Join(Path.GetTempPath(), $"{Guid.NewGuid():N}.xlsx");
        string output = Path.Join(Path.GetTempPath(), $"{Guid.NewGuid():N}.pdf");
        try
        {
            File.WriteAllBytes(input, Workbook());

            XlsxToPdfConverter.ConvertFile(input, output);

            Assert.True(PdfProbe.IsPdf(File.ReadAllBytes(output)));
        }
        finally
        {
            File.Delete(input);
            File.Delete(output);
        }
    }

    // =====================================================================================
    // PPTX
    // =====================================================================================

    // The page count is the assertion that catches a renderer returning success after emitting one
    // blank page. Three slides must be three pages.
    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public void PptxConvert_ProducesOnePagePerSlide(int slides)
    {
        byte[] pdf = PptxToPdfConverter.Convert(Deck(slides));

        Assert.True(PdfProbe.IsPdf(pdf));
        Assert.Equal(slides, PdfProbe.MediaBoxes(pdf).Count);
    }

    // 960 x 540 is the 16:9 slide geometry PptxDocumentWriter fixes, NOT a paper size. A PDF that
    // came back at 612 x 792 would mean the deck had been letterboxed onto US Letter.
    [Fact]
    public void PptxConvert_RendersAtTheSlideGeometryNotAPaperSize()
    {
        var boxes = PdfProbe.MediaBoxes(PptxToPdfConverter.Convert(Deck(2)));

        // A MediaBoxes returning nothing would satisfy Assert.All silently. The sibling test
        // above asserts the count for a different reason; this one needs its own, because a
        // broken probe and a correctly-sized deck are indistinguishable without it.
        Assert.Equal(2, boxes.Count);

        Assert.All(boxes, box =>
        {
            Assert.Equal(960, box.Width, 0);
            Assert.Equal(540, box.Height, 0);
        });
    }

    [Fact]
    public void PptxConvert_CarriesTheSlideText()
    {
        string text = PdfProbe.ExtractText(PptxToPdfConverter.Convert(Deck(2)));

        Assert.Contains("Slide 1", text, StringComparison.Ordinal);
        Assert.Contains("Slide 2", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PptxConvert_RejectsNullAndEmptyUnwrapped()
    {
        Assert.Throws<ArgumentNullException>(() => PptxToPdfConverter.Convert(null!));
        Assert.Throws<ArgumentException>(() => PptxToPdfConverter.Convert(Array.Empty<byte>()));
    }

    [Fact]
    public void PptxConvert_WrapsAFailureInDocumentConversionException()
    {
        Assert.Throws<DocumentConversionException>(
            () => PptxToPdfConverter.Convert(new byte[] { 1, 2, 3, 4 }));
    }

    [Fact]
    public async Task PptxConvertAsync_WritesThePdfToADestination()
    {
        using var source = new MemoryStream(Deck(2));
        using var destination = new MemoryStream();

        await PptxToPdfConverter.ConvertAsync(source, destination);

        Assert.Equal(2, PdfProbe.MediaBoxes(destination.ToArray()).Count);
    }

    [Fact]
    public void PptxConvertFile_WritesThePdf()
    {
        string input = Path.Join(Path.GetTempPath(), $"{Guid.NewGuid():N}.pptx");
        string output = Path.Join(Path.GetTempPath(), $"{Guid.NewGuid():N}.pdf");
        try
        {
            File.WriteAllBytes(input, Deck(2));

            PptxToPdfConverter.ConvertFile(input, output);

            Assert.Equal(2, PdfProbe.MediaBoxes(File.ReadAllBytes(output)).Count);
        }
        finally
        {
            File.Delete(input);
            File.Delete(output);
        }
    }

    // =====================================================================================
    // The offline guarantee, pinned at its source.
    //
    // AirGapGuardTests counts sockets, which is the behavioural proof. This asserts the POLICY
    // OBJECT the converters hand the renderer, because the two say different things: a socket count
    // of zero is also what you get from a document that happened to reference nothing, whereas this
    // fails the moment somebody constructs the options without the flags.
    // =====================================================================================

    [Fact]
    public void ResourcePolicy_RefusesRemoteAndLocalResources()
    {
        var (remote, local) = PdfRenderPolicy.DescribeForTests();

        Assert.False(remote, "AllowRemoteResourceResolution must be false - this package never fetches.");
        Assert.False(local, "AllowLocalFileAccess must be false - a document must not read the host's disk.");
    }

    /// <summary>
    /// EVERY options factory sets the policy, found by reflection rather than by a list.
    /// </summary>
    /// <remarks>
    /// <b>This asserts the EFFECTIVE policy, and deliberately no longer claims more than that.</b>
    /// It was written to catch a factory that forgets to set <c>ResourcePolicy</c>, and measured
    /// 2026-08-18 it does NOT: mutating both <c>ForDocument</c> and <c>ForWorkbook</c> to
    /// <c>new()</c> left every assertion here green.
    ///
    /// The reason is the one <c>PdfRenderPolicy</c>'s own comment gives - the upstream defaults
    /// already match, and the flags are stated anyway because a default is a policy somebody may
    /// revisit. <b>The consequence nobody had drawn is that no runtime test can tell the two
    /// apart.</b> Options built without the policy behave identically today, which is precisely
    /// what makes the omission survivable and invisible.
    ///
    /// So what this test is worth is the effective guarantee: whatever a factory returns, remote
    /// resolution and local file access are off. <b>Whether it was STATED rather than inherited is
    /// checked in source</b>, by <c>scripts/check-render-policy.py</c> - which does catch both
    /// mutants, and caught the real one: <c>DocxToPdfConverter</c> called a bare <c>ToPdf()</c> for
    /// as long as <c>PdfRenderPolicy</c> has existed.
    ///
    /// Derived rather than listed, so a fourth path is covered the day its factory appears.
    /// </remarks>
    /// <summary>
    /// There is deliberately no Word options factory, and this is what stops one coming back
    /// unmeasured.
    /// </summary>
    /// <remarks>
    /// <b>One was added on 2026-08-18 and reverted the same day.</b> Measured over 99 real
    /// documents, assigning a <c>ResourcePolicy</c> to <c>WordPdfSaveOptions</c> dropped DOCX to PDF
    /// from <b>71/99 to 57/99</b>, and the failures were text-encoding preflight errors rather than
    /// anything about resources - it puts the Word renderer into a mode that cannot find fonts.
    ///
    /// <b>The flag values are not the cause</b>, which is the part that makes this worth pinning:
    /// both flags <c>true</c>, both <c>false</c>, and a default-constructed policy all give 57/99,
    /// while no policy at all and an empty options object both give 71/99. Assigning the object is
    /// the change.
    ///
    /// <b>This asserts a DECISION, not a behaviour</b>, and deliberately so. The behaviour is
    /// host-dependent - which glyphs are encodable depends on the fonts a machine has, which this
    /// repository already records as varying PDF size a hundredfold - so an assertion on rendering
    /// would be a flake. What can be asserted is that nobody has quietly reintroduced the factory.
    /// If this test fails, re-run the table above before deleting it.
    /// </remarks>
    [Fact]
    public void ThereIsNoWordOptionsFactory_AndThatIsMeasured()
    {
        var wordFactories = typeof(PdfRenderPolicy)
            .GetMethods(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static)
            .Where(m => m.ReturnType.Name.Contains("Word", StringComparison.Ordinal))
            .Select(m => m.Name)
            .ToList();

        Assert.True(wordFactories.Count == 0,
            "PdfRenderPolicy has gained a Word options factory: " + string.Join(", ", wordFactories)
            + ". Assigning a ResourcePolicy to WordPdfSaveOptions was measured to drop DOCX to PDF "
            + "from 71/99 to 57/99 on real documents, with font-encoding failures. Re-run that "
            + "measurement before keeping it.");
    }

    [Fact]
    public void EveryRenderPolicyFactorySetsThePolicy()
    {
        var factories = typeof(PdfRenderPolicy)
            .GetMethods(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static)
            .Where(m => m.GetParameters().Length == 0
                     && m.ReturnType.Name.EndsWith("PdfSaveOptions", StringComparison.Ordinal))
            .ToList();

        // Non-vacuity: reflection that matched nothing would pass while asserting nothing at all.
        //
        // TWO, not three. The Word path deliberately has no factory - assigning a ResourcePolicy to
        // WordPdfSaveOptions was measured to drop DOCX to PDF from 71/99 to 57/99 on real documents,
        // for font reasons rather than resource ones. ThereIsNoWordOptionsFactory_AndThatIsMeasured
        // pins that; this floor only has to stop the reflection silently matching nothing.
        Assert.True(factories.Count >= 2,
            $"expected a factory for the XLSX and PPTX paths, found {factories.Count} - renamed?");

        foreach (var factory in factories)
        {
            var options = factory.Invoke(null, null);
            var policy = options!.GetType().GetProperty("ResourcePolicy")!.GetValue(options);

            Assert.True(policy is not null, $"{factory.Name} left ResourcePolicy null");

            var remoteFlag = (bool)policy!.GetType().GetProperty("AllowRemoteResourceResolution")!.GetValue(policy)!;
            var localFlag = (bool)policy.GetType().GetProperty("AllowLocalFileAccess")!.GetValue(policy)!;

            Assert.False(remoteFlag, $"{factory.Name} allows remote resource resolution");
            Assert.False(localFlag, $"{factory.Name} allows local file access");
        }
    }
}
