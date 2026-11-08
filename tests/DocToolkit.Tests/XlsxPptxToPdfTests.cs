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
    public void XlsxConvert_RendersTheEvaluatedFormulaValueRatherThanItsSourceText()
    {
        // A formula cell this package writes carries no cached value at all - see XlsxFormula's
        // own remarks - so a renderer that reads the raw file rather than going through
        // ClosedXML's lazy in-memory evaluation (the way ReadCell/ReadSheet do) shows the
        // formula's SOURCE TEXT where a value belongs. Measured 2026-08-31: this converter did
        // exactly that before the fix.
        var xlsx = WorkbookEditor.Create("S", new object?[][]
        {
            new object?[] { 2, 40, XlsxFormula.From("=SUM(A1:B1)") },
        });

        string text = PdfProbe.ExtractText(XlsxToPdfConverter.Convert(xlsx));

        Assert.Contains("42", text, StringComparison.Ordinal);
        Assert.DoesNotContain("SUM", text, StringComparison.Ordinal);
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
    public async Task XlsxConvertAsync_AlsoRendersTheEvaluatedFormulaValue()
    {
        // The Stream overload has its own Calculate() call, a separate line from Convert(byte[])'s
        // - this proves that call specifically, rather than trusting the two overloads agree.
        var xlsx = WorkbookEditor.Create("S", new object?[][]
        {
            new object?[] { 2, 40, XlsxFormula.From("=SUM(A1:B1)") },
        });
        using var source = new MemoryStream(xlsx);
        using var destination = new MemoryStream();

        await XlsxToPdfConverter.ConvertAsync(source, destination);

        string text = PdfProbe.ExtractText(destination.ToArray());
        Assert.Contains("42", text, StringComparison.Ordinal);
        Assert.DoesNotContain("SUM", text, StringComparison.Ordinal);
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
    /// There is deliberately no Word options factory, and this is what stops one coming back
    /// unmeasured.
    /// </summary>
    /// <remarks>
    /// <b>One was added on 2026-08-18 and reverted the same day.</b> Measured over 99 real
    /// documents, assigning a <c>ResourcePolicy</c> to <c>WordToPdfOptions</c> dropped DOCX to PDF
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
            + ". Assigning a ResourcePolicy to WordToPdfOptions was measured to drop DOCX to PDF "
            + "from 71/99 to 57/99 on real documents, with font-encoding failures. Re-run that "
            + "measurement before keeping it.");
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
    [Fact]
    public void EveryRenderPolicyFactorySetsThePolicy()
    {
        var factories = typeof(PdfRenderPolicy)
            .GetMethods(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static)
            .Where(m => m.GetParameters().Length == 0
                     && m.ReturnType.Name.EndsWith("ToPdfOptions", StringComparison.Ordinal))
            .ToList();

        // Non-vacuity: reflection that matched nothing would pass while asserting nothing at all.
        //
        // TWO, not three. The Word path deliberately has no factory - assigning a ResourcePolicy to
        // WordToPdfOptions was measured to drop DOCX to PDF from 71/99 to 57/99 on real documents,
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

    /// <summary>
    /// The font-taking factories keep the same policy, and hand the fonts over with shaping on.
    /// </summary>
    /// <remarks>
    /// The test above only reaches parameterless factories, so without this one a font overload
    /// could drop the policy and nothing would notice - the same silent shape that test exists for.
    /// </remarks>
    [Fact]
    public void EveryFontTakingRenderPolicyFactoryKeepsThePolicyAndShapes()
    {
        var factories = typeof(PdfRenderPolicy)
            .GetMethods(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static)
            .Where(m => m.GetParameters() is [{ ParameterType: var p }] && p == typeof(PdfFontOptions)
                     && m.ReturnType.Name.EndsWith("ToPdfOptions", StringComparison.Ordinal))
            .ToList();

        Assert.True(factories.Count >= 2,
            $"expected a font-taking factory for the XLSX and PPTX paths, found {factories.Count}");

        foreach (var factory in factories)
        {
            var options = factory.Invoke(null, [new PdfFontOptions("First", [1, 2, 3])]);
            var policy = options!.GetType().GetProperty("ResourcePolicy")!.GetValue(options);
            var pdf = (OfficeIMO.Pdf.PdfOptions?)options.GetType().GetProperty("PdfOptions")!.GetValue(options);

            Assert.True(policy is not null, $"{factory.Name}(fonts) left ResourcePolicy null");
            Assert.False((bool)policy!.GetType().GetProperty("AllowRemoteResourceResolution")!.GetValue(policy)!);
            Assert.False((bool)policy.GetType().GetProperty("AllowLocalFileAccess")!.GetValue(policy)!);
            Assert.Equal(["First"], pdf!.EmbeddedFontFallbacks!.Candidates.Select(c => c.FontName));
            Assert.NotNull(pdf.TextShapingProvider);
        }
    }

    /// <summary>
    /// Null fonts build exactly the options the parameterless factories build - no fallbacks, no
    /// shaping.
    /// </summary>
    /// <remarks>
    /// <b>Asserted on the options, because a render comparison cannot see it.</b> The converters'
    /// old overloads forward to the new ones with <see langword="null"/>, so comparing
    /// <c>Convert(x)</c> with <c>Convert(x, null)</c> compares a call with itself.
    /// </remarks>
    [Fact]
    public void NullFonts_BuildTheSameOptionsAsTheParameterlessFactories()
    {
        static (object? Fallbacks, object? Shaping) Of(OfficeIMO.Pdf.PdfOptions? o) =>
            (o?.EmbeddedFontFallbacks, o?.TextShapingProvider);

        Assert.Equal(Of(PdfRenderPolicy.ForWorkbook().PdfOptions), Of(PdfRenderPolicy.ForWorkbook(null).PdfOptions));
        Assert.Equal(Of(PdfRenderPolicy.ForPresentation().PdfOptions), Of(PdfRenderPolicy.ForPresentation(null).PdfOptions));
        Assert.Equal((null, null), Of(PdfRenderPolicy.ForWorkbook(null).PdfOptions));
        Assert.Equal((null, null), Of(PdfRenderPolicy.ForPresentation(null).PdfOptions));
    }

    // =====================================================================================
    // A122: fonts, and right-to-left shaping, on XLSX -> PDF
    // =====================================================================================

    private static byte[] SheetWith(string text) =>
        WorkbookEditor.Create("Sheet", new object?[][] { new object?[] { text } });

    /// <summary>
    /// A supplied font, or null on a host without one - in which case the calling test returns having
    /// asserted nothing, unless <c>DOCTOOLKIT_REQUIRE_FONT=1</c> makes that a failure. See
    /// <see cref="PdfFontOptionsTests.FindFont"/>.
    /// </summary>
    private static PdfFontOptions? HostFonts() =>
        PdfFontOptionsTests.FindFont() is { } font ? new PdfFontOptions("Supplied", font) : null;

    /// <summary>
    /// Like <see cref="HostFonts"/>, but a font whose cmap covers the Arabic presentation forms: the
    /// joining tests need glyphs to draw (see <see cref="ArabicFontTests"/>).
    /// </summary>
    private static PdfFontOptions? ArabicHostFonts() =>
        ArabicFontTests.FindArabicFont() is { } font ? new PdfFontOptions("Supplied", font) : null;

    /// <summary>The host-independent proof the bytes reach the renderer: its own refusal.</summary>
    [Theory]
    [InlineData(3)]
    [InlineData(1024)]
    public void XlsxConvert_WithInvalidFontBytes_FailsWithTheRenderersOwnMessage(int size)
    {
        var ex = Assert.Throws<DocumentConversionException>(
            () => XlsxToPdfConverter.Convert(SheetWith("Plain Latin text."), new PdfFontOptions("Fake", new byte[size])));

        Assert.Contains("TrueType", ex.InnerException!.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void XlsxConvert_WithARealFont_HebrewIsInReadingOrder()
    {
        if (HostFonts() is not { } fonts)
            return;

        PdfFontOptionsTests.AssertHebrewLineIsInReadingOrder(
            XlsxToPdfConverter.Convert(SheetWith(string.Join(" ", PdfFontOptionsTests.HebrewLine)), fonts));
    }

    [Theory]
    [InlineData("فاتورة ضريبية رقم 12")]
    public void XlsxConvert_WithARealFont_RightToLeftTextIsShaped(string text)
    {
        if (HostFonts() is not { } fonts)
            return;
        var xlsx = SheetWith(text);

        Assert.NotEqual(UnshapedRender.Xlsx(xlsx, fonts), XlsxToPdfConverter.Convert(xlsx, fonts));
    }

    [Fact]
    public void XlsxConvert_WithARealFont_LatinIsUnchangedByShaping()
    {
        if (HostFonts() is not { } fonts)
            return;
        var xlsx = SheetWith("Invoice total, office fi fl ffi: 60,000");

        Assert.Equal(UnshapedRender.Xlsx(xlsx, fonts), XlsxToPdfConverter.Convert(xlsx, fonts));
    }

    [Fact]
    public void XlsxConvert_WithARealFont_MixedLatinAndHebrewInOneCellRenders()
    {
        if (HostFonts() is not { } fonts)
            return;

        var pdf = XlsxToPdfConverter.Convert(SheetWith("Invoice חשבונית 12"), fonts);

        Assert.Contains("Invoice", PdfProbe.ExtractText(pdf), StringComparison.Ordinal);
    }

    /// <summary>Persian text, with its own letters, is joined (A124) and must render, not throw.</summary>
    [Fact]
    public void XlsxConvert_WithARealFont_PersianRendersWithoutThrowing()
    {
        if (HostFonts() is not { } fonts)
            return;

        Assert.True(PdfProbe.IsPdf(XlsxToPdfConverter.Convert(SheetWith("صورتحساب فروش کالا"), fonts)));
    }

    /// <summary>
    /// Supplied fonts replace the host's fallbacks, so text they do not cover is refused on every
    /// host, naming the character - never rendered blank.
    /// </summary>
    [Fact]
    public void XlsxConvert_WithFontsMissingTheScript_NamesTheCharacter()
    {
        if (HostFonts() is not { } fonts)
            return;

        var ex = Assert.Throws<DocumentConversionException>(
            () => XlsxToPdfConverter.Convert(SheetWith("发票"), fonts));

        Assert.Contains("U+53D1", ex.InnerException!.Message, StringComparison.Ordinal);
    }

    [Fact]
    public async Task XlsxConvertAsync_WithFonts_MatchesTheArrayOverload()
    {
        if (HostFonts() is not { } fonts)
            return;
        var xlsx = SheetWith("חשבונית מס מספר 12");

        using var destination = new MemoryStream();
        await XlsxToPdfConverter.ConvertAsync(new MemoryStream(xlsx), destination, fonts);

        Assert.Equal(XlsxToPdfConverter.Convert(xlsx, fonts), destination.ToArray());
    }

    // =====================================================================================
    // A122: fonts, and right-to-left shaping, on PPTX -> PDF
    // =====================================================================================

    private static byte[] SlideWith(string text) =>
        PresentationEditor.Create([PptxSlide.Titled(text, text)]);

    [Theory]
    [InlineData(3)]
    [InlineData(1024)]
    public void PptxConvert_WithInvalidFontBytes_FailsWithTheRenderersOwnMessage(int size)
    {
        var ex = Assert.Throws<DocumentConversionException>(
            () => PptxToPdfConverter.Convert(SlideWith("Plain Latin text."), new PdfFontOptions("Fake", new byte[size])));

        Assert.Contains("TrueType", ex.InnerException!.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void PptxConvert_WithARealFont_HebrewIsInReadingOrder()
    {
        if (HostFonts() is not { } fonts)
            return;

        PdfFontOptionsTests.AssertHebrewLineIsInReadingOrder(
            PptxToPdfConverter.Convert(
                PresentationEditor.Create([PptxSlide.Titled(string.Join(" ", PdfFontOptionsTests.HebrewLine))]), fonts));
    }

    [Theory]
    [InlineData("فاتورة ضريبية رقم 12")]
    public void PptxConvert_WithARealFont_RightToLeftTextIsShaped(string text)
    {
        if (HostFonts() is not { } fonts)
            return;
        var pptx = SlideWith(text);

        Assert.NotEqual(UnshapedRender.Pptx(pptx, fonts), PptxToPdfConverter.Convert(pptx, fonts));
    }

    [Fact]
    public void PptxConvert_WithARealFont_LatinIsUnchangedByShaping()
    {
        if (HostFonts() is not { } fonts)
            return;
        var pptx = SlideWith("Invoice total, office fi fl ffi: 60,000");

        Assert.Equal(UnshapedRender.Pptx(pptx, fonts), PptxToPdfConverter.Convert(pptx, fonts));
    }

    [Fact]
    public void PptxConvert_WithARealFont_MixedLatinAndHebrewRenders()
    {
        if (HostFonts() is not { } fonts)
            return;

        var pdf = PptxToPdfConverter.Convert(SlideWith("Invoice חשבונית 12"), fonts);

        Assert.Contains("Invoice", PdfProbe.ExtractText(pdf), StringComparison.Ordinal);
    }

    [Fact]
    public void PptxConvert_WithARealFont_PersianRendersWithoutThrowing()
    {
        if (HostFonts() is not { } fonts)
            return;

        Assert.True(PdfProbe.IsPdf(PptxToPdfConverter.Convert(SlideWith("صورتحساب فروش کالا"), fonts)));
    }

    [Fact]
    public void PptxConvert_WithFontsMissingTheScript_NamesTheCharacter()
    {
        if (HostFonts() is not { } fonts)
            return;

        var ex = Assert.Throws<DocumentConversionException>(
            () => PptxToPdfConverter.Convert(SlideWith("发票"), fonts));

        Assert.Contains("U+53D1", ex.InnerException!.Message, StringComparison.Ordinal);
    }

    [Fact]
    public async Task PptxConvertAsync_WithFonts_MatchesTheArrayOverload()
    {
        if (HostFonts() is not { } fonts)
            return;
        var pptx = SlideWith("חשבונית מס מספר 12");

        using var destination = new MemoryStream();
        await PptxToPdfConverter.ConvertAsync(new MemoryStream(pptx), destination, fonts);

        Assert.Equal(PptxToPdfConverter.Convert(pptx, fonts), destination.ToArray());
    }

    // =====================================================================================
    // A124 - Arabic letters join when fonts are supplied
    // =====================================================================================

    [Fact]
    public void XlsxConvert_WithARealFont_ArabicIsPreShaped()
    {
        if (ArabicFontTests.FindArabicFont() is not { } font)
            return;
        var fonts = new PdfFontOptions("Supplied", font);
        var xlsx = SheetWith(PdfFontOptionsTests.ArabicLine);

        PdfFontOptionsTests.AssertArabicLineIsJoined(XlsxToPdfConverter.Convert(xlsx, fonts), font);
    }

    [Fact]
    public async Task XlsxConvertAsync_WithARealFont_ArabicIsPreShaped()
    {
        if (ArabicFontTests.FindArabicFont() is not { } font)
            return;
        var fonts = new PdfFontOptions("Supplied", font);
        var xlsx = SheetWith(PdfFontOptionsTests.ArabicLine);
        using var destination = new MemoryStream();

        await XlsxToPdfConverter.ConvertAsync(new MemoryStream(xlsx), destination, fonts);

        PdfFontOptionsTests.AssertArabicLineIsJoined(destination.ToArray(), font);
    }

    [Fact]
    public void PptxConvert_WithARealFont_ArabicIsPreShaped()
    {
        if (ArabicFontTests.FindArabicFont() is not { } font)
            return;
        var fonts = new PdfFontOptions("Supplied", font);
        var pptx = SlideWith(PdfFontOptionsTests.ArabicLine);

        PdfFontOptionsTests.AssertArabicLineIsJoined(PptxToPdfConverter.Convert(pptx, fonts), font);
    }

    [Fact]
    public async Task PptxConvertAsync_WithARealFont_ArabicIsPreShaped()
    {
        if (ArabicFontTests.FindArabicFont() is not { } font)
            return;
        var fonts = new PdfFontOptions("Supplied", font);
        var pptx = SlideWith(PdfFontOptionsTests.ArabicLine);
        using var destination = new MemoryStream();

        await PptxToPdfConverter.ConvertAsync(new MemoryStream(pptx), destination, fonts);

        PdfFontOptionsTests.AssertArabicLineIsJoined(destination.ToArray(), font);
    }

    // ---- Formulas are calculated BEFORE the letters are joined ---------------------------------------

    private const string Salam = "سلام";   // seen (sin lam alef mim): lam-alef is one glyph
    private const string Paid = "مدفوع";

    /// <summary>A1 text, B1 its length, C1 a comparison against an Arabic literal, D1 an Arabic text result.</summary>
    private static byte[] FormulaSheet() =>
        WorkbookEditor.Create("Sheet", new object?[][]
        {
            new object?[]
            {
                Salam,
                XlsxFormula.From("LEN(A1)"),
                XlsxFormula.From($"IF(A1=\"{Salam}\",\"ok\",\"no\")"),
                XlsxFormula.From($"IF(TRUE,\"{Paid}\",\"\")"),
            },
        });

    /// <summary>The same sheet holding the results as plain values: what the formulas should render as.</summary>
    private static byte[] ValueSheet() =>
        WorkbookEditor.Create("Sheet", new object?[][] { new object?[] { Salam, 4, "ok", Paid } });

    private static void AssertFormulaResultsRender(byte[] pdf, PdfFontOptions fonts)
    {
        var runs = PdfProbe.TextRuns(pdf).Where(r => r.Text.Trim().Length > 0).ToList();

        Assert.Contains(runs, r => r.Text == "4");
        Assert.Contains(runs, r => r.Text == "ok");
        Assert.Equal(
            PdfProbe.TextRuns(XlsxToPdfConverter.Convert(ValueSheet(), fonts)).Where(r => r.Text.Trim().Length > 0),
            runs);
    }

    [Fact]
    public void XlsxConvert_WithARealFont_LenOfAnArabicCellCountsTheUnjoinedLetters()
    {
        if (ArabicHostFonts() is not { } fonts)
            return;

        var runs = PdfProbe.TextRuns(XlsxToPdfConverter.Convert(FormulaSheet(), fonts));

        Assert.Contains(runs, r => r.Text == "4");
    }

    [Fact]
    public void XlsxConvert_WithARealFont_ACompareAgainstAnArabicLiteralStillMatches()
    {
        if (ArabicHostFonts() is not { } fonts)
            return;

        var runs = PdfProbe.TextRuns(XlsxToPdfConverter.Convert(FormulaSheet(), fonts));

        Assert.Contains(runs, r => r.Text == "ok");
    }

    [Fact]
    public void XlsxConvert_WithARealFont_AFormulasArabicTextResultIsJoined()
    {
        if (ArabicHostFonts() is not { } fonts)
            return;

        var formulas = PdfProbe.TextRuns(XlsxToPdfConverter.Convert(FormulaSheet(), fonts)).Where(r => r.Text.Trim().Length > 0);
        var values = PdfProbe.TextRuns(XlsxToPdfConverter.Convert(ValueSheet(), fonts)).Where(r => r.Text.Trim().Length > 0);

        Assert.Equal(values, formulas);
    }

    [Fact]
    public async Task XlsxConvertAsync_WithARealFont_FormulaResultsAreCalculatedBeforeJoining()
    {
        if (ArabicHostFonts() is not { } fonts)
            return;
        using var destination = new MemoryStream();

        await XlsxToPdfConverter.ConvertAsync(new MemoryStream(FormulaSheet()), destination, fonts);

        AssertFormulaResultsRender(destination.ToArray(), fonts);
    }

    /// <summary>
    /// No fonts: Arabic is refused as it is today (measured 2026-11-09, on hosts with and without
    /// fonts), and nothing changes it. The refusal is the pinned behaviour; that the pre-shaper is
    /// not reached is the DOCX counterpart's job, where the outcome can differ.
    /// </summary>
    [Fact]
    public void WithoutFonts_ArabicXlsxAndPptxStillRefuseAsToday()
    {
        var xlsx = Assert.Throws<DocumentConversionException>(
            () => XlsxToPdfConverter.Convert(SheetWith(PdfFontOptionsTests.ArabicLine)));
        var pptx = Assert.Throws<DocumentConversionException>(
            () => PptxToPdfConverter.Convert(SlideWith(PdfFontOptionsTests.ArabicLine)));

        Assert.Contains("WinAnsi", xlsx.InnerException!.Message, StringComparison.Ordinal);
        Assert.Contains("WinAnsi", pptx.InnerException!.Message, StringComparison.Ordinal);
    }
}
