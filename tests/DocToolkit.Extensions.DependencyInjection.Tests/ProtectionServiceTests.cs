using DocToolkit;
using DocToolkit.Extensions.DependencyInjection;
using Microsoft.Extensions.Options;
using Xunit;

namespace DocToolkit.Extensions.DependencyInjection.Tests;

/// <summary>
/// The password-protection members mirrored in 0.28.0, plus the legacy .doc converter.
///
/// <b>These services are pure delegation, so the risk is not that the logic is wrong — it is that a
/// member delegates to the WRONG thing, or silently does nothing.</b> A test that only checked "it
/// did not throw" would miss both. So each assertion here is one a passthrough would fail: an
/// encrypted document is no longer a package, a wrong password is refused, and the wrapper's output
/// matches what the static API produces for the same input.
///
/// The DI package is held at 100% coverage precisely because it is delegation — an uncovered member
/// is a member nobody checked was wired to anything.
/// </summary>
public class ProtectionServiceTests
{
    private static byte[] DocxBytes() => DocxEditor.Create([DocxBlock.Paragraph("DI-SENTINEL")]);

    private static DocxEditorService Docx()
        => new(new TestOptionsMonitor<DocToolkitOptions>(new DocToolkitOptions()));
    private static byte[] Xlsx() => WorkbookEditor.Create("Sales", [new object?[] { "DI-SENTINEL" }]);
    private static byte[] Pptx() => PresentationEditor.Create([PptxSlide.Titled("DI-SENTINEL")]);

    private static bool IsZip(byte[] b) => b.Length >= 2 && b[0] == 0x50 && b[1] == 0x4B;

    // ---- the three Office editors --------------------------------------------------------------

    public static TheoryData<string> OfficeFormats => new() { "docx", "xlsx", "pptx" };

    private static byte[] Plain(string f) => f switch
    {
        "docx" => DocxBytes(),
        "xlsx" => Xlsx(),
        "pptx" => Pptx(),
        _ => throw new ArgumentOutOfRangeException(nameof(f)),
    };

    private static (Func<byte[], string, byte[]> Protect,
                   Func<byte[], string, byte[]> Unprotect,
                   Func<byte[], bool> IsProtected) Sut(string f) => f switch
                   {
                       // DocxEditorService is the only one of the three that takes the options monitor - it is
                       // the one whose other members can fetch remote images. Protection ignores it entirely.
                       "docx" => (Docx().Protect, Docx().Unprotect, Docx().IsProtected),
                       "xlsx" => (new WorkbookEditorService().Protect, new WorkbookEditorService().Unprotect,
                                  new WorkbookEditorService().IsProtected),
                       "pptx" => (new PresentationEditorService().Protect, new PresentationEditorService().Unprotect,
                                  new PresentationEditorService().IsProtected),
                       _ => throw new ArgumentOutOfRangeException(nameof(f)),
                   };

    [Theory]
    [MemberData(nameof(OfficeFormats))]
    public void Protect_ThroughTheService_ReallyEncrypts(string format)
    {
        var (protect, _, isProtected) = Sut(format);
        var plain = Plain(format);

        var locked = protect(plain, "s3cret");

        // A passthrough would leave a ZIP behind and fail both halves.
        Assert.False(IsZip(locked));
        Assert.True(isProtected(locked));
        Assert.False(isProtected(plain));
    }

    [Theory]
    [MemberData(nameof(OfficeFormats))]
    public void Unprotect_ThroughTheService_RestoresTheContent(string format)
    {
        var (protect, unprotect, _) = Sut(format);

        var opened = unprotect(protect(Plain(format), "s3cret"), "s3cret");

        Assert.True(IsZip(opened));
    }

    [Theory]
    [MemberData(nameof(OfficeFormats))]
    public void Unprotect_ThroughTheService_RefusesTheWrongPassword(string format)
    {
        var (protect, unprotect, _) = Sut(format);
        var locked = protect(Plain(format), "s3cret");

        Assert.Throws<DocumentConversionException>(() => unprotect(locked, "WRONG"));
    }

    private static (Func<Stream, Stream, string, Task> Protect,
                    Func<Stream, Stream, string, Task> Unprotect) StreamSut(string f) => f switch
                    {
                        "docx" => ((s, d, pw) => Docx().ProtectAsync(s, d, pw), (s, d, pw) => Docx().UnprotectAsync(s, d, pw)),
                        "xlsx" => ((s, d, pw) => new WorkbookEditorService().ProtectAsync(s, d, pw),
                                   (s, d, pw) => new WorkbookEditorService().UnprotectAsync(s, d, pw)),
                        "pptx" => ((s, d, pw) => new PresentationEditorService().ProtectAsync(s, d, pw),
                                   (s, d, pw) => new PresentationEditorService().UnprotectAsync(s, d, pw)),
                        _ => throw new ArgumentOutOfRangeException(nameof(f)),
                    };

    [Theory]
    [MemberData(nameof(OfficeFormats))]
    public async Task OfficeStreamOverloads_AreWiredForEveryFormat(string format)
    {
        // Written once for XLSX only, on the reasoning that "the three share a template". The 100%
        // coverage gate rejected that and was right: each service is a SEPARATE delegation, and a
        // template proves nothing about whether a given class calls the method it names. Two lines
        // in DocxEditorService and two in PresentationEditorService were uncovered.
        var (protect, unprotect) = StreamSut(format);

        using var source = new MemoryStream(Plain(format), writable: false);
        using var locked = new MemoryStream();
        await protect(source, locked, "s3cret");
        Assert.False(IsZip(locked.ToArray()));

        using var toOpen = new MemoryStream(locked.ToArray(), writable: false);
        using var opened = new MemoryStream();
        await unprotect(toOpen, opened, "s3cret");
        Assert.True(IsZip(opened.ToArray()));
    }

    // ---- PdfEditor -----------------------------------------------------------------------------

    [Fact]
    public void PdfProtect_ThroughTheService_ProducesADocumentThatNeedsThePassword()
    {
        var pdf = DocxToPdfConverter.Convert(DocxBytes());
        var sut = new PdfEditorService();

        var locked = sut.Protect(pdf, new PdfProtection { UserPassword = "s3cret" });

        // The observable form of "really locked": every other member refuses it.
        Assert.Throws<DocumentConversionException>(() => sut.PageCount(locked));
        Assert.Equal(PdfEditor.PageCount(pdf), sut.PageCount(sut.Unprotect(locked, "s3cret")));
    }

    [Fact]
    public void PdfProtect_ThroughTheService_PassesThePermissionsAlong()
    {
        // Discriminates between "wired to Protect" and "wired to Protect and actually handing over
        // the options object" - a service that dropped `protection` would still encrypt.
        var pdf = DocxToPdfConverter.Convert(DocxBytes());
        var sut = new PdfEditorService();

        var strong = sut.Protect(pdf, new PdfProtection
        {
            UserPassword = "s3cret",
            Strength = PdfEncryptionStrength.Aes256,
        });

        Assert.Contains("AESV3", System.Text.Encoding.Latin1.GetString(strong), StringComparison.Ordinal);
    }

    [Fact]
    public async Task PdfStreamOverloads_AreWiredToo()
    {
        var sut = new PdfEditorService();
        var pdf = DocxToPdfConverter.Convert(DocxBytes());

        using var source = new MemoryStream(pdf, writable: false);
        using var locked = new MemoryStream();
        await sut.ProtectAsync(source, locked, new PdfProtection { UserPassword = "s3cret" });

        using var toOpen = new MemoryStream(locked.ToArray(), writable: false);
        using var opened = new MemoryStream();
        await sut.UnprotectAsync(toOpen, opened, "s3cret");

        Assert.Equal(PdfEditor.PageCount(pdf), sut.PageCount(opened.ToArray()));
    }

    // ---- IDocToDocxConverter -------------------------------------------------------------------

    private static byte[] LegacyDoc() =>
        File.ReadAllBytes(Path.Join(AppContext.BaseDirectory, "assets", "legacy.doc"));

    private static byte[] LosslessDoc() =>
        File.ReadAllBytes(Path.Join(AppContext.BaseDirectory, "assets", "legacy-lossless.doc"));

    [Fact]
    public void DocToDocx_ExtractText_IsWired()
    {
        var text = new DocToDocxConverterService().ExtractText(LegacyDoc());

        Assert.Contains("SENTINEL-TWO", text, StringComparison.Ordinal);
    }

    [Fact]
    public void DocToDocx_Convert_KeepsTheRefusalDefault_AndTheOptionsOverloadWorks()
    {
        var sut = new DocToDocxConverterService();

        // The default must still refuse THROUGH the wrapper - a service that quietly passed
        // AllowContentLoss = true would be a security-shaped surprise, not a convenience.
        Assert.Throws<DocumentConversionException>(() => sut.Convert(LegacyDoc()));

        var converted = sut.Convert(LegacyDoc(), new LegacyDocOptions { AllowContentLoss = true });
        Assert.True(IsZip(converted));
        Assert.Contains("SENTINEL-TWO", DocxEditor.ExtractText(converted), StringComparison.Ordinal);
    }

    [Fact]
    public void DocToDocx_ConvertWithReport_IsWired()
    {
        var result = new DocToDocxConverterService()
            .ConvertWithReport(LegacyDoc(), new LegacyDocOptions { AllowContentLoss = true });

        Assert.Contains(result.Warnings, w => w.Code == "DOC-BINARY-DATA-STREAM-PRESENT");
    }

    [Fact]
    public async Task DocToDocx_StreamOverloads_AreWiredToo()
    {
        var sut = new DocToDocxConverterService();

        using var source = new MemoryStream(LosslessDoc(), writable: false);
        using var destination = new MemoryStream();
        await sut.ConvertAsync(source, destination);
        Assert.True(IsZip(destination.ToArray()));

        using var forText = new MemoryStream(LegacyDoc(), writable: false);
        Assert.Contains("SENTINEL-TWO", await sut.ExtractTextAsync(forText), StringComparison.Ordinal);
    }

    [Fact]
    public async Task DocToDocx_ConvertAsync_WithOptions_IsADistinctOverload()
    {
        // Without this, the four-argument overload would be uncovered - and an overload nobody
        // calls is exactly where a wrapper forgets to pass its options along.
        var sut = new DocToDocxConverterService();

        using var source = new MemoryStream(LegacyDoc(), writable: false);
        using var destination = new MemoryStream();
        await sut.ConvertAsync(source, destination, new LegacyDocOptions { AllowContentLoss = true });

        Assert.Contains("SENTINEL-TWO", DocxEditor.ExtractText(destination.ToArray()),
            StringComparison.Ordinal);
    }
}
