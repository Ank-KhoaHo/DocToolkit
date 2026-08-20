using System.Security.Cryptography;

namespace DocToolkit.Tests;

/// <summary>
/// File-level password protection for DOCX, XLSX and PPTX.
///
/// <b>Driven by a table so no format gets weaker coverage than another.</b> The three
/// implementations are separate — measured 2026-08-16, no OfficeIMO package exposes a
/// format-agnostic OOXML encryptor — so the risk is that one of them quietly does less than the
/// others. A per-format copy of each test is exactly how that goes unnoticed.
///
/// <b>The load-bearing assertions are the negative ones</b>: a wrong password must be REFUSED, and
/// the encrypted bytes must no longer be a ZIP. A round-trip proving the right password works
/// passes just as happily against a file that was never encrypted at all.
/// </summary>
public class OfficePasswordTests
{
    /// <summary>
    /// One row per format: a plain document, and the three entry points under test.
    /// </summary>
    public static TheoryData<string> Formats => new() { "docx", "xlsx", "pptx" };

    private static byte[] Plain(string format) => format switch
    {
        "docx" => DocxEditor.Create(new[] { DocxBlock.Paragraph("OFFICE-SENTINEL") }),
        "xlsx" => WorkbookEditor.Create("Sales", new[] { new object?[] { "OFFICE-SENTINEL", 1234 } }),
        "pptx" => PresentationEditor.Create(new[] { PptxSlide.Titled("OFFICE-SENTINEL") }),
        _ => throw new ArgumentOutOfRangeException(nameof(format)),
    };

    private static byte[] Protect(string format, byte[] doc, string password) => format switch
    {
        "docx" => DocxEditor.Protect(doc, password),
        "xlsx" => WorkbookEditor.Protect(doc, password),
        "pptx" => PresentationEditor.Protect(doc, password),
        _ => throw new ArgumentOutOfRangeException(nameof(format)),
    };

    private static byte[] Unprotect(string format, byte[] doc, string password) => format switch
    {
        "docx" => DocxEditor.Unprotect(doc, password),
        "xlsx" => WorkbookEditor.Unprotect(doc, password),
        "pptx" => PresentationEditor.Unprotect(doc, password),
        _ => throw new ArgumentOutOfRangeException(nameof(format)),
    };

    private static bool IsProtected(string format, byte[] doc) => format switch
    {
        "docx" => DocxEditor.IsProtected(doc),
        "xlsx" => WorkbookEditor.IsProtected(doc),
        "pptx" => PresentationEditor.IsProtected(doc),
        _ => throw new ArgumentOutOfRangeException(nameof(format)),
    };

    /// <summary>The text a round-tripped document must still contain, read the format's own way.</summary>
    private static string TextOf(string format, byte[] doc) => format switch
    {
        "docx" => DocxEditor.ExtractText(doc),
        "xlsx" => WorkbookEditor.ReadCell(doc, "Sales", "A1"),
        "pptx" => string.Join("\n", PresentationEditor.ExtractText(doc)),
        _ => throw new ArgumentOutOfRangeException(nameof(format)),
    };

    private static bool IsZip(byte[] b) => b.Length >= 2 && b[0] == 0x50 && b[1] == 0x4B;

    // ---- the guarantee -------------------------------------------------------------------------

    [Theory]
    [MemberData(nameof(Formats))]
    public void Protect_ProducesSomethingThatIsNoLongerAnOpenPackage(string format)
    {
        var plain = Plain(format);
        var locked = Protect(format, plain, "s3cret");

        // A plain OOXML file is a ZIP; the encrypted form is a compound file. The pair is the
        // control - asserting only "the locked one is not a ZIP" would pass if Plain() were broken.
        Assert.True(IsZip(plain));
        Assert.False(IsZip(locked));
        Assert.True(IsProtected(format, locked));
        Assert.False(IsProtected(format, plain));
    }

    [Theory]
    [MemberData(nameof(Formats))]
    public void Protect_MakesTheOrdinaryMethodsRefuseTheDocument(string format)
    {
        var locked = Protect(format, Plain(format), "s3cret");

        // The observable form of "it really is locked": the rest of the class cannot read it.
        Assert.ThrowsAny<Exception>(() => TextOf(format, locked));
    }

    [Theory]
    [MemberData(nameof(Formats))]
    public void Unprotect_ReturnsTheContent_NotJustADocument(string format)
    {
        var locked = Protect(format, Plain(format), "s3cret");

        var opened = Unprotect(format, locked, "s3cret");

        Assert.True(IsZip(opened));
        // The literal, not "is not empty" - the whole point is that the DATA survived.
        Assert.Contains("OFFICE-SENTINEL", TextOf(format, opened), StringComparison.Ordinal);
    }

    [Theory]
    [MemberData(nameof(Formats))]
    public void Unprotect_RefusesTheWrongPassword(string format)
    {
        var locked = Protect(format, Plain(format), "s3cret");

        var ex = Assert.Throws<DocumentConversionException>(() => Unprotect(format, locked, "WRONG"));

        // Must say the password was wrong, not that the file is broken - they are different faults
        // and the caller can only act on one of them.
        Assert.Contains("password", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.IsType<CryptographicException>(ex.InnerException);
    }

    [Theory]
    [MemberData(nameof(Formats))]
    public void Unprotect_SaysSo_WhenTheDocumentWasNeverEncrypted(string format)
    {
        // Silently returning the input would make a broken pipeline look like a working one.
        var ex = Assert.Throws<DocumentConversionException>(
            () => Unprotect(format, Plain(format), "s3cret"));

        Assert.Contains("not encrypted", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    // ---- arguments -----------------------------------------------------------------------------

    [Theory]
    [MemberData(nameof(Formats))]
    public void Protect_RejectsAnEmptyPassword(string format)
    {
        // Not a weak password - an unopenable file. The format cannot express "encrypted with
        // nothing", so this fails rather than producing something no reader can open.
        var ex = Assert.Throws<ArgumentException>(() => Protect(format, Plain(format), ""));
        Assert.Contains("empty password", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [MemberData(nameof(Formats))]
    public void EveryEntryPoint_RejectsNullAndEmptyInput(string format)
    {
        Assert.Throws<ArgumentNullException>(() => Protect(format, null!, "pw"));
        Assert.Throws<ArgumentNullException>(() => Protect(format, Plain(format), null!));
        Assert.Throws<ArgumentException>(() => Protect(format, Array.Empty<byte>(), "pw"));

        Assert.Throws<ArgumentNullException>(() => Unprotect(format, null!, "pw"));
        Assert.Throws<ArgumentException>(() => Unprotect(format, Array.Empty<byte>(), "pw"));

        Assert.Throws<ArgumentNullException>(() => IsProtected(format, null!));
    }

    [Theory]
    [MemberData(nameof(Formats))]
    public void IsProtected_IsFalseForBytesThatAreNotOfficeDocumentsAtAll(string format)
    {
        // Discriminates: a signature check that returned true for everything would still satisfy
        // the "encrypted file is protected" assertion above.
        _ = format;
        Assert.False(IsProtected(format, new byte[] { 1, 2, 3 }));
        Assert.False(IsProtected(format, Array.Empty<byte>()));
    }

    /// <summary>
    /// <c>IsProtected</c> answering <see langword="false"/> does NOT mean the other methods will
    /// accept the input, and the doc comment now says so.
    /// </summary>
    /// <remarks>
    /// The summary used to read <i>"whether <c>docx</c> is an encrypted Office document — that is,
    /// whether the other methods on this class will refuse it"</i>. The clause after "that is" was
    /// false for every input that is not a document at all: a JPEG is not encrypted, so
    /// <c>IsProtected</c> says <see langword="false"/>, while <c>ExtractText</c> refuses it.
    ///
    /// That reads as a guard - test it, and if false, proceed - so it took the wrong branch in
    /// exactly the situation it was written for. **The behaviour is correct and deliberate; only
    /// the sentence was wrong.** This test exists so the corrected sentence is a checkable claim
    /// rather than more prose, which is the failure mode this repository keeps correcting.
    /// </remarks>
    [Theory]
    [MemberData(nameof(Formats))]
    public void IsProtectedIsNotAValidityCheck(string format)
    {
        var notADocument = System.Text.Encoding.UTF8.GetBytes("this is plainly not an Office file");

        // Both halves together are the point. Either alone is satisfied by a weaker contract.
        Assert.False(IsProtected(format, notADocument));
        Assert.ThrowsAny<DocumentConversionException>(() => TextOf(format, notADocument));
    }

    // ---- the Stream overloads ------------------------------------------------------------------

    [Theory]
    [MemberData(nameof(Formats))]
    public async Task StreamOverloads_RoundTripTheSameWayTheByteArrayOnesDo(string format)
    {
        var plain = Plain(format);

        using var toProtect = new MemoryStream(plain, writable: false);
        using var locked = new MemoryStream();
        await ProtectAsync(format, toProtect, locked, "s3cret");

        Assert.False(IsZip(locked.ToArray()));

        using var toOpen = new MemoryStream(locked.ToArray(), writable: false);
        using var opened = new MemoryStream();
        await UnprotectAsync(format, toOpen, opened, "s3cret");

        Assert.Contains("OFFICE-SENTINEL", TextOf(format, opened.ToArray()), StringComparison.Ordinal);
    }

    private static Task ProtectAsync(string format, Stream s, Stream d, string pw) => format switch
    {
        "docx" => DocxEditor.ProtectAsync(s, d, pw),
        "xlsx" => WorkbookEditor.ProtectAsync(s, d, pw),
        "pptx" => PresentationEditor.ProtectAsync(s, d, pw),
        _ => throw new ArgumentOutOfRangeException(nameof(format)),
    };

    private static Task UnprotectAsync(string format, Stream s, Stream d, string pw) => format switch
    {
        "docx" => DocxEditor.UnprotectAsync(s, d, pw),
        "xlsx" => WorkbookEditor.UnprotectAsync(s, d, pw),
        "pptx" => PresentationEditor.UnprotectAsync(s, d, pw),
        _ => throw new ArgumentOutOfRangeException(nameof(format)),
    };
}
