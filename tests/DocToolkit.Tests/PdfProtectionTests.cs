using System.Text;

namespace DocToolkit.Tests;

/// <summary>
/// PDF password protection.
///
/// <b>The load-bearing assertion in this file is the NEGATIVE one</b>: that opening the output
/// <i>without</i> the password FAILS. A test proving the right password works passes just as
/// happily against a document that was never encrypted at all — which is not hypothetical, it is
/// the exact trap a probe fell into while surveying this capability, reporting a passing encryption
/// test for a plaintext file. Every "it encrypted" claim here is paired with a refusal.
/// </summary>
public class PdfProtectionTests
{
    private static byte[] Pdf() => DocxToPdfConverter.Convert(
        DocxEditor.Create(new[] { DocxBlock.Paragraph("PROTECTED-SENTINEL") }));

    private static bool IsEncrypted(byte[] pdf) =>
        Encoding.Latin1.GetString(pdf).Contains("/Encrypt", StringComparison.Ordinal);

    // ---- the guarantee -----------------------------------------------------------------------

    [Fact]
    public void Protect_ProducesADocumentThatCannotBeOpenedWithoutThePassword()
    {
        var locked = PdfEditor.Protect(Pdf(), new PdfProtection { UserPassword = "open-me" });

        // Every other operation on PdfEditor refuses an encrypted document, so PageCount failing is
        // the observable form of "it really is locked".
        var ex = Assert.Throws<DocumentConversionException>(() => PdfEditor.PageCount(locked));
        Assert.Contains("password", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Protect_MarksTheDocumentEncrypted_AndTheSourceWasNot()
    {
        var plain = Pdf();
        var locked = PdfEditor.Protect(plain, new PdfProtection { UserPassword = "open-me" });

        // The second half is the control: without it, a test asserting "/Encrypt is present" would
        // pass against a fixture that happened to contain the string already.
        Assert.True(IsEncrypted(locked));
        Assert.False(IsEncrypted(plain));
    }

    [Fact]
    public void Unprotect_RestoresADocumentTheRestOfTheApiCanUse()
    {
        var plain = Pdf();
        var locked = PdfEditor.Protect(plain, new PdfProtection { UserPassword = "open-me" });

        var opened = PdfEditor.Unprotect(locked, "open-me");

        Assert.False(IsEncrypted(opened));
        Assert.Equal(PdfEditor.PageCount(plain), PdfEditor.PageCount(opened));
        // Content survived the round trip - not just the page count.
        Assert.Contains("PROTECTED-SENTINEL", string.Concat(PdfEditor.ExtractText(opened)),
            StringComparison.Ordinal);
    }

    [Fact]
    public void Unprotect_RefusesTheWrongPassword()
    {
        var locked = PdfEditor.Protect(Pdf(), new PdfProtection { UserPassword = "open-me" });

        var ex = Assert.Throws<DocumentConversionException>(() => PdfEditor.Unprotect(locked, "wrong"));

        // The message must not repeat "it may be password-protected" at somebody who just supplied
        // a password - that hides the actual fault.
        Assert.Contains("password is", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Unprotect_NeedsTheOwnerPassword_WhenTheDocumentHasBoth()
    {
        // Measured 2026-08-16, and it inverted the assumption this test was first written with.
        // Removing protection is a MODIFICATION, and the PDF format reserves modification for the
        // owner password - so a perfectly correct user password is refused here. Pinned because
        // the behaviour is surprising and the error message has to explain it.
        var locked = PdfEditor.Protect(Pdf(),
            new PdfProtection { UserPassword = "user-pw", OwnerPassword = "owner-pw" });

        Assert.False(IsEncrypted(PdfEditor.Unprotect(locked, "owner-pw")));

        var ex = Assert.Throws<DocumentConversionException>(() => PdfEditor.Unprotect(locked, "user-pw"));
        Assert.Contains("owner password", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Unprotect_AcceptsTheUserPassword_WhenThatIsTheOnlyOne()
    {
        // The control for the test above: the refusal there is about the owner password EXISTING,
        // not about user passwords being useless. Without this pair, "Unprotect needs the owner
        // password" would read as a blanket rule and the API would look broken.
        var locked = PdfEditor.Protect(Pdf(), new PdfProtection { UserPassword = "user-pw" });

        Assert.False(IsEncrypted(PdfEditor.Unprotect(locked, "user-pw")));
    }

    // ---- the owner-password-only case, which is NOT a lock -------------------------------------

    [Fact]
    public void Protect_WithOnlyAnOwnerPassword_LeavesTheDocumentReadable()
    {
        // This is the PDF specification working as designed, and it is the single most likely way
        // somebody ships a "protected" file that is not protected. Pinned so the behaviour is a
        // documented property rather than a surprise.
        var restricted = PdfEditor.Protect(Pdf(),
            new PdfProtection { OwnerPassword = "owner-pw", AllowPrinting = false });

        Assert.True(IsEncrypted(restricted));
        Assert.Contains("PROTECTED-SENTINEL", string.Concat(PdfEditor.ExtractText(restricted)),
            StringComparison.Ordinal);
    }

    [Fact]
    public void Protect_RequiresAtLeastOnePassword_AndSaysWhichToChoose()
    {
        var ex = Assert.Throws<ArgumentException>(() => PdfEditor.Protect(Pdf(), new PdfProtection()));

        Assert.Contains("UserPassword", ex.Message, StringComparison.Ordinal);
        Assert.Contains("OwnerPassword", ex.Message, StringComparison.Ordinal);
    }

    // ---- strength ------------------------------------------------------------------------------

    [Theory]
    [InlineData(PdfEncryptionStrength.Aes128, "AESV2")]
    [InlineData(PdfEncryptionStrength.Aes256, "AESV3")]
    public void Protect_UsesTheRequestedCipher(PdfEncryptionStrength strength, string marker)
    {
        var locked = PdfEditor.Protect(Pdf(),
            new PdfProtection { UserPassword = "open-me", Strength = strength });

        // Reads the crypt-filter name out of the file rather than trusting the option was applied.
        Assert.Contains(marker, Encoding.Latin1.GetString(locked), StringComparison.Ordinal);

        // Strength must never weaken the guarantee: both ciphers still refuse without the password.
        Assert.Throws<DocumentConversionException>(() => PdfEditor.PageCount(locked));
        Assert.False(IsEncrypted(PdfEditor.Unprotect(locked, "open-me")));
    }

    [Fact]
    public void Protect_DefaultsToAes128()
    {
        var locked = PdfEditor.Protect(Pdf(), new PdfProtection { UserPassword = "open-me" });

        var text = Encoding.Latin1.GetString(locked);
        Assert.Contains("AESV2", text, StringComparison.Ordinal);
        Assert.DoesNotContain("AESV3", text, StringComparison.Ordinal);
    }

    // ---- permissions ---------------------------------------------------------------------------

    [Fact]
    public void Protect_ClearsThePermissionBitsItWasAskedTo()
    {
        var restricted = PdfEditor.Protect(Pdf(), new PdfProtection
        {
            UserPassword = "open-me",
            AllowPrinting = false,
            AllowCopying = false,
        });

        var p = PermissionBits(restricted);
        Assert.Equal(0, p & PrintBit);
        Assert.Equal(0, p & CopyBit);
    }

    [Fact]
    public void Protect_LeavesEveryPermissionSet_WhenNoneIsWithheld()
    {
        // The control for the test above. Without it, "the print bit is clear" would pass against
        // an implementation that cleared every bit regardless of what was asked - and adding a
        // password would then silently forbid printing a document that could be printed before.
        var p = PermissionBits(PdfEditor.Protect(Pdf(), new PdfProtection { UserPassword = "open-me" }));

        Assert.NotEqual(0, p & PrintBit);
        Assert.NotEqual(0, p & CopyBit);
        Assert.NotEqual(0, p & HighQualityPrintBit);
    }

    // PDF 32000-1, Table 22. The bits are 1-based in the spec; these are the values.
    private const int PrintBit = 1 << 2;              // bit 3
    private const int CopyBit = 1 << 4;               // bit 5
    private const int HighQualityPrintBit = 1 << 11;  // bit 12

    /// <summary>
    /// The <c>/P</c> permission bitfield, read straight out of the file.
    /// </summary>
    /// <remarks>
    /// Readable as plain text even though the document is encrypted, and necessarily so: a reader
    /// needs the <c>/Encrypt</c> dictionary before it can decrypt anything, so that dictionary is
    /// never itself encrypted. Verified 2026-08-16 rather than assumed.
    /// </remarks>
    private static int PermissionBits(byte[] locked)
    {
        var match = System.Text.RegularExpressions.Regex.Match(
            Encoding.Latin1.GetString(locked), @"/P\s+(-?\d+)");

        Assert.True(match.Success, "no /P entry found - the assertion below would be vacuous");
        return int.Parse(match.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture);
    }

    // ---- arguments -----------------------------------------------------------------------------

    [Fact]
    public void Protect_And_Unprotect_RejectNullAndEmpty()
    {
        Assert.Throws<ArgumentNullException>(() => PdfEditor.Protect(null!, new PdfProtection { UserPassword = "x" }));
        Assert.Throws<ArgumentNullException>(() => PdfEditor.Protect(Pdf(), null!));
        Assert.Throws<ArgumentException>(() => PdfEditor.Protect(Array.Empty<byte>(), new PdfProtection { UserPassword = "x" }));

        Assert.Throws<ArgumentNullException>(() => PdfEditor.Unprotect(null!, "x"));
        Assert.Throws<ArgumentNullException>(() => PdfEditor.Unprotect(Pdf(), null!));
        Assert.Throws<ArgumentException>(() => PdfEditor.Unprotect(Array.Empty<byte>(), "x"));
    }

    // ---- the Stream overloads ------------------------------------------------------------------

    [Fact]
    public async Task ProtectAsync_WritesADocumentThatStillRefusesWithoutThePassword()
    {
        using var source = new MemoryStream(Pdf(), writable: false);
        using var destination = new MemoryStream();

        await PdfEditor.ProtectAsync(source, destination, new PdfProtection { UserPassword = "open-me" });

        var locked = destination.ToArray();
        Assert.True(IsEncrypted(locked));
        Assert.Throws<DocumentConversionException>(() => PdfEditor.PageCount(locked));
    }

    [Fact]
    public async Task UnprotectAsync_RoundTripsThroughStreams()
    {
        var locked = PdfEditor.Protect(Pdf(), new PdfProtection { UserPassword = "open-me" });

        using var source = new MemoryStream(locked, writable: false);
        using var destination = new MemoryStream();

        await PdfEditor.UnprotectAsync(source, destination, "open-me");

        Assert.Contains("PROTECTED-SENTINEL", string.Concat(PdfEditor.ExtractText(destination.ToArray())),
            StringComparison.Ordinal);
    }
}
