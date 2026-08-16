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

    [Fact]
    public void PermissionBits_ReadsTheEncryptionDictionary_NotCiphertextThatLooksLikeIt()
    {
        // Pins the HELPER rather than the feature, for the same reason DocumentXml is pinned: every
        // other use of PermissionBits compares it against a constant, which holds however wrong the
        // helper is. This is the only test that would notice it reading the wrong part of the file.
        //
        // The decoy is placed where ciphertext lives - before the dictionary - because that is what
        // the previous implementation actually did: it took the first "/P <digits>" ANYWHERE, and
        // encrypted bytes are random, so it could pick one up.
        //
        // Laid out the way PDFsharp actually writes it - /CF nested BEFORE /Filter - because that
        // ordering broke the first attempt at the fix: "nearest << before /Filter" lands inside
        // /CF, which has no /P at all. A pin using a tidier layout would have gone green while the
        // helper read the wrong dictionary.
        var withDecoy =
            "%PDF-1.6\nstream\n£×/P 4á\nendstream\n"
            + "<</CF<</StdCF<</Type/CryptFilter/AuthEvent/DocOpen/CFM/AESV2/Length 16>>>>"
            + "/Filter/Standard/Length 128/P -24/R 4/StmF/StdCF/StrF/StdCF/V 4>>";

        Assert.Equal(-24, PermissionBits(Encoding.Latin1.GetBytes(withDecoy)));
    }

    // PDF 32000-1, Table 22. The bits are 1-based in the spec; these are the values.
    private const int PrintBit = 1 << 2;              // bit 3
    private const int CopyBit = 1 << 4;               // bit 5
    private const int HighQualityPrintBit = 1 << 11;  // bit 12

    /// <summary>
    /// The <c>/P</c> permission bitfield, read out of the <c>/Encrypt</c> dictionary specifically.
    /// </summary>
    /// <remarks>
    /// The dictionary is readable as plain text even though the document is encrypted, and
    /// necessarily so: a reader needs it before it can decrypt anything, so it is never itself
    /// encrypted.
    ///
    /// <b>It is located rather than searched for, and that is not fussiness.</b> The first version
    /// of this helper ran <c>/P\s+(-?\d+)</c> over the WHOLE file, which also scans the encrypted
    /// streams — and ciphertext differs on every run, so it could match random bytes ahead of the
    /// real entry. It failed CI on linux/net8.0 while the net10.0 leg of the same run passed.
    /// Measured afterwards: at ~15 KB of ciphertext, <b>1 run in 400</b> produced a file with more
    /// than one match where the first was wrong. A test that is right 399 times in 400 is not a
    /// test, it is a coin that usually lands the same way.
    ///
    /// So this finds the encryption dictionary by its <c>/Filter /Standard</c> entry and walks
    /// <c>&lt;&lt;</c>/<c>&gt;&gt;</c> depth to its end — depth matters, because <c>/CF</c> nests a
    /// dictionary inside it and a naive scan for the first <c>&gt;&gt;</c> stops early.
    /// </remarks>
    private static int PermissionBits(byte[] locked)
    {
        var text = Encoding.Latin1.GetString(locked);

        var filter = System.Text.RegularExpressions.Regex.Match(text, @"/Filter\s*/Standard");
        Assert.True(filter.Success, "no /Filter /Standard entry - the document is not encrypted, "
                                    + "so every assertion below would be vacuous");

        // Walk BACKWARDS with depth tracking to find the dictionary that ENCLOSES /Filter, not the
        // nearest "<<" before it. Those are different: PDFsharp writes the nested /CF crypt-filter
        // dictionary ahead of /Filter, so "nearest preceding" lands inside /CF and finds no /P at
        // all. Caught by this test failing rather than by reading the spec.
        var open = -1;
        for (int i = filter.Index - 1, depth = 0; i > 0; i--)
        {
            if (text[i] == '>' && text[i - 1] == '>') { depth++; i--; }
            else if (text[i] == '<' && text[i - 1] == '<')
            {
                if (depth == 0) { open = i - 1; break; }
                depth--; i--;
            }
        }

        Assert.True(open >= 0, "no dictionary encloses /Filter /Standard");

        var end = -1;
        for (int i = open, depth = 0; i < text.Length - 1; i++)
        {
            if (text[i] == '<' && text[i + 1] == '<') { depth++; i++; }
            else if (text[i] == '>' && text[i + 1] == '>')
            {
                depth--; i++;
                if (depth == 0) { end = i + 1; break; }
            }
        }

        Assert.True(end > open, "the /Encrypt dictionary never closes");

        var dictionary = text[open..end];
        var match = System.Text.RegularExpressions.Regex.Match(dictionary, @"/P\s+(-?\d+)");
        Assert.True(match.Success, $"no /P entry in the encryption dictionary: {dictionary}");

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
