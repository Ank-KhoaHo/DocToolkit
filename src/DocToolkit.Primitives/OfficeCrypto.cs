using System.Security.Cryptography;

namespace DocToolkit;

/// <summary>
/// The one place the three Office editors validate a password and translate an encryption failure.
/// </summary>
/// <remarks>
/// <b>The encryption itself is per format and cannot be shared</b> — measured 2026-08-16: there is
/// no format-agnostic OOXML encryptor in the public surface of any OfficeIMO package, so DOCX, XLSX
/// and PPTX each go through their own type. What <i>can</i> be shared is everything around it, and
/// that is what lives here: three editors must not disagree about when a password is invalid or
/// about what a wrong one says, because those are the parts a caller sees.
///
/// <b>An encrypted OOXML file is not a ZIP.</b> A plain .docx/.xlsx/.pptx is a ZIP package
/// (<c>PK</c>); the encrypted form is a CFB/OLE2 container (<c>D0 CF 11 E0</c>) with the package
/// sealed inside it. That is why an encrypted document cannot be handed to any other method on
/// these classes — it is not the format they read.
/// </remarks>
internal static class OfficeCrypto
{
    /// <summary>
    /// The first bytes of a compound-file container, which is what an encrypted OOXML document is.
    /// </summary>
    private static readonly byte[] CompoundFileSignature = { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };

    internal static void RequirePassword(string password, string paramName)
    {
        ArgumentNullException.ThrowIfNull(password, paramName);

        // An empty password is not a weak password, it is a broken file: the format has no way to
        // express "encrypted with nothing", so this fails here rather than producing a document
        // that cannot be opened by anything.
        if (password.Length == 0)
        {
            throw new ArgumentException(
                "A password is required. An empty password cannot be used to encrypt an Office "
                + "document — omit the call entirely if the document should not be protected.",
                paramName);
        }
    }

    /// <summary>Whether <paramref name="document"/> looks like an encrypted OOXML file.</summary>
    /// <remarks>
    /// A signature check, not a decryption attempt, so it costs nothing and needs no password. It
    /// answers "would the other methods on this class refuse this?" — which is the question a
    /// caller holding an unknown file actually has.
    /// </remarks>
    internal static bool IsEncrypted(byte[] document)
    {
        if (document.Length < CompoundFileSignature.Length) return false;

        for (var i = 0; i < CompoundFileSignature.Length; i++)
            if (document[i] != CompoundFileSignature[i])
                return false;

        return true;
    }

    /// <summary>
    /// Runs <paramref name="work"/>, turning every way decryption can fail into this library's own
    /// exception with a message that says which of them happened.
    /// </summary>
    /// <remarks>
    /// The two upstream failures are genuinely different and a caller can act on the difference, so
    /// they are not flattened into one message. A <see cref="CryptographicException"/> means the
    /// password was wrong; an <see cref="InvalidDataException"/> means the file was never encrypted
    /// — measured 2026-08-16 for all three formats, both with controls.
    /// </remarks>
    internal static T Translate<T>(Func<T> work, string format)
    {
        try
        {
            return work();
        }
        catch (CryptographicException ex)
        {
            throw new DocumentConversionException(
                $"The password did not open this {format}. Check the password — the file itself "
                + "looks like a correctly encrypted Office document.", ex);
        }
        catch (InvalidDataException ex)
        {
            throw new DocumentConversionException(
                $"This {format} is not encrypted, so it has no password to supply. Open it with the "
                + "ordinary methods on this class instead.", ex);
        }
        catch (Exception ex) when (ex is not DocumentConversionException and not ArgumentException)
        {
            throw new DocumentConversionException(
                $"Failed to read the encrypted {format}. See the inner exception for details.", ex);
        }
    }

    /// <summary>Runs <paramref name="work"/>, wrapping any failure to encrypt.</summary>
    internal static byte[] TranslateWrite(Func<byte[]> work, string format)
    {
        try
        {
            return work();
        }
        catch (Exception ex) when (ex is not DocumentConversionException and not ArgumentException)
        {
            throw new DocumentConversionException(
                $"Failed to encrypt the {format}. See the inner exception for details.", ex);
        }
    }
}
