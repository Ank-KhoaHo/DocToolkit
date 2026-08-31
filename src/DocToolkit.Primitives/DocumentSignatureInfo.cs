namespace DocToolkit;

/// <summary>
/// The result of a structural signature inspection — whether a document carries a signature and
/// who <b>claims</b> to have signed it, without validating anything cryptographically.
/// </summary>
/// <remarks>
/// <b><see cref="Signers"/> is unvalidated.</b> It is read directly from each signing certificate's
/// subject name, the same way the certificate itself claims an identity — this method does not
/// check whether the signature is genuine, whether the certificate chains to anything trusted, or
/// whether the signed content has been tampered with since signing. Use
/// <c>ValidateSignatures</c> before treating a claimed identity as real.
/// </remarks>
public sealed class DocumentSignatureInfo
{
    /// <param name="hasSignatures">Whether the document carries at least one signature.</param>
    /// <param name="signatureCount">How many signatures the document carries.</param>
    /// <param name="signers">The claimed signer subject names, one entry per signing certificate found.</param>
    /// <exception cref="ArgumentNullException"><paramref name="signers"/> is null.</exception>
    public DocumentSignatureInfo(bool hasSignatures, int signatureCount, IEnumerable<string> signers)
    {
        ArgumentNullException.ThrowIfNull(signers);
        HasSignatures = hasSignatures;
        SignatureCount = signatureCount;
        Signers = signers.ToList();
    }

    /// <summary>Whether the document carries at least one signature.</summary>
    public bool HasSignatures { get; }

    /// <summary>How many signatures the document carries.</summary>
    public int SignatureCount { get; }

    /// <summary>The claimed signer subject names — see the remarks on this type before trusting one.</summary>
    public IReadOnlyList<string> Signers { get; }
}
