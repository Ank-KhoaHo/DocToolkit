namespace DocToolkit;

/// <summary>One signature's independently validated state.</summary>
/// <remarks>
/// The three status fields are deliberately independent, matching what OfficeIMO's own validator
/// reports rather than this package computing a single opinion: a self-signed certificate that
/// never chases up to a trusted root is normal for an internal enterprise signer, and produces
/// <see cref="CryptographicStatus"/> = <see cref="DocumentSignatureStatus.Passed"/> (the signed
/// content was not tampered with) alongside <see cref="CertificateChainStatus"/> =
/// <see cref="DocumentSignatureStatus.Failed"/> (untrusted root) — measured directly, not assumed.
/// Read the field that answers the question you actually have.
/// </remarks>
public sealed class DocumentSignatureValidationResult
{
    /// <param name="cryptographicStatus">Whether the signed content matches what was signed — tamper detection.</param>
    /// <param name="certificateChainStatus">Whether the signing certificate chains to a certificate this machine trusts.</param>
    /// <param name="revocationStatus">Whether the signing certificate has been revoked.</param>
    /// <param name="signers">This signature's own claimed signer subject names.</param>
    /// <exception cref="ArgumentNullException"><paramref name="signers"/> is null.</exception>
    public DocumentSignatureValidationResult(
        DocumentSignatureStatus cryptographicStatus,
        DocumentSignatureStatus certificateChainStatus,
        DocumentSignatureStatus revocationStatus,
        IEnumerable<string> signers)
    {
        ArgumentNullException.ThrowIfNull(signers);
        CryptographicStatus = cryptographicStatus;
        CertificateChainStatus = certificateChainStatus;
        RevocationStatus = revocationStatus;
        Signers = signers.ToList();
    }

    /// <summary>Whether the signed content matches what was signed — tamper detection.</summary>
    public DocumentSignatureStatus CryptographicStatus { get; }

    /// <summary>Whether the signing certificate chains to a certificate this machine trusts.</summary>
    public DocumentSignatureStatus CertificateChainStatus { get; }

    /// <summary>Whether the signing certificate has been revoked. Never checked over the network — see <see cref="DocumentSignatureValidationReport"/>.</summary>
    public DocumentSignatureStatus RevocationStatus { get; }

    /// <summary>This signature's own claimed signer subject names.</summary>
    public IReadOnlyList<string> Signers { get; }
}
