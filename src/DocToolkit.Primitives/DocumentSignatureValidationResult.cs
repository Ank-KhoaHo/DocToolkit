namespace DocToolkit;

/// <summary>One signature's independently validated state.</summary>
/// <remarks>
/// The status fields are deliberately independent, matching what OfficeIMO's own validator
/// reports rather than this package computing a single opinion: a self-signed certificate that
/// never chases up to a trusted root is normal for an internal enterprise signer, and produces
/// <see cref="CryptographicStatus"/> = <see cref="DocumentSignatureStatus.Passed"/> (the signature
/// block itself is well-formed) alongside <see cref="CertificateChainStatus"/> =
/// <see cref="DocumentSignatureStatus.Failed"/> (untrusted root) — measured directly, not assumed.
/// Read the field that answers the question you actually have.
///
/// <b><see cref="CryptographicStatus"/> is not tamper detection — read
/// <see cref="DocumentSignatureValidationReport.IsCryptographicallyValid"/> for that.</b> Measured
/// directly: signing a document, then altering its content afterward without re-signing, leaves
/// <see cref="CryptographicStatus"/> at <see cref="DocumentSignatureStatus.Passed"/> on both the
/// untampered and the tampered copy — the signature value still verifies against its own
/// <c>SignedInfo</c> block either way. What changes is whether the covered content's digest still
/// matches, which this type does not expose per signature; only the report-level
/// <see cref="DocumentSignatureValidationReport.IsCryptographicallyValid"/> reflects it, and for a
/// document carrying more than one signature that is an aggregate across all of them, not a
/// per-signature answer.
/// </remarks>
public sealed class DocumentSignatureValidationResult
{
    /// <summary>Creates one signature's independently validated state.</summary>
    /// <param name="cryptographicStatus">
    /// Whether this signature's own cryptographic signature value verifies — not tamper detection,
    /// see the remarks on this type.
    /// </param>
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

    /// <summary>
    /// Whether this signature's own cryptographic signature value verifies against its
    /// <c>SignedInfo</c> block — confirms the signature block is well-formed, not that the covered
    /// content is unchanged since signing. See the remarks on this type before reading this as
    /// tamper detection.
    /// </summary>
    public DocumentSignatureStatus CryptographicStatus { get; }

    /// <summary>Whether the signing certificate chains to a certificate this machine trusts.</summary>
    public DocumentSignatureStatus CertificateChainStatus { get; }

    /// <summary>Whether the signing certificate has been revoked. Never checked over the network — see <see cref="DocumentSignatureValidationReport"/>.</summary>
    public DocumentSignatureStatus RevocationStatus { get; }

    /// <summary>This signature's own claimed signer subject names.</summary>
    public IReadOnlyList<string> Signers { get; }
}
