namespace DocToolkit;

/// <summary>Options for <c>ValidateSignatures</c>.</summary>
/// <remarks>
/// There is deliberately no way to enable revocation checking or certificate downloads here — see
/// <see cref="DocumentSignatureValidationReport"/>'s remarks. If a future version adds an opt-in
/// online check, it will be a new, explicitly named option, not a change to what this type's
/// current defaults mean.
///
/// <b>There is deliberately no way to trust an internal certificate authority here either — an
/// earlier draft of this type had one (<c>AdditionalTrustedCertificates</c>) and it was removed
/// before release.</b> Measured directly: OfficeIMO's underlying <c>ExtraCertificates</c> option
/// supplies certificates only for chain-<i>building</i> (resolving a missing intermediate), not
/// for chain-<i>trust</i> — passing a document's own issuing CA through it left
/// <c>CertificateChainStatus</c> at <see cref="DocumentSignatureStatus.Failed"/>, identically to
/// not passing it at all. Shipping that property with documentation claiming it worked would have
/// been exactly the kind of "present but wrong" answer about signature validity this feature
/// exists to avoid. To trust an internal CA, install it in the trust store this machine's chain
/// building already consults, and use <see cref="ValidateCertificateTrust"/> to opt entirely out
/// of chain checking if that is not available.
/// </remarks>
public sealed class DocumentSignatureValidationOptions
{
    /// <summary>
    /// Whether to check the signing certificate's chain against this machine's local trust store.
    /// Purely local — chain building against an already-present store needs no network access.
    /// Default <see langword="true"/>.
    /// </summary>
    public bool ValidateCertificateTrust { get; set; } = true;
}
