using System.Security.Cryptography.X509Certificates;

namespace DocToolkit;

/// <summary>Options for <c>ValidateSignatures</c>.</summary>
/// <remarks>
/// There is deliberately no way to enable revocation checking or certificate downloads here — see
/// <see cref="DocumentSignatureValidationReport"/>'s remarks. If a future version adds an opt-in
/// online check, it will be a new, explicitly named option, not a change to what this type's
/// current defaults mean.
/// </remarks>
public sealed class DocumentSignatureValidationOptions
{
    /// <summary>
    /// Whether to check the signing certificate's chain against this machine's local trust store
    /// and <see cref="AdditionalTrustedCertificates"/>. Purely local — chain building against an
    /// already-present store needs no network access. Default <see langword="true"/>.
    /// </summary>
    public bool ValidateCertificateTrust { get; set; } = true;

    /// <summary>
    /// Extra certificates to trust for chain building, beyond this machine's local store — the
    /// escape hatch for an internal enterprise certificate authority that is not, and should not
    /// be, in the OS trust store. Default empty.
    /// </summary>
    public IReadOnlyList<X509Certificate2> AdditionalTrustedCertificates { get; set; } = Array.Empty<X509Certificate2>();
}
