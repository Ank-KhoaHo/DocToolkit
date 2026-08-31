namespace DocToolkit;

/// <summary>The result of validating every signature a document carries.</summary>
/// <remarks>
/// <b>No revocation check and no network access ever happen, on any call, regardless of the
/// options passed.</b> <see cref="DocumentSignatureValidationResult.RevocationStatus"/> reflects
/// only what a purely local check can determine.
///
/// <b>An unsigned document reports <see cref="IsCryptographicallyValid"/> = <see langword="false"/>,
/// the same as a tampered one</b> — check <see cref="HasSignatures"/> first. "No signature" and
/// "a signature that failed" are different findings a caller must not conflate; measured directly
/// against a genuinely unsigned fixture, which reports <c>HasSignatures = false</c> with a single
/// finding, no exception.
/// </remarks>
public sealed class DocumentSignatureValidationReport
{
    /// <param name="hasSignatures">Whether the document carries at least one signature.</param>
    /// <param name="isCryptographicallyValid">
    /// Whether every signature's covered content matches what was signed — the aggregate tamper-detection verdict.
    /// </param>
    /// <param name="isValidUnderPolicy">
    /// Whether every signature passed every check this call actually performed (per the options given).
    /// </param>
    /// <param name="signatures">One result per signature the document carries.</param>
    /// <param name="findings">Human-readable diagnostic messages from the underlying validator.</param>
    /// <exception cref="ArgumentNullException"><paramref name="signatures"/> or <paramref name="findings"/> is null.</exception>
    public DocumentSignatureValidationReport(
        bool hasSignatures, bool isCryptographicallyValid, bool isValidUnderPolicy,
        IEnumerable<DocumentSignatureValidationResult> signatures, IEnumerable<string> findings)
    {
        ArgumentNullException.ThrowIfNull(signatures);
        ArgumentNullException.ThrowIfNull(findings);
        HasSignatures = hasSignatures;
        IsCryptographicallyValid = isCryptographicallyValid;
        IsValidUnderPolicy = isValidUnderPolicy;
        Signatures = signatures.ToList();
        Findings = findings.ToList();
    }

    /// <summary>Whether the document carries at least one signature.</summary>
    public bool HasSignatures { get; }

    /// <summary>
    /// Whether every signature's covered content matches what was signed. See the remarks on this
    /// type before reading <see langword="false"/> as "tampered" — it is also what an unsigned
    /// document reports.
    /// </summary>
    public bool IsCryptographicallyValid { get; }

    /// <summary>Whether every signature passed every check this call actually performed.</summary>
    public bool IsValidUnderPolicy { get; }

    /// <summary>One result per signature the document carries.</summary>
    public IReadOnlyList<DocumentSignatureValidationResult> Signatures { get; }

    /// <summary>Human-readable diagnostic messages from the underlying validator.</summary>
    public IReadOnlyList<string> Findings { get; }
}
