namespace DocToolkit;

/// <summary>
/// One dimension of a signature's validated state — mirrors OfficeIMO's own
/// <c>OfficePackageSignatureValidationState</c> 1:1, so nothing here can drift from what the
/// validator beneath actually reports.
/// </summary>
public enum DocumentSignatureStatus
{
    /// <summary>There was nothing to check — the document carries no signature at all.</summary>
    NotPresent,

    /// <summary>This dimension was not evaluated for this call — see the options that were passed.</summary>
    NotChecked,

    /// <summary>The check passed.</summary>
    Passed,

    /// <summary>The check failed.</summary>
    Failed,

    /// <summary>The signature uses a construct this validator does not evaluate.</summary>
    Unsupported,
}
