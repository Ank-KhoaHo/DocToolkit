using OfficeIMO.Security;

namespace DocToolkit;

/// <summary>
/// The one place a document's signature state is inspected or validated, and the one place
/// OfficeIMO's own result types are mapped into this package's — so DocxEditor, WorkbookEditor
/// and PresentationEditor cannot disagree about what a status means.
/// </summary>
/// <remarks>
/// <b>OfficeIMO's signature API takes a file PATH, not bytes or a Stream — no such overload
/// exists.</b> XML digital-signature verification is byte-sensitive, so re-serializing through an
/// in-memory object model risks validating bytes that differ from what was actually signed. This
/// writes the caller's bytes to a uniquely named temporary file (created user-only via
/// <see cref="UnixFileMode"/> — this is the only place this package writes caller document
/// content into a shared, non-per-user directory), runs the given OfficeIMO call against it, and
/// deletes it in a <c>finally</c> block — the one place this package needs that pattern; every
/// other OfficeIMO interaction goes through a <c>Stream</c>-based <c>Load</c>.
///
/// <b>No revocation check and no certificate download ever happen, on any call.</b> Both are
/// forced off before every validation, regardless of what a caller might want — see
/// <see cref="DocumentSignatureValidationOptions"/>'s own remarks for why this is not a caller
/// choice.
/// </remarks>
internal static class OfficeSignature
{
    internal static DocumentSignatureInfo Inspect(
        byte[] document, string fileExtension,
        Func<string, OfficePackageSignatureInspectionOptions, OfficePackageSignatureInfo> inspectPackageSignatures,
        string format)
    {
        var tempPath = TempFilePath(fileExtension);
        try
        {
            WriteToTempFile(tempPath, document);
            var info = inspectPackageSignatures(tempPath, new OfficePackageSignatureInspectionOptions());
            var signers = info.SignatureParts.SelectMany(p => p.X509SubjectNames).ToList();
            return new DocumentSignatureInfo(info.HasSignatures, info.SignatureParts.Count, signers);
        }
        catch (Exception ex) when (ex is not DocumentConversionException and not ArgumentException)
        {
            throw new DocumentConversionException(
                $"Failed to inspect {format} signatures. See the inner exception for details.", ex);
        }
        finally
        {
            DeleteTempFile(tempPath);
        }
    }

    internal static DocumentSignatureValidationReport Validate(
        byte[] document, string fileExtension,
        Func<string, IOfficeSecurityProvider, OfficePackageSignatureValidationOptions, OfficePackageSignatureValidationReport> validatePackageSignatures,
        DocumentSignatureValidationOptions? options, string format)
    {
        var resolvedOptions = options ?? new DocumentSignatureValidationOptions();
        var tempPath = TempFilePath(fileExtension);
        try
        {
            WriteToTempFile(tempPath, document);
            var officeOptions = new OfficeIMO.Security.OfficePackageSignatureValidationOptions
            {
                ValidateCertificateTrust = resolvedOptions.ValidateCertificateTrust,
            };
            // Forced off unconditionally - see this type's own remarks and
            // DocumentSignatureValidationOptions' remarks for why this is not a caller choice.
            officeOptions.CertificateValidation.RevocationMode = System.Security.Cryptography.X509Certificates.X509RevocationMode.NoCheck;
            officeOptions.CertificateValidation.DisableCertificateDownloads = true;

            var report = validatePackageSignatures(tempPath, OfficeSecurityProvider.Default, officeOptions);
            var signatures = report.Signatures.Select(ToValidationResult).ToList();
            return new DocumentSignatureValidationReport(
                report.SignatureInfo.HasSignatures, report.IsCryptographicallyValid, report.IsValidUnderPolicy,
                signatures, report.Findings);
        }
        catch (Exception ex) when (ex is not DocumentConversionException and not ArgumentException)
        {
            throw new DocumentConversionException(
                $"Failed to validate {format} signatures. See the inner exception for details.", ex);
        }
        finally
        {
            DeleteTempFile(tempPath);
        }
    }

    private static DocumentSignatureValidationResult ToValidationResult(
        OfficePackageSignaturePartValidationResult part)
    {
        return new DocumentSignatureValidationResult(
            ToStatus(part.CryptographicStatus),
            ToStatus(part.CertificateChainStatus),
            ToStatus(part.RevocationStatus),
            part.SignaturePart.X509SubjectNames);
    }

    private static DocumentSignatureStatus ToStatus(OfficePackageSignatureValidationState state) => state switch
    {
        OfficePackageSignatureValidationState.NotPresent => DocumentSignatureStatus.NotPresent,
        OfficePackageSignatureValidationState.NotChecked => DocumentSignatureStatus.NotChecked,
        OfficePackageSignatureValidationState.Passed => DocumentSignatureStatus.Passed,
        OfficePackageSignatureValidationState.Failed => DocumentSignatureStatus.Failed,
        OfficePackageSignatureValidationState.Unsupported => DocumentSignatureStatus.Unsupported,
        _ => throw new ArgumentOutOfRangeException(nameof(state), state, "Unrecognized OfficeIMO signature validation state."),
    };

    private static string TempFilePath(string fileExtension) =>
        Path.Combine(Path.GetTempPath(), $"doctoolkit-signature-{Guid.NewGuid():N}{fileExtension}");

    /// <summary>
    /// Writes <paramref name="document"/> to <paramref name="path"/> with the file created
    /// user-only on Unix-like systems (<see cref="UnixFileMode.UserRead"/> |
    /// <see cref="UnixFileMode.UserWrite"/>), rather than the shared temp directory's default
    /// permissions. This is the only place this package writes caller document content to disk
    /// under a shared, non-per-user directory (<c>/tmp</c> on Linux/macOS is world-readable).
    /// </summary>
    /// <remarks>
    /// <b>Setting <see cref="FileStreamOptions.UnixCreateMode"/> is not a harmless no-op on
    /// Windows — it throws <see cref="PlatformNotSupportedException"/> there, measured directly.</b>
    /// <c>%TEMP%</c> is already per-user on Windows, so the guard below is not merely satisfying
    /// the platform-compatibility analyzer; setting it unconditionally would have broken every
    /// signature operation on the platform this was built and tested on.
    /// </remarks>
    private static void WriteToTempFile(string path, byte[] document)
    {
        var options = new FileStreamOptions
        {
            Mode = FileMode.CreateNew,
            Access = FileAccess.Write,
            Share = FileShare.None,
        };
        if (!OperatingSystem.IsWindows())
        {
            options.UnixCreateMode = UnixFileMode.UserRead | UnixFileMode.UserWrite;
        }
        using var stream = new FileStream(path, options);
        stream.Write(document);
    }

    private static void DeleteTempFile(string path)
    {
        try { File.Delete(path); }
        catch (IOException) { /* best effort */ }
        catch (UnauthorizedAccessException) { /* best effort - e.g. a locked handle from an AV scanner */ }
    }
}
