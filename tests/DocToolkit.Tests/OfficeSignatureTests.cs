using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using OfficeIMO.Word;

namespace DocToolkit.Tests;

public class OfficeSignatureTests
{
    private static X509Certificate2 SelfSignedCertificate()
    {
        using var rsa = RSA.Create(2048);
        var request = new CertificateRequest(
            "CN=DocToolkit Test Signer", rsa, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
        using var cert = request.CreateSelfSigned(
            DateTimeOffset.UtcNow.AddDays(-1), DateTimeOffset.UtcNow.AddDays(30));
        var pfxBytes = cert.Export(X509ContentType.Pfx, "pw");
        return X509CertificateLoader.LoadPkcs12(pfxBytes, "pw", X509KeyStorageFlags.Exportable);
    }

    /// <summary>
    /// A signed DOCX, built the same way OfficeIMO itself authors one — WordDocument.SignPackage
    /// requires a real file path, so this writes to a TempFile the same way OfficeSignature itself
    /// does, matching this repo's "author fixtures the way the library under test authors them"
    /// discipline.
    /// </summary>
    private static byte[] SignedDocx(X509Certificate2 cert)
    {
        using var input = new TempFile();
        using (var doc = WordDocument.Create(input.Path))
        {
            doc.AddParagraph("Hello, this is a signed document.");
            doc.Save();
        }
        WordDocument.SignPackage(input.Path, OfficeIMO.Security.OfficeSecurityProvider.Default, cert,
            new WordPackageSigningOptions());
        return File.ReadAllBytes(input.Path);
    }

    [Fact]
    public void Inspect_ReportsAnUnsignedDocumentCleanly()
    {
        using var input = new TempFile();
        using (var doc = WordDocument.Create(input.Path))
        {
            doc.AddParagraph("Unsigned.");
            doc.Save();
        }
        var bytes = File.ReadAllBytes(input.Path);

        var info = OfficeSignature.Inspect(
            bytes, ".docx", WordDocument.InspectPackageSignatures, "DOCX");

        Assert.False(info.HasSignatures);
        Assert.Equal(0, info.SignatureCount);
        Assert.Empty(info.Signers);
    }

    [Fact]
    public void Inspect_ReportsTheSignerOfASignedDocument()
    {
        using var cert = SelfSignedCertificate();
        var signed = SignedDocx(cert);

        var info = OfficeSignature.Inspect(
            signed, ".docx", WordDocument.InspectPackageSignatures, "DOCX");

        Assert.True(info.HasSignatures);
        Assert.Equal(1, info.SignatureCount);
        Assert.Contains(info.Signers, s => s.Contains("DocToolkit Test Signer"));
    }

    [Fact]
    public void Validate_SeparatesCryptographicIntegrityFromCertificateChainTrust()
    {
        // The ticket's own central question, measured directly: a self-signed certificate never
        // chains to a trusted root, but that must not make an untampered signature report as
        // cryptographically invalid - the two are independent findings.
        using var cert = SelfSignedCertificate();
        var signed = SignedDocx(cert);

        var report = OfficeSignature.Validate(
            signed, ".docx", WordDocument.ValidatePackageSignatures,
            new DocumentSignatureValidationOptions(), "DOCX");

        Assert.True(report.HasSignatures);
        Assert.True(report.IsCryptographicallyValid);
        var signature = Assert.Single(report.Signatures);
        // signature.CryptographicStatus does NOT discriminate tamper (see
        // Validate_RejectsATamperedDocument and DocumentSignatureValidationResult's own remarks) -
        // it is asserted here only to pin that an untrusted chain doesn't drag it to Failed either,
        // i.e. it really is independent of CertificateChainStatus in both directions.
        Assert.Equal(DocumentSignatureStatus.Passed, signature.CryptographicStatus);
        Assert.Equal(DocumentSignatureStatus.Failed, signature.CertificateChainStatus);
        // The untrusted root correctly fails the overall policy verdict even though the content
        // itself was never tampered with.
        Assert.False(report.IsValidUnderPolicy);
    }

    [Fact]
    public void Validate_WithTrustCheckDisabled_IgnoresTheUntrustedChain()
    {
        using var cert = SelfSignedCertificate();
        var signed = SignedDocx(cert);

        var report = OfficeSignature.Validate(
            signed, ".docx", WordDocument.ValidatePackageSignatures,
            new DocumentSignatureValidationOptions { ValidateCertificateTrust = false }, "DOCX");

        Assert.True(report.IsCryptographicallyValid);
        Assert.True(report.IsValidUnderPolicy);
    }

    [Fact]
    public void Validate_RejectsATamperedDocument()
    {
        using var cert = SelfSignedCertificate();
        var signed = SignedDocx(cert);

        // Tamper AFTER signing, without re-signing - exactly the scenario a signature exists to
        // catch. Uses the raw OpenXml SDK directly, matching this session's own probe.
        using var tampered = new TempFile();
        File.WriteAllBytes(tampered.Path, signed);
        using (var pkg = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(tampered.Path, true))
        {
            var text = pkg.MainDocumentPart!.Document!.Body!.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().First();
            text.Text = "TAMPERED CONTENT - changed after signing";
            pkg.MainDocumentPart.Document.Save();
        }
        var tamperedBytes = File.ReadAllBytes(tampered.Path);

        var untamperedReport = OfficeSignature.Validate(
            signed, ".docx", WordDocument.ValidatePackageSignatures,
            new DocumentSignatureValidationOptions { ValidateCertificateTrust = false }, "DOCX");
        var tamperedReport = OfficeSignature.Validate(
            tamperedBytes, ".docx", WordDocument.ValidatePackageSignatures,
            new DocumentSignatureValidationOptions { ValidateCertificateTrust = false }, "DOCX");

        Assert.True(untamperedReport.IsCryptographicallyValid);
        Assert.False(tamperedReport.IsCryptographicallyValid);
        Assert.False(tamperedReport.IsValidUnderPolicy);

        // PINS A MEASURED, DOCUMENTED LIMITATION - see DocumentSignatureValidationResult's own
        // remarks. The per-signature CryptographicStatus does NOT detect this tamper: it verifies
        // only that the signature block's own SignatureValue is well-formed against SignedInfo,
        // which a post-signing content edit does not touch. Confirmed unchanged on BOTH copies -
        // report.IsCryptographicallyValid above is the field that actually discriminates. If a
        // future OfficeIMO version starts reflecting the digest mismatch here, this assertion
        // fails and is the signal to revisit the doc comments rather than a silent behavior change.
        var untamperedSignature = Assert.Single(untamperedReport.Signatures);
        var tamperedSignature = Assert.Single(tamperedReport.Signatures);
        Assert.Equal(DocumentSignatureStatus.Passed, untamperedSignature.CryptographicStatus);
        Assert.Equal(DocumentSignatureStatus.Passed, tamperedSignature.CryptographicStatus);
    }

    [Fact]
    public void Validate_NeverEnablesRevocationCheckingOrCertificateDownloads()
    {
        // A caller cannot even ask for this - DocumentSignatureValidationOptions has no such
        // property - so this proves the INTERNAL options OfficeSignature builds are also always
        // forced off, regardless of what a future maintainer might otherwise wire through.
        using var cert = SelfSignedCertificate();
        var signed = SignedDocx(cert);

        var report = OfficeSignature.Validate(
            signed, ".docx", WordDocument.ValidatePackageSignatures,
            new DocumentSignatureValidationOptions(), "DOCX");

        // RevocationStatus must never be anything that implies an online check was attempted -
        // NotChecked is the only value a purely local validation can produce.
        var signature = Assert.Single(report.Signatures);
        Assert.Equal(DocumentSignatureStatus.NotChecked, signature.RevocationStatus);
    }

    [Fact]
    public void DocumentSignatureStatus_MirrorsOfficePackageSignatureValidationStateByName()
    {
        Assert.Equal(
            Enum.GetNames<OfficeIMO.Security.OfficePackageSignatureValidationState>().OrderBy(n => n, StringComparer.Ordinal),
            Enum.GetNames<DocumentSignatureStatus>().OrderBy(n => n, StringComparer.Ordinal));
    }
}
