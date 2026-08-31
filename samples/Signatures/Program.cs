using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using DocToolkit;
using OfficeIMO.Word;

Console.WriteLine("Signatures");
Console.WriteLine("==========");

// This library reads and validates digital signatures but does not create one itself - signing
// goes through OfficeIMO.Word.WordDocument.SignPackage directly, the library
// InspectSignatures/ValidateSignatures are themselves built on. SignPackage needs a real file
// path - XML digital-signature verification is byte-sensitive, so OfficeIMO offers no byte[]/
// Stream form of it - so this sample writes to a temp file the same way the library's own
// internals do (see OfficeSignature.cs's own remarks on why).

using var rsa = RSA.Create(2048);
var request = new CertificateRequest(
    "CN=Sample Signer", rsa, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
using var selfSigned = request.CreateSelfSigned(
    DateTimeOffset.UtcNow.AddDays(-1), DateTimeOffset.UtcNow.AddDays(30));
using var cert = X509CertificateLoader.LoadPkcs12(
    selfSigned.Export(X509ContentType.Pfx, "pw"), "pw", X509KeyStorageFlags.Exportable);

byte[] unsigned = DocxEditor.Create(new[] { DocxBlock.Paragraph("Q3 board resolution") });

string tempPath = Path.Join(Path.GetTempPath(), $"doctoolkit-sample-{Guid.NewGuid():N}.docx");
byte[] signed;
try
{
    File.WriteAllBytes(tempPath, unsigned);
    WordDocument.SignPackage(
        tempPath, OfficeIMO.Security.OfficeSecurityProvider.Default, cert, new WordPackageSigningOptions());
    signed = File.ReadAllBytes(tempPath);
}
finally
{
    // Best effort: a delete failure (e.g. a locked handle from an AV scanner) must not mask
    // whatever exception the try block itself raised.
    try { File.Delete(tempPath); } catch (IOException) { } catch (UnauthorizedAccessException) { }
}

// ---------------------------------------------------------------------------------------------
Console.WriteLine("\n1. Inspecting - who claims to have signed it, unvalidated");

DocumentSignatureInfo unsignedInfo = DocxEditor.InspectSignatures(unsigned);
DocumentSignatureInfo signedInfo = DocxEditor.InspectSignatures(signed);

Console.WriteLine($"   unsigned doc : HasSignatures={unsignedInfo.HasSignatures}");
Console.WriteLine($"   signed doc   : HasSignatures={signedInfo.HasSignatures}, "
                  + $"signer(s)=\"{string.Join(", ", signedInfo.Signers)}\"");
Console.WriteLine("   ^ InspectSignatures reports a CLAIMED identity - it does not check anything "
                  + "cryptographically. Use ValidateSignatures before trusting a name.");

// ---------------------------------------------------------------------------------------------
Console.WriteLine("\n2. The trap: a self-signed certificate is not a broken signature");

DocumentSignatureValidationReport report = DocxEditor.ValidateSignatures(signed);
DocumentSignatureValidationResult signature = report.Signatures[0];

Console.WriteLine($"   report.IsCryptographicallyValid : {report.IsCryptographicallyValid}  (content matches what was signed)");
Console.WriteLine($"   signature.CertificateChainStatus: {signature.CertificateChainStatus}  (nobody told this machine to trust \"Sample Signer\")");
Console.WriteLine($"   report.IsValidUnderPolicy       : {report.IsValidUnderPolicy}");
Console.WriteLine("   ^ these are independent findings on purpose. A self-signed certificate is");
Console.WriteLine("     normal for an internal signer - it does not mean the content is untrustworthy,");
Console.WriteLine("     only that this machine has not been told to trust that certificate's issuer.");

DocumentSignatureValidationReport withTrustOff = DocxEditor.ValidateSignatures(
    signed, new DocumentSignatureValidationOptions { ValidateCertificateTrust = false });
Console.WriteLine("\n   With ValidateCertificateTrust = false:");
Console.WriteLine($"   report.IsValidUnderPolicy       : {withTrustOff.IsValidUnderPolicy}  (chain check skipped entirely)");

// ---------------------------------------------------------------------------------------------
Console.WriteLine("\n3. What tampering actually looks like");

string tamperedPath = Path.Join(Path.GetTempPath(), $"doctoolkit-sample-{Guid.NewGuid():N}.docx");
byte[] tampered;
try
{
    File.WriteAllBytes(tamperedPath, signed);
    using (var pkg = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(tamperedPath, true))
    {
        var text = pkg.MainDocumentPart!.Document!.Body!
            .Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().First();
        text.Text = "Someone changed this after it was signed";
        pkg.MainDocumentPart.Document.Save();
    }
    tampered = File.ReadAllBytes(tamperedPath);
}
finally
{
    try { File.Delete(tamperedPath); } catch (IOException) { } catch (UnauthorizedAccessException) { }
}

DocumentSignatureValidationReport tamperedReport = DocxEditor.ValidateSignatures(
    tampered, new DocumentSignatureValidationOptions { ValidateCertificateTrust = false });
DocumentSignatureValidationResult tamperedSignature = tamperedReport.Signatures[0];

Console.WriteLine($"   report.IsCryptographicallyValid : {tamperedReport.IsCryptographicallyValid}  (this is the field that catches it)");
Console.WriteLine($"   signature.CryptographicStatus   : {tamperedSignature.CryptographicStatus}  (unchanged - see below)");
Console.WriteLine("   ^ the per-signature CryptographicStatus only checks that the signature block");
Console.WriteLine("     itself is well-formed against its own SignedInfo - it does NOT re-check the");
Console.WriteLine("     content, so it reads Passed on the tampered copy too. Read");
Console.WriteLine("     report.IsCryptographicallyValid for tamper detection, not the per-signature field.");

Console.WriteLine("\n4. The same four members exist on WorkbookEditor and PresentationEditor");
byte[] xlsx = WorkbookEditor.Create("Sheet1", new object?[][] { new object?[] { "x" } });
byte[] pptx = PresentationEditor.Create(new[] { PptxSlide.Titled("Untitled") });
Console.WriteLine($"   WorkbookEditor.InspectSignatures(xlsx).HasSignatures     : {WorkbookEditor.InspectSignatures(xlsx).HasSignatures}");
Console.WriteLine($"   PresentationEditor.InspectSignatures(pptx).HasSignatures : {PresentationEditor.InspectSignatures(pptx).HasSignatures}");

Console.WriteLine("\nDone.");
