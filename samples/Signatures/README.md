# Signatures

Inspecting and validating digital signatures on DOCX, XLSX and PPTX — and the trap in reading
either result too literally.

```bash
dotnet run --project samples/Signatures
```

Signs a Word document with a self-signed certificate, then inspects and validates it, then
tampers with the content after signing and validates that copy too. Needs no fixture: the
unsigned document is produced by the library first, and signing goes through
`OfficeIMO.Word.WordDocument.SignPackage` directly — this library validates and inspects
signatures but does not create one.

## The non-obvious part

**A self-signed certificate is not a broken signature.** `ValidateSignatures` reports
`CryptographicStatus = Passed` (the signature itself is well-formed) alongside
`CertificateChainStatus = Failed` (nobody told this machine to trust the issuer) on the exact same
signature. That is normal for an internal signer, not a defect — the two are independent findings
on purpose. `ValidateCertificateTrust = false` skips the chain check entirely, for exactly that
case.

**`CryptographicStatus` is not tamper detection, despite reading like it should be.** This sample
tampers with the signed document's content — no re-signing — and validates both copies. The
per-signature `CryptographicStatus` stays `Passed` on **both**: it only confirms the signature
block itself is well-formed against its own `SignedInfo`, not that the content it covers is
unchanged. `report.IsCryptographicallyValid` is the field that actually goes `false` on the
tampered copy — read that one, not the per-signature field, if the question is "was this altered
after signing".

**`InspectSignatures` reports a claimed identity, not a proven one.** It reads a signer's name
straight from their certificate's subject — the same way the certificate itself claims an
identity. Anyone can put any name on a self-signed certificate. Use `ValidateSignatures` before
treating a signer's name as real.
