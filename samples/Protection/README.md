# Protection

Putting a password on a PDF, DOCX, XLSX or PPTX — and finding out why a "protected" file is still
opening for everyone.

```bash
dotnet run --project samples/Protection
```

Encrypts a PDF and proves it cannot be read back, demonstrates that an **owner** password is not a
lock by reading the text out of one with no password at all, then does the same three operations on
a workbook and shows the two failure modes that are worth telling apart. Writes `protected.pdf` next
to the built binary — it opens with the password `s3cret`.

Needs no fixture: every document it protects is produced by the library first.

## The non-obvious part

**A user password and an owner password are not two strengths of the same thing.**

| | enforced by | what it stops |
|---|---|---|
| `UserPassword` | cryptography | opening the file at all |
| `OwnerPassword` | convention | nothing, on its own |

An owner password sets the permission flags and asks a reader to honour them. The file still opens
for anyone — a cooperative reader greys out printing, an uncooperative one need not. **This sample
prints the text of an owner-protected PDF with no password supplied**, because that is far more
convincing than a sentence saying it would.

That is the PDF specification working as designed, not a defect here or in any reader. **If the
content must not be read, set `UserPassword`.**

Two smaller things the sample shows rather than tells:

- **`Unprotect` needs the OWNER password when the document has one**, even if you also know the user
  password. Removing protection is a modification, and the format reserves that for the owner.
- **An encrypted Office file is not a package any more.** A plain `.docx`/`.xlsx`/`.pptx` is a ZIP —
  the encrypted form is a compound file with the ZIP sealed inside, which is why the output starts
  `ÐÏ` rather than `PK`, and why every other method on those classes refuses it until you
  `Unprotect`. `IsProtected` answers "would they refuse this?" from the signature alone, so it needs
  no password.

Finally, a wrong password and a file that was never encrypted are reported as **different**
failures. Only one of them is fixed by asking the user to type it again.
