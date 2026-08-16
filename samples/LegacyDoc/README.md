# LegacyDoc

Reading the Word 97-2003 binary `.doc` files still sitting on an old share drive, and converting
them to `.docx`.

```bash
dotnet run --project samples/LegacyDoc
```

Reads a real `.doc` saved by Word, prints its text including the table cells, shows the conversion
**refusing** by default, then accepts the loss deliberately and prints exactly what was dropped.
Writes `quarterly-report.docx` next to the built binary.

## The non-obvious part

**Converting refuses by default, and that will be the common case rather than a rare one.**

A legacy `.doc` keeps pictures, drawings and form fields in a binary stream that a `.docx` cannot
carry. Rather than hand back a document quietly missing them, `Convert` throws. Measured: **any
`.doc` containing a table has such a stream** — plain text, bold runs and headings do not. Tables are
ordinary, so expect to meet this on real files rather than on unusual ones.

```csharp
DocToDocxConverter.Convert(doc);                                   // throws
DocToDocxConverter.Convert(doc, new LegacyDocOptions { AllowContentLoss = true });   // converts
```

`ConvertWithReport` returns the same bytes plus a list of what was dropped, so the loss is
**recorded** rather than merely permitted. This sample prints that list — one entry for the binary
payload, and one noting that quick-save revision history is readable but not carried across as
editable revisions.

**What survives is more than the refusal implies**: text, tables with every cell, and character
formatting all convert intact. What is lost is the unprojected binary payload, and nothing else.

**`ExtractText` takes no options and never refuses**, because what that stream holds is not text.
Reading a `.doc` someone sent you is the common case and needs no policy decision.

**Reading only.** There is no `.doc` *writing* and there will not be — the underlying library
reports native `.doc` saving as unsupported, so offering it would mean claiming something that does
not work.

## Why this sample carries a fixture and the others do not

Every other sample here builds its own input with the library. This one cannot: a `.doc` cannot be
produced by a library that deliberately does not write the format. `assets/quarterly-report.doc` is
therefore committed — this sample's own copy, not a link into the test project, because a sample
reaching into `tests/` made a moved fixture surface as an opaque MSBuild error once already.
