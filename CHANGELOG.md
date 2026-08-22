# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/).

Ank.DocToolkit and Ank.DocToolkit.Extensions.DependencyInjection release together, at the same
version, from a single tag (see README.md > Releasing). Entries below are prefixed **Core:** or
**Extensions:** when they apply to only one package; unprefixed entries apply to both or to
repo-wide tooling (CI, release pipeline).

## [0.33.4](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.33.3...v0.33.4) (2026-08-22)


### Fixed

* **Core: a cancellation token now reaches the PDF render, not only the HTML-to-DOCX stage.**
  Every HTML-to-PDF overload documents an `OperationCanceledException`. Until now the render -
  the slower half, which a repair can run more than once - could not be interrupted once it had
  started ([#333](https://github.com/Ank-KhoaHo/DocToolkit/issues/333)) ([d4a058c](https://github.com/Ank-KhoaHo/DocToolkit/commit/d4a058cb8207abac35365be2efcce70b4f6c9586))


### Changed

* **Core: a file-path overload handed a ZERO-BYTE FILE now names the parameter you passed.**
  It raises `ArgumentException` with `ParamName` `path` or `inputPath`, rather than the one
  belonging to the `byte[]` method underneath (`docx`, `xlsx`, `pdf`, and so on). Nineteen
  overloads across seven types behaved the old way, and each one's documented `<exception>` tag
  already described the new behaviour.

  The exception **type** is unchanged, so `catch (ArgumentException)` needs no edit. **Code that
  switches on `ParamName`, or matches the message text, will need updating** - see *Migrating*
  in the package README ([#334](https://github.com/Ank-KhoaHo/DocToolkit/issues/334)) ([2b643e7](https://github.com/Ank-KhoaHo/DocToolkit/commit/2b643e70f71150297275303a36382d45fc0a4795))
* **Core: HTML with many in-page links converts faster.** Link targets are resolved in one pass
  rather than one document query per link. Measured 258 ms to 20 ms on a page of 2000 blocks and
  300 links, of which 14 ms is parsing the HTML ([#336](https://github.com/Ank-KhoaHo/DocToolkit/issues/336)) ([fd3e52e](https://github.com/Ank-KhoaHo/DocToolkit/commit/fd3e52ec0069017cf4abffc3db8a5034790d0883))
* **Core: two internal paths stop doing work nothing uses** - a DOCX-to-PDF retry that could not
  have helped is no longer attempted, and opening a workbook no longer copies it ([#338](https://github.com/Ank-KhoaHo/DocToolkit/issues/338)) ([8604d21](https://github.com/Ank-KhoaHo/DocToolkit/commit/8604d212efa175b0da0572852c2bac63c75ebb68))

## [0.33.3](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.33.2...v0.33.3) (2026-08-21)


### Changed

* **Core: a replaced image is stored once, however many times its placeholder appears.**
  `DocxEditor.ReplaceImage` added a fresh copy of the image bytes for every occurrence, so a
  logo repeated down a document was embedded once per match.

  Measured with a 40 KB image: three occurrences produced three byte-identical media parts and a
  **122,483-byte** package. The same document is now **42,026 bytes**. The saving grows with the
  number of occurrences and with the size of the image.

  **Nothing changes about where the parts live.** An image still belongs to the container that
  owns its paragraph - a header image to the header part - because a relationship id resolving in
  the wrong scope opens in Word showing nothing. Only duplication *within* one container is
  removed. Each occurrence keeps its own drawing id, which is what Word requires.
  ([#326](https://github.com/Ank-KhoaHo/DocToolkit/issues/326))

## [0.33.2](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.33.1...v0.33.2) (2026-08-21)


### Fixed

* **Core: JPEG images with fill bytes are no longer refused.** JPEG permits any number of 0xFF
  padding bytes before a marker, and encoders emit them when aligning to a boundary. The size
  scanner treated one as a marker, misread the next real marker as a segment length, and skipped
  the frame header - so a **valid** image could not be embedded at all.

  Measured: **one fill byte was enough**. Affected images failed in `DocxBlock.Image` and
  `DocxEditor.ReplaceImage` with a message blaming the file - that it "is truncated, or not
  actually a well-formed JPEG" - which sent you to check two things that were both fine.

  Every `.docx` image needs an explicit size and this package deliberately carries no
  image-decoding dependency, so that scanner is the only thing that can supply it; there was no
  workaround short of re-encoding the image.
  ([#324](https://github.com/Ank-KhoaHo/DocToolkit/issues/324))

## [0.33.1](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.33.0...v0.33.1) (2026-08-21)


### Fixed

* **Core: the `Stream` overloads now convert everything the `byte[]` ones do.** If you write PDFs
  to a stream - an HTTP response body, a file, a pipe - this is worth reading, because those
  overloads were refusing documents their array siblings converted.

  They wrote the PDF straight through as the renderer produced it, which meant they could not
  apply the repairs that retry a failed render. **Measured over real files: 4 of 99 real Word
  documents converted through `DocxToPdfConverter.Convert` and were refused through its stream
  overload**, and on the HTML path a construct present in 27 of 181 real `.gov` pages did the
  same. `AddDocToolkit`'s `IHtmlToPdfConverter` routes to the stream path, so ASP.NET callers
  streaming to the response body had the worse behaviour by default.

  **What changes for you.** Affected documents now convert. The PDF is rendered whole and then
  written, so a failure leaves your destination **untouched** rather than carrying a truncated
  PDF - previously documented as the expected behaviour and now no longer the case. Peak memory
  rises by one rendered PDF; this was never a memory optimisation, since the source was already
  buffered either way.

  Streams are still never disposed, closed or sought, and may still be forward-only.
  ([#322](https://github.com/Ank-KhoaHo/DocToolkit/issues/322))

* **Core: `XlsxToPdfConverter.ConvertAsync` refuses a legacy `.xls` workbook**, as its `byte[]`
  sibling already did. A binary Excel 97-2003 file could previously reach the renderer through the
  stream overload and consume minutes of CPU on input the API never claimed to support - which
  matters most for the upload endpoints that overload exists for.

## [0.33.0](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.32.0...v0.33.0) (2026-08-20)


### Added

* **Extensions: configure PDF fonts once, on `DocToolkitOptions`.** The `PdfFontOptions` added in
  0.32.0 could only be passed per call, which is the wrong shape for dependency injection - needing
  a font is a property of the deployment, not of the document.

  ```csharp
  services.AddDocToolkit(o =>
      o.Fonts = new PdfFontOptions("Noto Sans", File.ReadAllBytes("NotoSans-Regular.ttf")));
  ```

  Configure it once at registration and every resolved converter uses it. Nothing is fetched and
  nothing is read from disk - the bytes come from you. ([#312](https://github.com/Ank-KhoaHo/DocToolkit/issues/312))

  **Known limitation:** this reaches `IDocxToPdfConverter` but not yet `IHtmlToPdfConverter`, whose
  service composes page setup and remote-image settings and has no core overload taking all three.
  Callers needing fonts on the HTML path can use the static `HtmlToPdfConverter.ConvertAsync`
  overload directly.

* **Core: legacy PowerPoint 97-2003 `.ppt` to PDF is now a supported, tested capability.**
  `PptxToPdfConverter.Convert` already read binary decks; that was inherited from a dependency,
  undocumented and uncovered. **The behaviour is unchanged** - what is new is that it is claimed,
  pinned by tests against a real deck, and published with its rate.

  **It succeeds on 60.2%** of 88 real decks from a government crawl - a lower bar than the OOXML
  path, stated rather than rounded up. Twenty of the thirty-five refusals are a single upstream
  limitation; none produced a corrupt PDF. ([#319](https://github.com/Ank-KhoaHo/DocToolkit/issues/319))

### Fixed

* **Core: passing something that is not HTML to `HtmlToDocxConverter` now says so.** The refusal
  was correct and unreadable - *"See the inner exception for details"*, wrapping *"hexadecimal
  value 0x10, is an invalid character"*, which is a message about a character you never typed.

  It now names the character and both causes it cannot distinguish between: content that is not
  HTML at all (an image, a PDF, an Office file - each has its own reader here), or genuine HTML
  carrying a stray control character. Ordinary markup is unaffected: tabs, newlines, character
  entities, accented text, CJK and emoji all convert. ([#315](https://github.com/Ank-KhoaHo/DocToolkit/issues/315))

## [0.32.0](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.31.1...v0.32.0) (2026-08-19)


**Both entries below are about the same thing: rendering documents whose text or layout the PDF
renderer would otherwise refuse.** Measured over 99 documents carrying real content, DOCX → PDF went
from **71.7% to 75.8%**, and to **77.8%** when the caller supplies fonts.

### Added

* **Core: supply your own fonts for characters the renderer cannot otherwise encode.** Whether a
  document containing Cyrillic, Greek or CJK renders has been a property of the **machine** — the
  renderer falls back to whatever fonts the host happens to have, so the same document converts on
  one and is refused on another.

  ```csharp
  var fonts = new PdfFontOptions("Noto Sans", File.ReadAllBytes("NotoSans-Regular.ttf"));
  byte[] pdf = DocxToPdfConverter.Convert(docx, fonts);
  byte[] web = await HtmlToPdfConverter.ConvertAsync(html, fonts);
  ```

  Nothing is fetched and nothing is read from disk by this library — the bytes come from you, so it
  works air-gapped. **No font ships inside the package**: one covering Cyrillic, Greek and CJK is
  measured in megabytes against a package measured in tens of kilobytes.

  **Read this before using it — supplying too few fonts is worse than supplying none.** The fonts you
  pass **replace** the host's own fallbacks rather than adding to them. Measured over 99 documents:
  none → 71/99, **one font → 63/99**, four fonts → 77/99. One font fixed the four documents needing
  Cyrillic and broke twelve the host had been covering. Supply fonts covering everything your
  documents use, not only the script that failed; the refusal names the character still missing.
  It also changes how fonts are embedded generally — an ordinary Latin document went from 128,755
  bytes to 1,306, both rendering correctly
  ([#307](https://github.com/Ank-KhoaHo/DocToolkit/issues/307),
  [#309](https://github.com/Ank-KhoaHo/DocToolkit/issues/309)).

* **Core: a document whose paragraphs have a negative indent now renders.** The PDF renderer refuses
  any negative left or right paragraph indent at any magnitude — a 0.35pt value fails exactly as a
  36pt one does — while Word honours it, which is how content is deliberately set outside the margin
  in a letterhead or pull-quote.

  **This one changes the layout slightly**, and is the only repair in this package that does: the
  indent is clamped to zero, so content that overhung the margin now stops at it. Ordinary hanging
  indents are untouched. Recovered 4 documents in 99
  ([#310](https://github.com/Ank-KhoaHo/DocToolkit/issues/310)).

## [0.31.1](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.31.0...v0.31.1) (2026-08-19)


**A patch release with one change in it.** Three commits appear below; two of them are a change and
its revert, and net to nothing.

### Fixed

* **Core: the two commonest reasons a Word document will not render to PDF now name themselves.**
  Measured over 99 documents carrying real content, DOCX → PDF succeeds on **71.7%**, and 15 of the
  28 failures are a negative paragraph indent (8) or header/footer content wider than the page (7).

  **Both are legal in Word**, which is what the messages now say. The renderer's own wording is
  accurate — *"Paragraph right indent must be a non-negative finite value"* — and leaves a reader
  hunting for a mistake in a document that does not contain one. Content set outside the margin, and
  a wide header, are ordinary in a letterhead.

  **An ordinary hanging indent is unaffected**, and the message says so: `w:hanging` and a negative
  `w:firstLine` both convert. Only a negative `w:left` or `w:right` is refused, at any magnitude
  ([#303](https://github.com/Ank-KhoaHo/DocToolkit/issues/303)).


### Changed

* **Nothing else behaves differently, and the other two entries are why that needs saying.**
  A resource policy was added to the DOCX → PDF path and reverted the same day
  ([#305](https://github.com/Ank-KhoaHo/DocToolkit/issues/305),
  [#306](https://github.com/Ank-KhoaHo/DocToolkit/issues/306)). It never reached a release.

  It is recorded rather than hidden because the measurement is worth having: assigning a resource
  policy to the Word renderer — **with any flag values, including permissive ones** — drops DOCX →
  PDF from 71/99 to 57/99 on real documents, because it stops the renderer resolving **fonts**. The
  XLSX and PPTX paths are unaffected and keep theirs.

  If you have been passing `WordPdfSaveOptions` with a `ResourcePolicy` in your own code, that is
  worth knowing: the flags are not what matters, the presence of the object is.

## [0.31.0](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.30.0...v0.31.0) (2026-08-18)


**Core: HTML → PDF now converts most real-world pages.** Measured over **181 genuine `.gov` pages**
from a public crawl, before and after:

| | before | after |
|---|---|---|
| HTML → PDF | 106 / 181 — **58.6%** | 159 / 181 — **87.8%** |
| HTML → DOCX | 163 / 181 — 90.1% | 177 / 181 — **97.8%** |

Every one of the five repairs below has the same shape: the page was **valid**, browsers render it
correctly, and the converter refused it over something the author had no reason to think was a
problem. **None of them can change a document that converts today** — each is applied only to input
that already fails, and the tests assert that by reference rather than by argument.

### Added

* **Core: a `rowspan` reaching past the last row of its table no longer fails the conversion.** It
  raised an unhandled index error naming nothing. Browsers clamp such a rowspan to the rows that
  exist, so this markup is common — the sample held spans of 2, 3, 14, 100 and 103 against tables of
  one to three rows — and it is now clamped the same way
  ([#293](https://github.com/Ank-KhoaHo/DocToolkit/issues/293)).

* **Core: a table cell holding a link that wraps only an image now renders.** A logo linking home, a
  "skip navigation" button. The link is labelled with the image's `alt` text, so the navigation
  survives; an image with no usable `alt` has the link unwrapped instead, because there is nothing to
  label it with ([#299](https://github.com/Ank-KhoaHo/DocToolkit/issues/299)).


### Fixed

* **Core: internal links whose targets use the obsolete `<a name="x">` form now resolve.** HTML5
  replaced `name` with `id` and browsers honour both, so these pages navigate correctly everywhere
  else while the converter produced a link to a bookmark it never created. The identity is moved to
  the nearest block that can carry one; a target that exists in no usable form has its link dropped
  and its text kept ([#296](https://github.com/Ank-KhoaHo/DocToolkit/issues/296),
  [#297](https://github.com/Ank-KhoaHo/DocToolkit/issues/297)).

* **Core: a table whose spacer cell collapses to nothing now renders.** An empty cell beside a cell
  of long text was given a near-zero width by automatic layout, which the renderer then refused —
  with no width specified anywhere in the document
  ([#298](https://github.com/Ank-KhoaHo/DocToolkit/issues/298)).

* **Core: the two things Markdown conversion rejects now say what they are.** A line feed written as
  `&#10;` raised a bare `NullReferenceException`, and an ordered list starting below 1 — which
  CommonMark permits — an `ArgumentOutOfRangeException`. Both messages now name the construct and the
  remedy. **The behaviour is unchanged and deliberately so**: both repairs would alter what the
  document says. Worth knowing that the ordered-list limit is **PDF-only** — `0. item` converts to
  DOCX perfectly well ([#301](https://github.com/Ank-KhoaHo/DocToolkit/issues/301)).


### Changed

* **Nothing that converted before behaves differently.** Stated explicitly because the last release
  did contain such a change: every repair here is reached only by input that already failed.

* **A failing HTML → PDF conversion can now take up to twice as long.** Two of the repairs cannot
  tell in advance whether a document needs them, so the render retries once with the repair that
  matches the failure raised. Only a conversion that already failed pays this; a successful one is
  unaffected.

* **`AngleSharp` is now a direct dependency.** It was already in the resolved graph — it is the
  parser the HTML converter uses underneath — so **nothing new ships**: the same 30 packages, the
  same licences. Noted for anyone auditing direct references.

## [0.30.0](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.29.0...v0.30.0) (2026-08-17)


### Added

* **Core: legacy PowerPoint `.ppt` renders to PDF, and is now claimed rather than accidental.**
  `PptxToPdfConverter` already read the binary format — it worked, was tested nowhere and was
  documented nowhere. Measured across 15 real `.ppt` files: **11 convert (73%)**, the slowest in
  1.7 s. It is now tested, documented and bounded. The editors are unchanged and still refuse
  `.ppt`: `PresentationEditor` is OOXML-only, and claiming the format on the PDF path does not
  claim it everywhere ([#290](https://github.com/Ank-KhoaHo/DocToolkit/issues/290)).

* **Core: Word documents containing bulleted lists now render to PDF.** They previously could not —
  at all. Word's default bullet is `U+F0B7`, a Symbol-font glyph in the Unicode private-use area,
  and the PDF renderer refused to encode it, so any `.docx` or `.doc` with a bulleted list failed.
  Since a table also carries one, that is most real documents
  ([#289](https://github.com/Ank-KhoaHo/DocToolkit/issues/289)).


### Changed

* **Core: `XlsxToPdfConverter` now REFUSES a legacy `.xls` immediately, where it used to attempt
  the render.** This is the one behaviour change in this release that can break working code. The
  renderer underneath does read the binary format, so a `.xls` sometimes succeeded on this one path
  while every other entry point refused it — but the cost was unbounded: measured, a 101 KB
  workbook took **10.9 s**, a 2.3 MB one did not finish in **ten minutes**, and a 7.7 MB one spent
  **161 s** before failing anyway. The supported `.xlsx` path renders 20,000 rows in 3.7 s, so this
  is the legacy path and not a property of large workbooks. Accepting caller-chosen input that
  costs minutes of CPU through an endpoint that never claimed the format is not a capability worth
  keeping.

  **Migrating:** if you were passing `.xls` bytes to `XlsxToPdfConverter`, you now get a
  `DocumentConversionException` at once, naming the format and telling you to save as `.xlsx`. The
  same message covers an encrypted `.xlsx` — both are compound files — and points at
  `WorkbookEditor.Unprotect` for that case.

* **Core: the bullet glyph in a rendered PDF is a Unicode bullet, not the Symbol-font one.** This is
  the trade the bullet fix makes, and it is on by default: `U+F0B7` becomes `U+2022`, `U+F0A7`
  becomes `U+00B7`. Visually near-identical, and the alternative is not a faithful conversion but no
  conversion at all. Only list markers in `word/numbering.xml` are touched — document text is never
  rewritten, and a document with no list is returned byte-for-byte unchanged.


### Fixed

* **Core: converting HTML now names the failure that real pages hit most often.** Measured across
  179 real `.gov` pages, **14 of them — 7.7% —** failed with a bare `IndexOutOfRangeException` and
  the message *"See the inner exception for details"*, which named no table, no cell and no remedy.
  The cause is a table cell whose `rowspan` reaches past the last row of its table. The message now
  names the construct, says the markup is **valid** — browsers clamp such a rowspan, so the page
  renders correctly in a browser — and gives the remedy; the inner exception is unchanged. Those
  pages still do not convert. They now say why
  ([#292](https://github.com/Ank-KhoaHo/DocToolkit/issues/292)).

* **Core: two failure messages now name a cause they can actually distinguish.** Both previously
  asserted one specific reason as fact when the code could not tell which of several had occurred,
  which sent at least one investigation after a network regression that did not exist. The first
  sentence of each is unchanged ([#287](https://github.com/Ank-KhoaHo/DocToolkit/issues/287)).

## [0.29.0](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.28.0...v0.29.0) (2026-08-16)


**Extensions package only. The core package is unchanged, and nothing existing behaves
differently.**

### Added

* **Extensions:** everything 0.28.0 added to the core package is now injectable ([#276](https://github.com/Ank-KhoaHo/DocToolkit/issues/276)) ([48e5d4a](https://github.com/Ank-KhoaHo/DocToolkit/commit/48e5d4a56ea502869cc59b985973b05404714fda)).

    ```csharp
    public sealed class InvoiceService(IPdfEditor pdf, IDocToDocxConverter legacy)
    {
        public byte[] Lock(byte[] file) =>
            pdf.Protect(file, new PdfProtection { UserPassword = "s3cret" });
    }
    ```

    New interface `IDocToDocxConverter` (legacy Word 97-2003 `.doc`), and the password members on
    the ones you already inject: `Protect`/`Unprotect` on `IPdfEditor`, and
    `Protect`/`Unprotect`/`IsProtected` on `IDocxEditor`, `IWorkbookEditor` and
    `IPresentationEditor` — `byte[]` for `byte[]` and `Stream` for `Stream`, as everywhere else here.
    All are registered by `services.AddDocToolkit()`; no new call is needed.

    **One thing to know if you pin versions.** `Ank.DocToolkit.Extensions.DependencyInjection` now
    requires `Ank.DocToolkit` **>= 0.28.0**, up from 0.26.0 — it cannot wrap a method that has not
    shipped, so the mirror always lands one release after the core feature. Upgrading the extensions
    package will bring the core package with it.

## [0.28.0](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.27.5...v0.28.0) (2026-08-16)


**Three new capabilities, all additive. Nothing existing changed and no upgrade action is needed.**

> **This entry was expanded after publication, on 2026-08-16.** Nothing published was removed or
> reworded — the first bullet below was **missing entirely**, and the other two carried only their
> commit titles. PDF password protection reached this release inside a squash-merged stacked pull
> request whose title named only the Office formats, so release-please could not see it. Recording
> the repair here rather than silently rewriting: everything already published is still present
> below, verbatim.

### Added

* **core:** **password-protect and unprotect a PDF** ([#271](https://github.com/Ank-KhoaHo/DocToolkit/pull/271), shipped in [e34291e](https://github.com/Ank-KhoaHo/DocToolkit/commit/e34291e77f5986daffce099601042b865e8904b8))

    ```csharp
    byte[] locked = PdfEditor.Protect(pdf, new PdfProtection { UserPassword = "s3cret" });
    byte[] opened = PdfEditor.Unprotect(locked, "s3cret");
    ```

    `PdfProtection` carries both passwords, seven permission flags and the cipher. **The two
    passwords are not interchangeable**: a *user* password is required to open the file and is
    enforced by cryptography; an *owner* password leaves the document readable and only asks readers
    to honour the permissions. If content must not be read, set `UserPassword`.

    Two behaviours worth knowing before relying on them. **`Unprotect` needs the owner password when
    the document has one**, even if you also know the user password — removing protection is a
    modification, which the PDF format reserves for the owner. And **every permission defaults to
    allowed**, so adding a password does not silently stop a document being printed.

    `PdfEncryptionStrength.Aes128` is the default because every reader in service opens it; `Aes256`
    is available and needs a PDF 2.0 reader (Acrobat X and later).

* **core:** password-protect DOCX, XLSX and PPTX ([#272](https://github.com/Ank-KhoaHo/DocToolkit/issues/272)) ([e34291e](https://github.com/Ank-KhoaHo/DocToolkit/commit/e34291e77f5986daffce099601042b865e8904b8))

    ```csharp
    byte[] locked = WorkbookEditor.Protect(xlsx, "s3cret");
    bool needsPassword = WorkbookEditor.IsProtected(bytes);   // no password required to ask
    ```

    The same three members on `DocxEditor`, `WorkbookEditor` and `PresentationEditor`, with `Stream`
    overloads both ways. **This is file encryption, not the "restrict editing" flag** — Office puts
    both under one menu, and only this one stops anyone reading the file.

    **An encrypted Office file is not a package any more**: a plain `.docx`/`.xlsx`/`.pptx` is a ZIP,
    the encrypted form is a compound file with the package sealed inside. Every other method on those
    classes therefore refuses one — call `Unprotect` first. A wrong password and a file that was never
    encrypted are reported as different failures.

* **core:** read and convert legacy Word 97-2003 binary .doc ([#269](https://github.com/Ank-KhoaHo/DocToolkit/issues/269)) ([2ef72c4](https://github.com/Ank-KhoaHo/DocToolkit/commit/2ef72c45022458a0fd3def40d4c0e67901206ed7))

    ```csharp
    string text = DocToDocxConverter.ExtractText(doc);   // never refuses
    byte[] docx = DocToDocxConverter.Convert(doc);        // refuses if content would be lost
    ```

    **Converting refuses by default, and that is the common case rather than a rare one.** A legacy
    `.doc` keeps pictures, drawings and form fields in a binary stream a `.docx` cannot carry, so
    `Convert` throws `DocumentConversionException` rather than quietly returning a document with those
    payloads missing. Measured: any `.doc` containing a **table** has such a stream — plain text, bold
    runs and headings do not.

    Accept the loss deliberately with `new LegacyDocOptions { AllowContentLoss = true }`, or call
    `ConvertWithReport` for the same bytes plus a list of exactly what was dropped. Text, tables
    (every cell) and character formatting survive either way. `ExtractText` takes no options and never
    refuses. **Reading only** — native `.doc` saving is unsupported upstream, so none is offered.


### Fixed

* **build:** describe what actually ships on both nuget.org pages ([#273](https://github.com/Ank-KhoaHo/DocToolkit/issues/273)) ([29476c8](https://github.com/Ank-KhoaHo/DocToolkit/commit/29476c8db91a9a98ebb164626d1ae06cfb4434e2))
* **test:** read /P from the encryption dictionary, not from ciphertext ([#274](https://github.com/Ank-KhoaHo/DocToolkit/issues/274)) ([5b687e5](https://github.com/Ank-KhoaHo/DocToolkit/commit/5b687e516b3a69988f312e824acb7e86d8f1a54c))

## [0.27.5](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.27.4...v0.27.5) (2026-08-15)


**Exception message text changed. No API changed and no call behaves differently** — but if you
match on `DocumentConversionException.Message`, read the migration note below before upgrading.

### Changed

* **`DocumentConversionException` messages now say what to do, not only what failed.** They
  previously named the failure and stopped — `"Failed to read the PDF."`, `"A template row had no
  parent table."` — which is little help in a log. Fifteen messages that had a *specific* remedy now
  carry one sentence of guidance; around twenty generic wrappers now point at the inner exception
  rather than inventing a cause they cannot know.

  **Migrating:** the first sentence of every message is unchanged, so a substring match on the old
  text still matches. **An exact-equality match on `.Message` will not.** If you branch on message
  text, match the prefix or — better — catch the type and read `InnerException`
  ([#266](https://github.com/Ank-KhoaHo/DocToolkit/issues/266)).

  One message deliberately names *two* causes rather than the likely one: a JPEG with no
  Start-Of-Frame segment may be truncated **or** simply not a well-formed JPEG, and the code cannot
  tell those apart. Naming a cause a failure cannot distinguish is how a wrong message costs
  somebody an afternoon.

## [0.27.4](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.27.3...v0.27.4) (2026-08-15)


**Package metadata only. No code changed, nothing behaves differently, and the assemblies are
identical to 0.27.3** — if you are not reading this package's nuget.org page, there is nothing here
you need.

### Changed

* **The nuget.org description now names all eleven converters.** It listed eight, omitting
  DOCX → PDF, XLSX → PDF and PPTX → PDF — all three of which have shipped for several releases.
  Nothing was wrong with what it said; three capabilities were simply absent from it, so the
  package understated itself to anyone evaluating it from the listing
  ([#262](https://github.com/Ank-KhoaHo/DocToolkit/issues/262)).

## [0.27.3](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.27.2...v0.27.3) (2026-08-15)


**A consistency release for argument handling. Nothing new, and nothing renders differently** —
but **three calls now throw a different exception type for the same bad input**, so read *Changed*
before upgrading if you catch exceptions around `PdfEditor` or `ConvertWithReport`.

### Changed

* **`PdfEditor`: an empty `byte[]` is now an `ArgumentException`, not a
  `DocumentConversionException`.** It affects `PageCount`, `Merge`, `ExtractPages`, `RemovePages`,
  `RotatePages`, `ReorderPages`, `InsertPages`, `ReadMetadata` and `WithMetadata`. `ExtractText`
  already behaved this way, so one class was answering the same mistake two different ways
  depending on which method you called.

  **Migrating:** if you catch `DocumentConversionException` to handle a truncated or empty upload,
  add `ArgumentException` — or check `.Length` before calling, which is what the new behaviour is
  telling you to do ([#258](https://github.com/Ank-KhoaHo/DocToolkit/issues/258)).

* **`PdfEditor`'s file-path overloads reject a whitespace-only path.**
  `PageCountAsync("   ")` and `ExtractTextAsync("   ")` now throw `ArgumentException` naming
  `path`, matching every other file-path overload in the library, instead of surfacing a
  platform-dependent framework exception from the file system
  ([#258](https://github.com/Ank-KhoaHo/DocToolkit/issues/258)).

* **`DocxToHtmlConverter.ConvertWithReport` and `DocxToMarkdownConverter.ConvertWithReport` now
  wrap a conversion-internal `ArgumentException`** in `DocumentConversionException`, matching their
  sibling `Convert`. Previously the same malformed package produced a different exception type
  depending on which of the two you called, so a caller catching `DocumentConversionException`
  around both crashed on one ([#258](https://github.com/Ank-KhoaHo/DocToolkit/issues/258)).

### Fixed

* **Markdown → PDF through the `Stream` overload** handed the renderer a read-only buffer where the
  renderer documents needing an expandable one. No reported failure — the contract and the call
  site disagreed, and only one of them could be right
  ([#258](https://github.com/Ank-KhoaHo/DocToolkit/issues/258)).

* **HTML → DOCX now validates `RemoteImageOptions` at the single point every overload passes
  through**, rather than at each entry point that remembered to. An unvalidated timeout could
  abort the conversion instead of skipping the image it applied to
  ([#258](https://github.com/Ank-KhoaHo/DocToolkit/issues/258)).

## [0.27.2](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.27.1...v0.27.2) (2026-08-15)


**A concurrency fix, and the only people it affects are the ones it affects badly.** If you call
this package's async methods with `await`, nothing changes. If you **block** on one of them —
`.Result`, `.GetAwaiter().GetResult()` — from a thread with a `SynchronizationContext`, which means
WPF, WinForms or classic ASP.NET, some of them could **deadlock**. They no longer can.

### Fixed

* **Core:** every `await` in the library now carries `ConfigureAwait(false)`. Nine did not, and a
  blocking caller on a UI or classic-ASP.NET synchronisation context could deadlock on any of them:
  `HtmlToPdfConverter.ConvertAsync` (both overloads) and `ConvertToFileAsync`,
  `HtmlToDocxConverter.ConvertToFileAsync` and the shared HTML → DOCX build path,
  `PdfEditor.PageCountAsync(string)`, and the remote-image read used by the opt-in image fetch.
  **No behaviour change for a caller who awaits**, which is the normal case
  ([e91a653](https://github.com/Ank-KhoaHo/DocToolkit/commit/e91a65310b9279b498783ba78f25d4d6347242ef)).

## [0.27.1](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.27.0...v0.27.1) (2026-08-15)

**A packaging and metadata release. Nothing in the public API changed**, and no call behaves
differently — if you are on 0.27.0 and not reading the nuget.org page or stepping into the source,
there is nothing here you need.

### Changed

* **The package description now describes what actually ships.** It had gone stale over four
  releases, naming only HTML → DOCX/PDF while Markdown → DOCX/PDF, DOCX → HTML/Markdown,
  XLSX → CSV/HTML and PDF text extraction had all shipped. If you evaluated this package on its
  nuget.org summary before today, it understated it
  ([e0fdef5](https://github.com/Ank-KhoaHo/DocToolkit/commit/e0fdef5f863892ec10ef0d4be47b5bc8677a6867)).
* **Both packages now carry a repository URL and a copyright notice.** Source Link previously
  emitted a commit hash with nothing saying which repository it belonged to, so stepping into the
  library's source while debugging did not resolve. It does now
  ([c54ad91](https://github.com/Ank-KhoaHo/DocToolkit/commit/c54ad91d2f2744710235c0b3e49d05f01ded13d8)).

### Fixed

* **Core:** `WorkbookEditor.SetCell` and `WorkbookEditor.Format` now share one implementation with
  their `Stream` overloads instead of carrying a second copy. **No behaviour change** — the two
  paths were verified identical before the change — but they had already begun to diverge in how
  they opened the workbook, which is how the two forms of a method come to disagree
  ([7b54430](https://github.com/Ank-KhoaHo/DocToolkit/commit/7b54430825e0e33107d251f479495ef3f214c766)).

## [0.27.0](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.26.0...v0.27.0) (2026-08-14)


### Added

* **core:** claim native AOT compatibility, now that it is verified ([45ff45b](https://github.com/Ank-KhoaHo/DocToolkit/commit/45ff45b7cf72590f6b329ffb2ac28ac915d6638c))
* **core:** claim native AOT compatibility, now that it is verified ([6a0434c](https://github.com/Ank-KhoaHo/DocToolkit/commit/6a0434cdd55ffb0b2a2a9a352a46520b0298febf))
* **extensions:** restore 1:1 parity and enforce it with a derived check ([2996756](https://github.com/Ank-KhoaHo/DocToolkit/commit/2996756c7d42a870720b326edd9da54016480089))
* **extensions:** restore 1:1 parity and enforce it with a derived check ([c7f7cf1](https://github.com/Ank-KhoaHo/DocToolkit/commit/c7f7cf14e4cbdf2b419777723bb2637d133ebe63))


### Fixed

* **core:** make XlsxFormat's number-format map genuinely immutable ([a699c81](https://github.com/Ank-KhoaHo/DocToolkit/commit/a699c813a922cfe3dcb73fc23308f6b8b53640f1))
* **core:** make XlsxFormat's number-format map genuinely immutable ([45be8a0](https://github.com/Ank-KhoaHo/DocToolkit/commit/45be8a03cf41ac5ca09b798febd59f474d24dd40))

## [0.26.0](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.25.0...v0.26.0) (2026-08-13)


### Added

* **core:** convert Markdown to DOCX ([8266b3f](https://github.com/Ank-KhoaHo/DocToolkit/commit/8266b3fcac60dbf970a71dd81a9947c0c71d9015)).
  `MarkdownToDocxConverter`, completing the round trip `DocxToMarkdownConverter` opened.
  **Nothing in it reaches the network or the disk**: an image URL becomes a hyperlink rather than
  a fetch, a local file reference is refused, and `data:` images are inlined. Also available as
  `ConvertWithReport`, which reports what the conversion could not carry across.
  **No new dependency** — the capability was already in the resolved graph.
* **core:** convert Markdown to PDF ([dc310c2](https://github.com/Ank-KhoaHo/DocToolkit/commit/dc310c24893966a0324e89646f3da0bbbf4894f5)).
  `MarkdownToPdfConverter`, which pivots through DOCX exactly as `HtmlToPdfConverter` does and
  inherits the offline guarantee above unchanged. `ConvertWithReport` here carries the
  Markdown → DOCX half's warnings only; the render half produces no report.
* **core:** export a sheet as CSV or as an HTML table ([4b70084](https://github.com/Ank-KhoaHo/DocToolkit/commit/4b70084138f20a8bff665493ad0d945a0fa55a9a)).
  `XlsxToCsvConverter` and `XlsxToHtmlConverter`, one named sheet at a time. A formula exports its
  computed **value**. The HTML is a `<table>` **fragment**, not a whole document — deliberately
  unlike `DocxToHtmlConverter` — and every cell is escaped.
  **Read the note under _Changed_ below about culture before using these.**
* **core:** format a sheet — bold header, freeze, auto-fit, number formats ([694a1f4](https://github.com/Ank-KhoaHo/DocToolkit/commit/694a1f4319c7a93fe504f27c27c2e3ba16cf18f4)).
  `WorkbookEditor.Format` and `XlsxFormat`, applied to an existing workbook so it composes with
  `Create`, `AppendRows` and workbooks this library never made. `XlsxFormat.Report` is the
  three-setting preset for a readable report. **The set is deliberately small and closed**: if you
  need fonts, borders, fills or conditional rules, use ClosedXML directly rather than through a
  thinner API in front of it.


### Changed

* **core:** the two new spreadsheet exporters are **culture-invariant**, and this differs on purpose
  from `WorkbookEditor.ReadSheet`.

  `XlsxToCsvConverter` and `XlsxToHtmlConverter` render numbers with a dot and dates as ISO 8601
  regardless of the machine's regional settings. `ReadSheet` is unchanged and still follows
  `CurrentCulture`, as its documentation has always said.

  That asymmetry is deliberate rather than an oversight. `ReadSheet` returns data you inspect in
  process; an exporter produces a file you hand to something else, and on a machine whose culture
  uses a decimal comma, `1234.5` would export as `1234,5` — **a decimal comma inside a
  comma-delimited file**, which every downstream reader would see as an extra column. Measured
  across en-US, de-DE and fr-FR.

  If you were relying on export output matching `ReadSheet`'s text on a non-invariant culture,
  that is the one behaviour here that will look different — and it is the direction that was
  previously wrong.

## [0.25.0](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.24.0...v0.25.0) (2026-08-13)


### Added

* **core:** report what a DOCX text conversion could not carry across ([6148021](https://github.com/Ank-KhoaHo/DocToolkit/commit/6148021220142908e144f92ed4430c72184f8a61)).
  `DocxToHtmlConverter.ConvertWithReport` and `DocxToMarkdownConverter.ConvertWithReport` return
  the same output as `Convert` plus a list of `ConversionWarning`, each carrying a `Code`, a
  `Message` and a `ConversionLossKind`. The existing `Convert` overloads are unchanged.
  Worth knowing: a DOCX → HTML conversion has **always** reported a loss
  (`SectionLayoutFlattened`, an approximation of section page geometry) — it was computed on every
  call and discarded, and this is the first release in which you can see it.
* **extensions:** mirror `PdfEditor.ExtractText` on `IPdfEditor` ([907df58](https://github.com/Ank-KhoaHo/DocToolkit/commit/907df58219ca9fbe35531382b89b5ff6a146a264)).
  Adds `ExtractText(byte[])` and `ExtractTextAsync(Stream)`. **Source-breaking for anyone who
  implements `IPdfEditor` by hand**; additive for anyone who injects it.


### Changed

* **core:** `PdfEditor`'s seven other async methods now guard their streams the way every other
  `Stream` overload in the library does. **Four behaviour changes** to `PageCountAsync`,
  `MergeAsync`, `ExtractPagesAsync`, `RemovePagesAsync`, `RotatePagesAsync`, `ReorderPagesAsync`
  and `InsertPagesAsync`:
  * an unreadable source or unwritable destination now throws `ArgumentException` naming the
    parameter, where it previously surfaced `NotSupportedException` from inside `CopyToAsync`;
  * an empty source now throws `ArgumentException`, where it previously passed empty bytes on and
    reported them later as `DocumentConversionException`;
  * a source is read from its current `Position` and left drained, where a `MemoryStream` had its
    whole buffer read regardless of position and was not advanced;
  * a failure writing to your destination stream is now wrapped in `DocumentConversionException`.

  `ExtractText`, added in 0.24.0, is unaffected — it already behaved this way.


### Fixed

* **core:** build PdfEditor's Stream overloads on StreamPipeline ([cb413bd](https://github.com/Ank-KhoaHo/DocToolkit/commit/cb413bda8b57925cf53526f495424513fa61cc4e)).
  `PdfEditor` was the only Stream-overload class not built on the shared pipeline, and none of its
  async methods was covered by the suite that holds that surface to one shape. See **Changed**
  above for what this alters for callers.

## [0.24.0](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.23.0...v0.24.0) (2026-08-12)


### Added

* **core:** read text out of a PDF ([#209](https://github.com/Ank-KhoaHo/DocToolkit/issues/209)) ([0d3dd90](https://github.com/Ank-KhoaHo/DocToolkit/commit/0d3dd908e9bd1ac1aff8a3d76adcf862c7e51dd9))

## [0.23.0](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.22.0...v0.23.0) (2026-08-12)


### Added

* **di:** mirror DocxEditor table read-back onto IDocxEditor ([#205](https://github.com/Ank-KhoaHo/DocToolkit/issues/205)) ([71039e5](https://github.com/Ank-KhoaHo/DocToolkit/commit/71039e54e45a4e19bc2d8812315cedd25cf854d3))

## [0.22.0](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.21.0...v0.22.0) (2026-08-12)


### Added

* **core:** read a DOCX table back as data ([#204](https://github.com/Ank-KhoaHo/DocToolkit/issues/204)) ([739e443](https://github.com/Ank-KhoaHo/DocToolkit/commit/739e44318008466bb4c349b90a0113598f35a036))
* **extensions:** mirror PdfEditor page ops and PresentationEditor.ReplaceImage ([#202](https://github.com/Ank-KhoaHo/DocToolkit/issues/202)) ([6ddbd8e](https://github.com/Ank-KhoaHo/DocToolkit/commit/6ddbd8ef2a71204456d27744553dda5def3272e1))

## [0.21.0](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.20.0...v0.21.0) (2026-08-11)


### ⚠ BREAKING CHANGES

* **core:** DocxEditor.ExtractText and ExtractTextAsync now separate blocks with \n and table cells with \t. Previously all text was concatenated with no separator. text.Replace("\n", "").Replace("\t", "") reproduces the old output.

### Added

* **core:** add PdfEditor.RemovePages, the complement of ExtractPages ([#198](https://github.com/Ank-KhoaHo/DocToolkit/issues/198)) ([5ede858](https://github.com/Ank-KhoaHo/DocToolkit/commit/5ede858200f9e76f04f30e6ab40431436ee9a13b))
* **core:** add PdfEditor.ReorderPages and InsertPages, closing A25 ([#200](https://github.com/Ank-KhoaHo/DocToolkit/issues/200)) ([92e7768](https://github.com/Ank-KhoaHo/DocToolkit/commit/92e77682785774a69be84ef4f4ddb5ce8b138d3f))
* **core:** add PdfEditor.RotatePages, turning pages a quarter at a time ([#199](https://github.com/Ank-KhoaHo/DocToolkit/issues/199)) ([5877fff](https://github.com/Ank-KhoaHo/DocToolkit/commit/5877fffc370179a590170611c30d12e993be5613))
* **core:** add PresentationEditor.ReplaceImage ([#201](https://github.com/Ank-KhoaHo/DocToolkit/issues/201)) ([4e17d12](https://github.com/Ank-KhoaHo/DocToolkit/commit/4e17d1257c58d986ef3f6105eae3e35b4c9dcdfc))


### Fixed

* **core:** separate block boundaries in DocxEditor.ExtractText ([#195](https://github.com/Ank-KhoaHo/DocToolkit/issues/195)) ([3f67241](https://github.com/Ank-KhoaHo/DocToolkit/commit/3f67241416c30a8d15eefc54b523729d5601e1fc))

## [0.20.0](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.19.1...v0.20.0) (2026-08-10)


### Added

* **extensions:** raise the core floor so DI consumers can use headers ([#192](https://github.com/Ank-KhoaHo/DocToolkit/issues/192)) ([5c973ef](https://github.com/Ank-KhoaHo/DocToolkit/commit/5c973effd68a9b6d3533e939be803a463433b913))

## [0.19.1](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.19.0...v0.19.1) (2026-08-10)


### Fixed

* **ci:** require a Conventional Commit pull request title ([#190](https://github.com/Ank-KhoaHo/DocToolkit/issues/190)) ([367d57b](https://github.com/Ank-KhoaHo/DocToolkit/commit/367d57b2cbac4270a87866bea99aa2e706900e37))

## [0.19.0](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.18.1...v0.19.0) (2026-08-10)


### Added

* **extensions:** a default page setup on DocToolkitOptions ([#186](https://github.com/Ank-KhoaHo/DocToolkit/issues/186)) ([4e76adc](https://github.com/Ank-KhoaHo/DocToolkit/commit/4e76adc9980f3f6e31dda66c2c951d2619e21825))
* **core:** headers and footers on generated documents ([#189](https://github.com/Ank-KhoaHo/DocToolkit/pull/189)) ([9dcf579](https://github.com/Ank-KhoaHo/DocToolkit/commit/9dcf579))

  Added by hand. release-please could not parse #189's squash-commit subject - the pull
  request was titled without a Conventional Commit prefix, per advice that was correct for
  merge commits and wrong for squash - so it discarded the commit and every `feat:` line in
  it. The feature is in 0.19.0; only this entry was missing.

## [0.18.1](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.18.0...v0.18.1) (2026-08-09)


### Fixed

* **ci:** restore local tools in release.yml, and guard that it stays ([#182](https://github.com/Ank-KhoaHo/DocToolkit/issues/182)) ([9830b9f](https://github.com/Ank-KhoaHo/DocToolkit/commit/9830b9f745c83b7d76441189e4931f034706f3df))
* **core:** reject null html on the overloads that take a page setup ([#185](https://github.com/Ank-KhoaHo/DocToolkit/issues/185)) ([a9ef5ec](https://github.com/Ank-KhoaHo/DocToolkit/commit/a9ef5ecaf12e0e816b7ece390c5b774254af57ee))

## [0.18.0](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.17.0...v0.18.0) (2026-08-09)


### Added

* **core:** allow a page setup and remote images in the same call ([#180](https://github.com/Ank-KhoaHo/DocToolkit/issues/180)) ([de2c3b2](https://github.com/Ank-KhoaHo/DocToolkit/commit/de2c3b2352eba52ddb3be219d5082a0272fff7c0))

## [0.17.0](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.16.0...v0.17.0) (2026-08-09)

> **Never published to nuget.org.** The release workflow failed before pushing: its new
> SBOM step used a local tool that the job had not restored. The packages for this version
> do not exist and the version number is skipped - 0.16.0 is followed by 0.18.0 on the feed.
> Nothing is lost: the change here was to the release pipeline, not to either library, so
> 0.17.0 and 0.16.0 would have been identical to a consumer. The SBOMs it describes ship
> from 0.18.0 onwards.


### Added

* **ci:** publish an attested CycloneDX SBOM with every release ([#177](https://github.com/Ank-KhoaHo/DocToolkit/issues/177)) ([8919496](https://github.com/Ank-KhoaHo/DocToolkit/commit/8919496c80fbdb641a7b48d447e46e763cda3249))

## [0.16.0](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.15.0...v0.16.0) (2026-08-09)


### Added

* **extensions:** mirror PdfEditor on IPdfEditor ([91e958b](https://github.com/Ank-KhoaHo/DocToolkit/commit/91e958b3078e917af50dad3af65973b280c45e10))

## [0.15.0](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.14.0...v0.15.0) (2026-08-09)


### Added

* **core:** read an existing PDF - page count, merge, extract, metadata ([971fffa](https://github.com/Ank-KhoaHo/DocToolkit/commit/971fffa0ce6b63237925ff75fd15ab39efbbd7ac))

## [0.14.0](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.13.0...v0.14.0) (2026-08-09)


### Added

* **extensions:** mirror the 0.13.0 API on the injectable interfaces ([67bda53](https://github.com/Ank-KhoaHo/DocToolkit/commit/67bda539cac2840c6f48e8aec48a5c9c42a2fd46))


### Fixed

* **extensions:** apply option changes without a restart ([f0f0508](https://github.com/Ank-KhoaHo/DocToolkit/commit/f0f0508c37bc02f2a51e4eb8aa7d7c3c4dcb320e))

## [0.13.0](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.12.0...v0.13.0) (2026-08-09)


### Added

* collect nuget.org download counts into a daily CSV ([ced271c](https://github.com/Ank-KhoaHo/DocToolkit/commit/ced271c9477229de571e4156c6855a36eb6059b6))
* **core:** add PageSetup, the page size and margins a document is laid out on ([38df59d](https://github.com/Ank-KhoaHo/DocToolkit/commit/38df59df8ef48be853421301c37bead8525f665d))
* **core:** convert DOCX to HTML and Markdown ([4a851ff](https://github.com/Ank-KhoaHo/DocToolkit/commit/4a851ff3485b6321515c7f33c1195f59e5e1e18a))
* **core:** default every producer to A4 and document the change ([87b6c62](https://github.com/Ank-KhoaHo/DocToolkit/commit/87b6c6211aaaf6258c7b94321fb097751e9ca712))
* **core:** give DocxEditor.Create a page setup, defaulting to A4 ([8c13aeb](https://github.com/Ank-KhoaHo/DocToolkit/commit/8c13aeb612bbd80e54d58dcd42d158f78d9cf86c))
* **core:** give HtmlToDocxConverter a page setup, defaulting to A4 ([f447e33](https://github.com/Ank-KhoaHo/DocToolkit/commit/f447e33fb49c5088b40ae0030660abb0bedc91d4))
* **core:** give HtmlToPdfConverter a page setup, defaulting to A4 ([c588f85](https://github.com/Ank-KhoaHo/DocToolkit/commit/c588f854a7d6430c9c84e6539436a0723d9fda30))
* **core:** render XLSX and PPTX to PDF ([3521dc9](https://github.com/Ank-KhoaHo/DocToolkit/commit/3521dc9b4386d183134633b8a88b7633f53c2741))
* **core:** report remote-image fetch outcomes as traces and metrics ([8427b7f](https://github.com/Ank-KhoaHo/DocToolkit/commit/8427b7f725d2f4027319fe233f1b1bdfb8d57e78))
* **extensions:** mirror the XLSX writing methods on IWorkbookEditor ([f8529ba](https://github.com/Ank-KhoaHo/DocToolkit/commit/f8529baad23de02f892dd5e6aa2aac1838a022ee))
* render the download history as markdown and html reports ([6b885b3](https://github.com/Ank-KhoaHo/DocToolkit/commit/6b885b33b245ce8e97b263cdec08f2f03d74b026))
* schedule the nuget usage tracker daily ([a7ac2d2](https://github.com/Ank-KhoaHo/DocToolkit/commit/a7ac2d285003ecd00f8f804072650f8a920deb77))


### Fixed

* degrade per-region and never record an unverifiable first reading ([8ac9523](https://github.com/Ank-KhoaHo/DocToolkit/commit/8ac9523d1cc4eff704b5777887e1503a36a39157))
* grant actions read scope and keep partial data on a bad day ([26b1f02](https://github.com/Ank-KhoaHo/DocToolkit/commit/26b1f02963bc2292b8677199d632c9f129cc59c3))
* label the CI runs window and survive an unparseable date ([d754831](https://github.com/Ank-KhoaHo/DocToolkit/commit/d7548311ba30b03569288b64c757d109506bebd7))
* never let an unreadable runs.json block the downloads write ([4de5fe6](https://github.com/Ank-KhoaHo/DocToolkit/commit/4de5fe66ca71ec9f9686e48930492a41331094a1))
* remove the adoption signal that could never fire ([b58edff](https://github.com/Ank-KhoaHo/DocToolkit/commit/b58edff625896154cf104a614aede3fc443c9817))
* weigh quiet downloads by calendar days, not by sample count ([70a6fca](https://github.com/Ank-KhoaHo/DocToolkit/commit/70a6fcaa36f692398a4a69bc0292826658d25ffb))
* write the CSVs atomically so a failed write cannot truncate history ([aaa4c28](https://github.com/Ank-KhoaHo/DocToolkit/commit/aaa4c28808c800dc67404785b0b0146f7ee905db))


### Changed

* keep the csproj version in step with the release ([6795ce9](https://github.com/Ank-KhoaHo/DocToolkit/commit/6795ce943404e5ba084a8fdffa5fe4f636b5160a))

## [0.12.0](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.11.0...v0.12.0) (2026-08-08)


### Added

* **Core: the XLSX writing surface, which was previously read-mostly.** A workbook could only ever
  be created with **one sheet**, rows could not be added to an existing one, and a cell could not
  hold a formula. All three are now possible.

  ```csharp
  byte[] workbook = WorkbookEditor.Create(new[]
  {
      XlsxSheet.Named("Sales",   new[] { new object?[] { "Region", "Total" },
                                         new object?[] { "EMEA", 1200 } }),
      XlsxSheet.Named("Summary", new[] { new object?[] { "Grand total",
                                         XlsxFormula.From("SUM(Sales!B2:B2)") } }),
  });

  workbook = WorkbookEditor.AppendRows(workbook, "Sales", moreRows);
  ```

  `XlsxSheet` completes the typed-creation trio alongside `DocxBlock` and `PptxSlide`.
  `Create(IEnumerable<XlsxSheet>)` and `AppendRows` each have the usual `Stream` and file-path
  forms. The single-sheet `Create` is unchanged.

  **`XlsxFormula` is a cell value, not a method** — it works anywhere a cell value is accepted,
  including as the `value` argument to the existing `SetCell`.

  **Formulas carry no cached value.** The cell holds the formula and nothing else. Excel
  recalculates when it opens the file, and `ReadCell`/`ReadSheet` compute the value on read — but a
  third-party reader that only reads cached values, such as openpyxl with `data_only=True`, sees an
  empty cell until Excel has opened and saved the file. A formula that cannot be evaluated reads
  back as its Excel error string (`#DIV/0!`, `#NAME?`, `#REF!`) rather than throwing.
  ([49982ad](https://github.com/Ank-KhoaHo/DocToolkit/commit/49982ad14f29be6ce845845e29d1d20ea496bf19),
  [8af39c8](https://github.com/Ank-KhoaHo/DocToolkit/commit/8af39c8e3a68c9a945e1c9f05d9dc8645ee92543),
  [f426aaa](https://github.com/Ank-KhoaHo/DocToolkit/commit/f426aaa4b6b09bbb85fb296edb64b45be039638b),
  [5d8b2fe](https://github.com/Ank-KhoaHo/DocToolkit/commit/5d8b2fe0a8db19fd531344a0851668866559e9b4))

* **Both packages are marked trimmable.** Assemblies carry the `IsTrimmable` metadata, so an app
  published with `PublishTrimmed` keeps only the parts of DocToolkit it uses. CI checks the claim
  rather than asserting it: it trim-publishes an application over the real dependency graph and
  **runs it**, verifying every capability still works.

  Two limits worth knowing. **ClosedXML emits a trim warning** (`IL2090`, in
  `DescribedEnumParser<T>`); spreadsheet reading and writing work correctly in the trimmed app CI
  runs, but the warning will appear in your publish output, and it is a dependency's rather than
  ours. And **Native AOT is not claimed** — `IsAotCompatible` is a strictly stronger promise that
  has not been verified end to end here, and an unverified compatibility claim is worse than an
  absent one.
  ([83f182f](https://github.com/Ank-KhoaHo/DocToolkit/commit/83f182f593a0ed382600fe0d45b9f8177e9bfc5a),
  [46adf28](https://github.com/Ank-KhoaHo/DocToolkit/commit/46adf28c7962a412dddaabfdd82a89fde46137c2))

* **Extensions: `IPresentationEditor` gains `Create` and `CreateAsync`.** The deck-building methods
  added to `PresentationEditor` in 0.11.0 are now reachable through dependency injection, restoring
  the 1:1 mirror between the interface and the static class. The extensions package builds against
  the *published* core package, so a new core method can only be mirrored one release later; this
  closes that gap for 0.11.0's addition.
  ([7578053](https://github.com/Ank-KhoaHo/DocToolkit/commit/7578053f9509c85930c4f3bc98860d013a310936))


### Fixed

* **Core: slides built by `PresentationEditor.Create` are now real title and body placeholders.**
  They were previously loose text boxes. Decks rendered correctly and passed schema validation, but
  PowerPoint had no way to tell which shape was the title — so **Outline View listed every slide
  with no title**, "Reset Slide" had no placeholder to restore geometry to, and the layout gallery
  showed the internal part name "SlideLayout2". Measured in PowerPoint 16.0, `Shapes.HasTitle` went
  from `False` to `True` and the layout is now labelled "Title and Content". Decks produced by
  earlier versions are unaffected on disk; rebuild them with 0.12.0 to pick this up.
  ([a8dcc36](https://github.com/Ank-KhoaHo/DocToolkit/commit/a8dcc3646fc677e585849b546af4d4abc89b1d52))

* **Core: an invalid sheet name now fails fast, and the exception type changed.**
  `WorkbookEditor.Create` accepted a sheet name Excel cannot use — longer than 31 characters, or
  containing any of `: \ / ? * [ ]` — and let it reach ClosedXML, surfacing as a
  `DocumentConversionException` wrapping a third-party message. It now throws `ArgumentException`
  naming the `sheetName` parameter and stating the rule.

  **If you catch `DocumentConversionException` around workbook creation to handle a bad sheet
  name, that handler no longer fires.** Excel's rules are now applied by every path that names a
  sheet, so the two cannot disagree.
  ([aa806aa](https://github.com/Ank-KhoaHo/DocToolkit/commit/aa806aa29109c32225ea5219cdf9351fa24bf8f0))


### Changed

* **Published builds are now reproducible, and stepping into the library works.** `Deterministic`
  alone does not normalize source paths, so every previously published assembly and PDB carried the
  absolute paths of the machine that built it — meaning builds from two machines differed, and
  SourceLink could not reliably map a frame back to a file. CI builds now set
  `ContinuousIntegrationBuild`, which normalizes them.
  ([52e4c9a](https://github.com/Ank-KhoaHo/DocToolkit/commit/52e4c9a67ecdfcb03fcce77721d660c8f1170a4d))

* **Core:** `OfficeIMO.Word.Pdf` updated to 3.2.0. Measured before shipping, because 3.1.0 changed
  PDF font embedding: the same document through `HtmlToDocx` then `DocxToPdf` produces a
  **byte-identical** PDF on 3.1.0 and 3.2.0, and text still extracts correctly. No package was
  added or removed; the dependency closure stays at 16.
  ([3061462](https://github.com/Ank-KhoaHo/DocToolkit/commit/3061462d4c137921ea5565ae9e41266fbed1838e))

## [0.11.0](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.10.0...v0.11.0) (2026-08-07)


### Added

* **Core: create a PPTX from scratch, from a typed slide model.** `PresentationEditor.Create`,
  `CreateAsync` and `CreateToFileAsync` build a deck from `PptxSlide` values — a title and bullet
  lines per slide. Previously `PresentationEditor` could count slides, extract text and replace
  text, but there was no way to obtain a `.pptx` in the first place. As with the DOCX equivalent,
  content comes from **data** rather than a template, so there is nothing to escape and no source
  file to edit.

  ```csharp
  byte[] deck = PresentationEditor.Create(new[]
  {
      PptxSlide.Titled("Q3 Results", "Revenue up 12%", "Costs flat"),
      PptxSlide.Titled("Outlook", "Hiring 3 engineers"),
  });
  ```

  ([faac36b](https://github.com/Ank-KhoaHo/DocToolkit/commit/faac36b194963f993716ef4bc4a47f15e3f35a5d),
  [5fc9738](https://github.com/Ank-KhoaHo/DocToolkit/commit/5fc97388e1e4a8386aba0f9d6de143f418deb5c4),
  [6756955](https://github.com/Ank-KhoaHo/DocToolkit/commit/675695556f048a5a5cb471f2ddb8a88de38b7d36),
  [f832b17](https://github.com/Ank-KhoaHo/DocToolkit/commit/f832b17a5fc754700be451039109e155711dd943),
  [321d3e0](https://github.com/Ank-KhoaHo/DocToolkit/commit/321d3e0d5ff9d5936f08ef625fac8f3152a4cfa3))

* **Extensions: `IDocxEditor` regains parity with `DocxEditor`.**
  `FillRows`/`FillRowsAsync` — repeating table rows — have been missing from the interface since
  **0.4.0** shipped them in the core package, and `ReplaceImage`/`ReplaceImageAsync` since **0.5.0**.
  Both features were therefore unreachable through dependency injection entirely, and available only
  through the static API. Both are now present, along with `Create`/`CreateAsync` for the block model
  added in 0.10.0.

  **If you use `IDocxEditor` and worked around either gap by calling `DocToolkit.DocxEditor`
  statically, you can now inject it instead.** File-path overloads remain deliberately unmirrored,
  as they are for every interface in this package.
  ([0c3d90e](https://github.com/Ank-KhoaHo/DocToolkit/commit/0c3d90e8c4ee0142e12e2fe15eeb38ad106f1b92))

## [0.10.0](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.9.0...v0.10.0) (2026-08-07)


### Added

* **Core: create a DOCX from scratch, from a typed block model.** `DocxEditor.Create`,
  `CreateAsync` and `CreateToFileAsync` build a document from `DocxBlock` values — headings,
  paragraphs, tables, and inline images with alt text. Previously the only way to obtain a `.docx`
  was to convert HTML. This path takes content from **data** rather than markup, so there is no
  HTML to escape and a value containing `<` cannot corrupt the document's structure.

  ```csharp
  byte[] docx = DocxEditor.Create(new[]
  {
      DocxBlock.Heading("Quarterly Report", 1),
      DocxBlock.Paragraph("Revenue rose 12% against a flat cost base."),
      DocxBlock.Table(new[] { "Region", "Revenue" },
                      new[] { new object[] { "EMEA", 1200 } }),
      DocxBlock.Image(logoBytes, widthPoints: 120, altText: "Contoso logo"),
  });
  ```

  ([5c30c52](https://github.com/Ank-KhoaHo/DocToolkit/commit/5c30c52ce3991b3def51402334d959750a6c034e),
  [0c785b0](https://github.com/Ank-KhoaHo/DocToolkit/commit/0c785b04524bab49da012d4a188bded722915b31),
  [60b484e](https://github.com/Ank-KhoaHo/DocToolkit/commit/60b484e0d8deb468b4f5f1f160ae4c330a710ea1),
  [b8efc54](https://github.com/Ank-KhoaHo/DocToolkit/commit/b8efc54a3f7d07cf7c6c44b1c58ea45aaa2ad2dd),
  [5353d20](https://github.com/Ank-KhoaHo/DocToolkit/commit/5353d20536b6d41660c9731ef8285a28fc46ec83),
  [57a8add](https://github.com/Ank-KhoaHo/DocToolkit/commit/57a8add09514bfd7b6ad11af44cfa00a0e2f595d),
  [20a1602](https://github.com/Ank-KhoaHo/DocToolkit/commit/20a1602c277ad0b702cc81b43b32839d7091442d))


### Fixed

* **Core:** an oversized image produced a **corrupt document instead of an error**. The drawing
  extent overflowed to a negative value, giving a file with four schema violations and no exception
  thrown. Image dimensions are now bounded at 2,147,483,647 EMU per side (about 2,348 inches).

  **This also changes `DocxEditor.ReplaceImage`**, which shares the same size arithmetic: a call
  with an out-of-range `widthPoints`/`heightPoints` that returned a document in 0.9.0 now throws
  `ArgumentOutOfRangeException`. The bound applies to a dimension derived from the aspect ratio as
  well as to one you supply.
  ([030e321](https://github.com/Ank-KhoaHo/DocToolkit/commit/030e3217ccced3963678ae8469922cf8b173685b))


### Changed

* **Core:** `OfficeIMO.Word.Pdf` updated to 3.1.0. Generated PDFs now **embed a subset of the font
  they use** rather than relying on the viewer to substitute one, so text renders consistently
  wherever the file is opened, and remains selectable and searchable. The cost is file size: a
  short document grows by roughly 130 KB, because the font travels with it. The resolved dependency
  closure shrank from 18 packages to 16.
  ([fe2ad4a](https://github.com/Ank-KhoaHo/DocToolkit/commit/fe2ad4ab81bee52c21f69b0260e0139f51d2d580))

* **Extensions:** `THIRD-PARTY-NOTICES.txt` omitted `Microsoft.Extensions.Primitives` entirely. It
  reaches this package through `Microsoft.Extensions.Options` rather than through `Ank.DocToolkit`,
  so the notices' pointer to the core package's file never covered it and it was attributed nowhere.
  **If you have run an open-source attribution or licence-compliance review against an earlier
  version, it was incomplete** — that package is MIT, as is the rest of the closure. Both packages'
  notices are now generated from the resolved dependency graph and checked in CI, rather than
  maintained by hand.

## [0.9.0](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.8.0...v0.9.0) (2026-08-06)


### Added

* **Extensions:** `DocToolkitOptions.RemoteImage` carries the `RemoteImageOptions` bounds applied
  when `AllowRemoteImageDownload` is `true` — per-fetch timeout, byte cap, host allow-list and the
  private-address block. It is get-only and configured in place, so a restrictive default cannot be
  lost by assigning an object that missed one. The bool keeps its meaning as the only switch
  deciding whether anything is fetched at all. Requires `Ank.DocToolkit` 0.8.0 or later.

### Changed

* **Extensions:** remote image download through `AddDocToolkit` now refuses loopback, private and
  link-local addresses by default, inheriting the guard core gained in 0.8.0. **This affects anyone
  registering `AllowRemoteImageDownload = true` to fetch images from an intranet host**: the
  conversion still succeeds, but that image is now silently left out, with no exception raised. Add
  `o.RemoteImage.AllowPrivateAddresses = true` to restore the old reach. Every fetch is also now
  bounded by a 10-second timeout and a 5 MB cap. This is still not a complete SSRF defence; see
  `SECURITY.md`.

## [0.8.0](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.7.0...v0.8.0) (2026-08-06)

### Added

* **Core:** `RemoteImageOptions`, and four `ConvertAsync` overloads taking it on
  `HtmlToDocxConverter` and `HtmlToPdfConverter`. They bound the remote-image opt-in with a
  per-fetch timeout, a byte cap, an optional host allow-list and a private-address block, all
  restrictive by default. See **Changed** below — this also alters what the existing
  `allowRemoteImageDownload: true` reaches.
* **Core:** file-path overloads on `DocxEditor`, `PresentationEditor` and `WorkbookEditor`, so a
  capability can now be reached as `byte[]`, as a `Stream`, or by path.
* **Extensions:** `IWorkbookEditor` gained `SheetNames`, `SheetNamesAsync`, `ReadSheet` and
  `ReadSheetAsync`, restoring 1:1 parity with `WorkbookEditor` after the core package shipped them
  in 0.7.0.

### Changed

* **Core:** the remote-image opt-in (`allowRemoteImageDownload: true`, and the new
  `RemoteImageOptions` overloads on `HtmlToDocxConverter`/`HtmlToPdfConverter`) now refuses
  loopback, private and link-local addresses by default, including `169.254.169.254` — the cloud
  metadata endpoint. **This affects anyone currently converting intranet-hosted `<img>` markup with
  `allowRemoteImageDownload: true`**: the conversion still succeeds, but that image is now silently
  left out, with no exception raised. Pass `new RemoteImageOptions { AllowPrivateAddresses = true }`
  in place of the bool overload to restore the old reach. The opt-in is also now bounded by a
  10-second per-fetch timeout, a 5 MB cap enforced on bytes actually read (never on a `Content-Length`
  header), an optional host allow-list, and refuses every scheme but `http`/`https` — closing a
  `file://` local-disclosure path that existed before. This is still not a complete SSRF defence;
  see the package README.

## [0.7.0](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.6.0...v0.7.0) (2026-08-04)


### Added

* **core:** ship a package icon on both packages ([b5dfdff](https://github.com/Ank-KhoaHo/DocToolkit/commit/b5dfdff321f40028326f116cb5dd8a0bd18c71b5))
* **core:** ship a package icon on both packages ([ba1424f](https://github.com/Ank-KhoaHo/DocToolkit/commit/ba1424f1591e809598a7279f66e9bd52b2de8be6))


### Changed

* close D11, the package icon shipped ([b1d396d](https://github.com/Ank-KhoaHo/DocToolkit/commit/b1d396df86ef49178102e5d4d1c89d7868d5a6ce))
* design for splitting the samples per capability ([2cc29e3](https://github.com/Ank-KhoaHo/DocToolkit/commit/2cc29e3eae609665a2b98e150de67ced0d00a913))
* design for splitting the samples per capability ([7966887](https://github.com/Ank-KhoaHo/DocToolkit/commit/796688722f30af41441c4b1245c9766009383c69))
* implementation plan for splitting the samples per capability ([de2ae51](https://github.com/Ank-KhoaHo/DocToolkit/commit/de2ae5144aeafd9e873275dfa83038338a1a8c01))
* record JPEG sample-coverage gap in enhancement backlog ([8b99058](https://github.com/Ank-KhoaHo/DocToolkit/commit/8b99058671391b9329c12c0e70a44ced2b13eee2))
* record the missing NuGet package icon as D11 ([73275af](https://github.com/Ank-KhoaHo/DocToolkit/commit/73275af19ee7174cdfc98b4a7067e18e7056764b))
* record the missing NuGet package icon as D11 ([f224e39](https://github.com/Ank-KhoaHo/DocToolkit/commit/f224e3915015aa51fd497e382e3b50d1424e7a74))
* **samples:** add the DocxImages sample ([eab3f44](https://github.com/Ank-KhoaHo/DocToolkit/commit/eab3f443480907b9182bb5f397528434f0e67b4c))
* **samples:** add the DocxTemplating sample ([ed38169](https://github.com/Ank-KhoaHo/DocToolkit/commit/ed38169373d0d718abfda4085e32c3c7c53bef59))
* **samples:** add the HtmlConversion sample ([0e4a1a8](https://github.com/Ank-KhoaHo/DocToolkit/commit/0e4a1a862dc4fd242188274b8ee9181b2cb6e1c3))
* **samples:** add the Presentations sample ([8070797](https://github.com/Ank-KhoaHo/DocToolkit/commit/80707971e508a8791edefa3dc60d594eb01f3839))
* **samples:** add the Spreadsheets sample ([7451242](https://github.com/Ank-KhoaHo/DocToolkit/commit/7451242f6a9293c406eba0fcda03224b80e75500))
* **samples:** correct ExtractText and FillRows/ReplaceText claims ([382177f](https://github.com/Ank-KhoaHo/DocToolkit/commit/382177fae89a6eb95e5540d3f0f68d8e03de4437))
* **samples:** fix "one per capability" wording in layout blocks ([5b309b7](https://github.com/Ank-KhoaHo/DocToolkit/commit/5b309b72197a094d33a1163fddb36b4e3ecfdcbb))
* **samples:** rename MinimalApiSample to MinimalApi ([185ac61](https://github.com/Ank-KhoaHo/DocToolkit/commit/185ac61c754f6a7b80c812696410759072bdcfc8))
* **samples:** retire ConsoleSample in favour of per-capability samples ([d3c4b9b](https://github.com/Ank-KhoaHo/DocToolkit/commit/d3c4b9b7815d372546ee9ff5321f5d6b68b12087))
* **samples:** split the samples per capability ([8fd9f09](https://github.com/Ank-KhoaHo/DocToolkit/commit/8fd9f0969c9bcd94c1ed9443cb4a55aaf901cf5d))

## [0.6.0](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.5.0...v0.6.0) (2026-08-04)


### Added

* **core:** list a workbook's sheets with SheetNames ([b26244b](https://github.com/Ank-KhoaHo/DocToolkit/commit/b26244b0ad781503832037d526e2bb6834e923b0))
* **core:** read a whole sheet with ReadSheet ([1d0dbdd](https://github.com/Ank-KhoaHo/DocToolkit/commit/1d0dbdd03f8d96f541d910139ddfe25aff77cdf3))
* **core:** read XLSX sheets in bulk ([d2a72fc](https://github.com/Ank-KhoaHo/DocToolkit/commit/d2a72fcf2e7abbccc2fcc5840d2ee1fdeb38ff09))


### Fixed

* **core:** cap ReadSheet's used-range at 2,000,000 cells ([061d7e4](https://github.com/Ank-KhoaHo/DocToolkit/commit/061d7e4746df4e57361b0632b80bc4e10a3cf80e))


### Changed

* close A16 - verified in Word, the location is cosmetic ([52491d7](https://github.com/Ank-KhoaHo/DocToolkit/commit/52491d7bb94aeec97eca342a62f547e8e7c361b1))
* **core:** close the culture-rule gap in ReadSheetAsync and ReadCell ([fd2930c](https://github.com/Ank-KhoaHo/DocToolkit/commit/fd2930cb8fe95d3a3f2ebab1724fef4e243edb31))
* **core:** correct three inaccuracies found in final review ([984697d](https://github.com/Ank-KhoaHo/DocToolkit/commit/984697d103d65adfeb4351438e2c6c93e9ba580a))
* **core:** document XLSX bulk reading ([d4176b3](https://github.com/Ank-KhoaHo/DocToolkit/commit/d4176b3cfc38ae5af9c4ccde1ac69534be0ad485))
* **core:** update test counts for rebased XLSX reading tests ([d72fecb](https://github.com/Ank-KhoaHo/DocToolkit/commit/d72fecb574b32e3fecc6550f4160eec1011e99f7))
* correct the used-range semantics in the XLSX reading design ([51e7d90](https://github.com/Ank-KhoaHo/DocToolkit/commit/51e7d90193a7f9ced611b0dfe3048a2f34393c05))
* design for XLSX bulk reading (A3, reading slice) ([1441801](https://github.com/Ank-KhoaHo/DocToolkit/commit/1441801c02af55a9a820c97d2538e0bec6f01fbd))
* design for XLSX bulk reading (A3, reading slice) ([dfcb4a3](https://github.com/Ank-KhoaHo/DocToolkit/commit/dfcb4a3e4bf51ca170bc588079ea8601fa9865cb))
* implementation plan for XLSX bulk reading ([4b9e824](https://github.com/Ank-KhoaHo/DocToolkit/commit/4b9e824974b3677424d40b637fbc8f4deb125b21))
* keep the version below 1.0.0 ([3074725](https://github.com/Ank-KhoaHo/DocToolkit/commit/3074725a1f9dcab6ca561d1bca0c48b687216c36))
* keep the version below 1.0.0 ([c6a446c](https://github.com/Ank-KhoaHo/DocToolkit/commit/c6a446c613a3a32cfdc588b896ea441f49daf49e))
* record where ReplaceImage puts image parts (closed — verified in Word) ([edbbe2d](https://github.com/Ank-KhoaHo/DocToolkit/commit/edbbe2da6e9e2f7639541e3a4342308f86448bfa))
* record where ReplaceImage puts image parts, and why it was not changed ([8aedb52](https://github.com/Ank-KhoaHo/DocToolkit/commit/8aedb526718292decc834ce0c4411f4f20c261ba))
* **samples:** demonstrate image placeholders in ConsoleSample ([6901d3b](https://github.com/Ank-KhoaHo/DocToolkit/commit/6901d3b0ce78d982f89bd41ebfb160211ae66651))
* **samples:** demonstrate image placeholders in ConsoleSample ([5940cf0](https://github.com/Ank-KhoaHo/DocToolkit/commit/5940cf09f5c0791736aac1e4266937ce517eec36))

## [0.5.0](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.4.0...v0.5.0) (2026-08-03)


### Added

* **core:** add DocxEditor.ReplaceImage ([a8ead4f](https://github.com/Ank-KhoaHo/DocToolkit/commit/a8ead4fb5295d1a6f1be5562d078c792d6566575))
* **core:** add DocxEditor.ReplaceImageAsync, and document images ([869fe14](https://github.com/Ank-KhoaHo/DocToolkit/commit/869fe14cfa0a1aac2ff8dbf3673af76be12735c2))
* **core:** image placeholders for DOCX templates ([f1d7618](https://github.com/Ank-KhoaHo/DocToolkit/commit/f1d7618a21920ff246c212c35904e5dfaa515c96))
* **core:** read image format and size from the header, and resolve to EMUs ([dc1443b](https://github.com/Ank-KhoaHo/DocToolkit/commit/dc1443b92fd266cdb5b3eacf4f7cd4df2d8ebcc4))


### Fixed

* **core:** document every DrawingFactory parameter, and record why CI caught it ([4dfc227](https://github.com/Ank-KhoaHo/DocToolkit/commit/4dfc227ba3c49fc0a538062a61384976a5fd06ae))
* **core:** drop the flaky timing assertion from the cancellation test ([d59815e](https://github.com/Ank-KhoaHo/DocToolkit/commit/d59815e07f44c490e75d5b083ccd9eb205c2f0f3))
* **core:** drop the flaky timing assertion from the cancellation test ([a3394df](https://github.com/Ank-KhoaHo/DocToolkit/commit/a3394df8d0d5dd39bab57249dafa40250462a3bd))


### Changed

* add the image-placeholders implementation plan ([19308d2](https://github.com/Ank-KhoaHo/DocToolkit/commit/19308d2b0f68ca469389f14fd98dab11e0dce935))
* design image placeholders for DOCX templates ([f4ac790](https://github.com/Ank-KhoaHo/DocToolkit/commit/f4ac7900b66ce74182aea10c4a29f757bf3197de))
* design image placeholders for DOCX templates ([54b2a3e](https://github.com/Ank-KhoaHo/DocToolkit/commit/54b2a3efe7bca74a0c9f510f27796ecbcc313261))
* **samples:** demonstrate repeating table rows in ConsoleSample ([de80d4b](https://github.com/Ank-KhoaHo/DocToolkit/commit/de80d4bd232ebf9e9e40aecdb47f5c8a46d84261))
* **samples:** demonstrate repeating table rows in ConsoleSample ([1f857f8](https://github.com/Ank-KhoaHo/DocToolkit/commit/1f857f887a668e4c379f1366146083909a2b02f2))

## [0.4.0](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.3.12...v0.4.0) (2026-08-03)


### Added

* **core:** add DocxEditor.FillRows for repeating table rows ([432e4f3](https://github.com/Ank-KhoaHo/DocToolkit/commit/432e4f30cd795bc865d9b07b70e5b6a6ddb0f6ff))
* **core:** add DocxEditor.FillRowsAsync ([6ff7061](https://github.com/Ank-KhoaHo/DocToolkit/commit/6ff70610ffb15ea6949cca7112e43e50d306cd8b))
* **core:** find template table rows without descending into nested tables ([7a4f39d](https://github.com/Ank-KhoaHo/DocToolkit/commit/7a4f39db5bffe7f3be3652cae027da7c21f8bbaa))
* **core:** repeating table rows for DOCX templates ([0a25d50](https://github.com/Ank-KhoaHo/DocToolkit/commit/0a25d50db75190dc18f31a28263952362a6eaf29))


### Changed

* add the repeating-table-rows implementation plan ([5d2f1cb](https://github.com/Ank-KhoaHo/DocToolkit/commit/5d2f1cb91890d35899768b1b85e616abc16cae1d))
* add the repeating-table-rows implementation plan ([8b87f34](https://github.com/Ank-KhoaHo/DocToolkit/commit/8b87f342687456ab31c565d6ac74e680fc59b280))
* add the repeating-table-rows implementation plan ([b82aadd](https://github.com/Ank-KhoaHo/DocToolkit/commit/b82aadd1841eee68d0fb74ece81ae052bb9ef1e3))
* **core:** document repeating table rows, and guard them offline ([e33b0c7](https://github.com/Ank-KhoaHo/DocToolkit/commit/e33b0c7a1da001470b319a62d795ec0a7f786e44))
* exempt merge commits from the Conventional Commits guard ([5a313d1](https://github.com/Ank-KhoaHo/DocToolkit/commit/5a313d14ea4876e8f11dc4e117bfd9314121a2fd))
* exempt merge commits from the Conventional Commits guard ([4abaf2a](https://github.com/Ank-KhoaHo/DocToolkit/commit/4abaf2a66d5619d2a209d7eea0c3b5e3defcc5a1))
* record the flaky cancellation test as B10 ([8ffa6f0](https://github.com/Ank-KhoaHo/DocToolkit/commit/8ffa6f00e6a203788c5789b11ec94a9347e96478))
* stop auto-merging the Release PR ([7b0f0ac](https://github.com/Ank-KhoaHo/DocToolkit/commit/7b0f0acab67ea26005375f017956223636308f26))
* stop auto-merging the Release PR ([02f2ace](https://github.com/Ank-KhoaHo/DocToolkit/commit/02f2acedb4a9458fa7abcf3f8cf5926a4c0df803))

## [0.3.12](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.3.11...v0.3.12) (2026-08-03)


### Changed

* design repeating table rows for DOCX templates ([eb447d3](https://github.com/Ank-KhoaHo/DocToolkit/commit/eb447d360e8f51dd1fb13701f07c0e37cbe3f1b0))
* design repeating table rows for DOCX templates ([8193879](https://github.com/Ank-KhoaHo/DocToolkit/commit/8193879aed0386f028b267afadb498f5e392f174))

## [0.3.11](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.3.10...v0.3.11) (2026-08-03)


### Changed

* refresh the backlog against what shipped today ([570ba61](https://github.com/Ank-KhoaHo/DocToolkit/commit/570ba61ec6422729f782fabf33445c6bfb261797))

## [0.3.10](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.3.9...v0.3.10) (2026-08-03)


### Changed

* bump googleapis/release-please-action from 4 to 5 ([8bbe072](https://github.com/Ank-KhoaHo/DocToolkit/commit/8bbe0723a75504d9e23f6dedeaf132aa8cf66c35))

## [0.3.9](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.3.8...v0.3.9) (2026-08-03)


### Changed

* record the single-branch verification outcome ([696c3bc](https://github.com/Ank-KhoaHo/DocToolkit/commit/696c3bc0c2e4e28ac459cb94f6881c5d897bf688))
* record the single-branch verification outcome ([1fed371](https://github.com/Ank-KhoaHo/DocToolkit/commit/1fed3718cdef5214a61ea8cddfe403f56d6763a2))

## [0.3.8](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.3.7...v0.3.8) (2026-08-03)


### Changed

* bump five GitHub Actions across major versions ([963ec57](https://github.com/Ank-KhoaHo/DocToolkit/commit/963ec57b79584ae3c1e2413b4354cf20c81494fc))
* bump five GitHub Actions across major versions ([8f42d36](https://github.com/Ank-KhoaHo/DocToolkit/commit/8f42d36dcfc0284acb092423590f49fc7800d7d3))
* fold the two remaining docs.yml action majors in ([ae2873e](https://github.com/Ank-KhoaHo/DocToolkit/commit/ae2873ebc213e2ed6abe4588039c13cb8390f9f1))

## [0.3.7](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.3.6...v0.3.7) (2026-08-03)


### Fixed

* stop asserting byte equality on two ClosedXML saves ([5edce27](https://github.com/Ank-KhoaHo/DocToolkit/commit/5edce2746189fce8a34edf2ae6e4527fcb9b4cc3))
* stop asserting byte equality on two ClosedXML saves ([2697f13](https://github.com/Ank-KhoaHo/DocToolkit/commit/2697f139d25538f0b930aa78348feb71c47fa2f0))

## [0.3.6](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.3.5...v0.3.6) (2026-08-03)


### Changed

* float the samples to the newest published release ([87b01d0](https://github.com/Ank-KhoaHo/DocToolkit/commit/87b01d028bafab85b0042e8d9842cdf94a515706))
* float the samples to the newest published release ([9d2b564](https://github.com/Ank-KhoaHo/DocToolkit/commit/9d2b564c8971433d45cec1eef3e47ee17bc79f73))

## [0.3.5](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.3.4...v0.3.5) (2026-08-03)


### Changed

* keep major dependency bumps out of grouped PRs ([fea3e49](https://github.com/Ank-KhoaHo/DocToolkit/commit/fea3e49453c00d45983b932ff0ec6ea2861c649c))
* keep major dependency bumps out of grouped PRs ([468307b](https://github.com/Ank-KhoaHo/DocToolkit/commit/468307bf5f6f899d58424c595d50938ad8c6355b))
* repeat the src guards in the tests update block ([d0bf6c7](https://github.com/Ank-KhoaHo/DocToolkit/commit/d0bf6c7021fb1bb7487dc68e0d2d594c7c2fb93a))

## [0.3.4](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.3.3...v0.3.4) (2026-08-03)


### Changed

* bump SixLabors.Fonts to 1.0.1 ([61e462a](https://github.com/Ank-KhoaHo/DocToolkit/commit/61e462a4859c51799a336a3a9421a73fe9aa859f))
* bump SixLabors.Fonts to 1.0.1 ([c38eeea](https://github.com/Ank-KhoaHo/DocToolkit/commit/c38eeead26a7b0b87af68257bb200de058d5b43a))

## [0.3.3](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.3.2...v0.3.3) (2026-08-03)


### Changed

* account for admin bypass of branch protection ([6481f62](https://github.com/Ank-KhoaHo/DocToolkit/commit/6481f6225a320d24f7e024a94e5db205b5bd1536))
* add the dependency-automation implementation plan ([d7d1588](https://github.com/Ank-KhoaHo/DocToolkit/commit/d7d158840dd28b03dbce3e3c99867bdda0307c49))
* add the enhancement backlog and dependency-automation design ([d408789](https://github.com/Ank-KhoaHo/DocToolkit/commit/d408789a407f129e2e59bbee45533f19c39f4fa7))
* add the single-branch collapse implementation plan ([11718a0](https://github.com/Ank-KhoaHo/DocToolkit/commit/11718a0dcafa5cc7cfe5f803d3a1b48dea8533e8))
* assert the resolved graph matches the committed lockfiles ([3715ba9](https://github.com/Ank-KhoaHo/DocToolkit/commit/3715ba9b11a05b4c7ed6dd938e8b33b3ae8ad58e))
* auto-merge the Release PR so every merge releases ([cd5edf8](https://github.com/Ank-KhoaHo/DocToolkit/commit/cd5edf897442d5ad81f9352c9c579612b9d32513))
* automate dependency updates and lock the shipped dependency graph ([3ff0a9f](https://github.com/Ank-KhoaHo/DocToolkit/commit/3ff0a9f2165fd5faeda8d925c2261b973e7de866))
* bump the sample package floors so the canary re-arms ([aedffc9](https://github.com/Ank-KhoaHo/DocToolkit/commit/aedffc9aa1b9c34a567fac1c18c394b7f8f5a747))
* collapse to a single main branch with automatic releases ([f26d492](https://github.com/Ank-KhoaHo/DocToolkit/commit/f26d49215374ace4ab16651aabc05ec30af55173))
* describe the single-branch model ([8b71c4b](https://github.com/Ank-KhoaHo/DocToolkit/commit/8b71c4bea07860502c548cd82db67e3c74fdbc52))
* design the collapse to a single main branch ([e50bc37](https://github.com/Ank-KhoaHo/DocToolkit/commit/e50bc3752a0d7a726525934061e4aa0361339b46))
* lock the resolved dependency graph of both packages ([26fe5c4](https://github.com/Ank-KhoaHo/DocToolkit/commit/26fe5c4efd5f95a2d5a60dff313b0fb953873e9e))
* point Dependabot at main ([cd1e54e](https://github.com/Ank-KhoaHo/DocToolkit/commit/cd1e54ee47ee94ea1937952c30a4f373c5606395))
* propose dependency updates automatically ([b342bc1](https://github.com/Ank-KhoaHo/DocToolkit/commit/b342bc1a2fcdd0079e748d820760c8dd45dcc165))
* record that Dependabot only activates once config reaches main ([c2a59d2](https://github.com/Ank-KhoaHo/DocToolkit/commit/c2a59d29ca772b25cfbcf782ac08688680357eb0))
* record the IDE-restore race that fakes a passing lockfile guard ([53a7760](https://github.com/Ank-KhoaHo/DocToolkit/commit/53a77604318fc5ada16bf0f6d71939474d5ad66d))
* remove the promote machinery ([e160d20](https://github.com/Ank-KhoaHo/DocToolkit/commit/e160d2085e9c96844dc32d5d8aed8e7403eb329a))
* require a conventional message on the migration merge ([9facd75](https://github.com/Ank-KhoaHo/DocToolkit/commit/9facd7500ba5a26cef23383a2f0a60d5e2ef8483))
* restore CLAUDE.md, docs and spike onto main ([e49e189](https://github.com/Ank-KhoaHo/DocToolkit/commit/e49e189889c5e840832ae4dc2b871a9a6996c34f))

## [0.3.2](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.3.1...v0.3.2) (2026-08-03)


### Changed

* tighten the root README and drop the AutoLnD reference ([8a42e41](https://github.com/Ank-KhoaHo/DocToolkit/commit/8a42e413588c3fd0545ca059cc0276e248c7ca95))

## [0.3.1](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.3.0...v0.3.1) (2026-08-03)


### Fixed

* move PDF fixtures off spike/out so main can build ([a6b8604](https://github.com/Ank-KhoaHo/DocToolkit/commit/a6b860409c2de343c7373477ab5ce877341ab657))


### Changed

* add the develop-to-main promote script and its test ([c389c1c](https://github.com/Ank-KhoaHo/DocToolkit/commit/c389c1cfcfa6ed893c44e9edbfe5477767204213))
* add the implementation plan for the two-branch model ([4dce9fc](https://github.com/Ank-KhoaHo/DocToolkit/commit/4dce9fc1d8d07d72d7fdc8fbd4c6f3e522d327dd))
* address second-pass review findings on promote-to-main ([a140c3d](https://github.com/Ank-KhoaHo/DocToolkit/commit/a140c3de066773edbec2557a1296b823ab2cd96d))
* amend the branching plan for state that moved under it ([c777c23](https://github.com/Ank-KhoaHo/DocToolkit/commit/c777c23ad79a159c410cc2ec4c0c58ec50663cb0))
* amend the plan from the whole-branch review ([ed36ca0](https://github.com/Ank-KhoaHo/DocToolkit/commit/ed36ca0067f71446df913e889a9c83d4c2f45455))
* correct branching-model doc inaccuracies from whole-branch review ([41d4e12](https://github.com/Ank-KhoaHo/DocToolkit/commit/41d4e128fbf552a1b88f618d5389e494b3df3606))
* correct the stale test counts ([6bb6466](https://github.com/Ank-KhoaHo/DocToolkit/commit/6bb64668d1d830e505a33adbd90c62ca015396c0))
* correct why scripts/ stays on main ([e6ae369](https://github.com/Ank-KhoaHo/DocToolkit/commit/e6ae369d186b40230a9e5c9e4e1050bc025149f0))
* describe the two-branch model and its one dangerous operation ([9398f55](https://github.com/Ank-KhoaHo/DocToolkit/commit/9398f5568e2abc94018c359f7a6c24166ee985b0))
* design a two-branch model separating releases from development ([1b50fef](https://github.com/Ank-KhoaHo/DocToolkit/commit/1b50fefc72e34b52553e0b894fef1a0cea558ea6))
* guard against shell injection via github.head_ref in branch-policy ([d1da75a](https://github.com/Ank-KhoaHo/DocToolkit/commit/d1da75aabc70d8c9abf2d2f4c7ada9c9366bbdbf))
* harden promote-to-main merge, purge and test guarantees ([77bdae8](https://github.com/Ank-KhoaHo/DocToolkit/commit/77bdae85dcd5ba54a393499692935f6b4da1b27c))
* point contributors at develop and drop the dev-phase rows ([1df5126](https://github.com/Ank-KhoaHo/DocToolkit/commit/1df5126a530d64db37992dee280d13bee152e406))
* run on develop, and keep main release-only ([968501d](https://github.com/Ank-KhoaHo/DocToolkit/commit/968501d771f32a28468fb5c8dfb10e08e0491189))
* Update branch merge/rebase guidance for develop-based workflow ([dd2d089](https://github.com/Ank-KhoaHo/DocToolkit/commit/dd2d08939bb85ab543ea2d600e760b70aadf774d))

## [0.3.0](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.2.2...v0.3.0) (2026-08-03)

### Added

- **Extensions:** `Stream`-based async members on all six injectable interfaces, so a document can
  be read from and written to a request body, a response body or a file without being buffered
  into a caller-visible `byte[]`. The core package has had these since 0.2.0; only the `byte[]`
  surface was reachable through DI until now.
  ([21890a2](https://github.com/Ank-KhoaHo/DocToolkit/commit/21890a2b9041863ada07a9ac44deba9e98b06c4c))

  | Interface | Added |
  |---|---|
  | `IDocxEditor` | `ReplaceTextAsync`, `ExtractTextAsync` (two overloads) |
  | `IPresentationEditor` | `SlideCountAsync`, `ExtractTextAsync`, `ReplaceTextAsync` |
  | `IWorkbookEditor` | `CreateAsync`, `ReadCellAsync`, `SetCellAsync` |
  | `IHtmlToDocxConverter` | `ConvertAsync(string, Stream, CancellationToken)` |
  | `IHtmlToPdfConverter` | `ConvertAsync(string, Stream, CancellationToken)` |
  | `IDocxToPdfConverter` | `ConvertAsync(Stream, Stream, CancellationToken)` |

  Each new member delegates to the identically-shaped core static method. The two HTML converters
  thread the registration-time `DocToolkitOptions.AllowRemoteImageDownload`, exactly as their
  `byte[]` counterparts do — remote image download stays a composition-time decision, never a
  per-call argument.

  The file-path helpers (`ConvertToFileAsync`, `ConvertFile`) and the per-call
  `allowRemoteImageDownload` argument remain deliberately absent from these interfaces.

### Fixed

- **Extensions:** the package referenced `Ank.DocToolkit` at floor `[0.1.0, )`. NuGet resolves a
  minimum-version range to the *lowest* satisfying version, so the package built against a core
  release predating the `Stream` API it wraps. Floor is now `[0.2.0, )`.
  ([51646e2](https://github.com/Ank-KhoaHo/DocToolkit/commit/51646e2c17fecffcb7ae6e410d25a70d11180c58))

### Changed

- **Extensions:** `IHtmlToPdfConverter.ConvertAsync` and `IDocxToPdfConverter.ConvertAsync` now
  document that their output is written as it is rendered rather than assembled first — so a
  failure part-way through leaves partial output on the destination, and against an HTTP response
  body the status and headers are already committed.
  ([43ff9ac](https://github.com/Ank-KhoaHo/DocToolkit/commit/43ff9acab2722b09d45f3d8f15eaeda97b329c97))
- **Extensions:** the package README no longer claims the interfaces mirror the static API
  "one-for-one" — it now names what is deliberately excluded and why.
  ([a570e2d](https://github.com/Ank-KhoaHo/DocToolkit/commit/a570e2de98c69d3f887d5895b2e85060e67156dc),
  [5ed5596](https://github.com/Ank-KhoaHo/DocToolkit/commit/5ed55969fcf03bd092864edb33f5b74eb68f7d1a))
- The API-reference site is now linked from the README.
  ([812e433](https://github.com/Ank-KhoaHo/DocToolkit/commit/812e4339dd06776ed5eb773e80ec792bdcbbd32a))

### Upgrading

Adding members to a shipped interface is a **source-breaking change for anyone who implements one
of these six interfaces themselves** — a hand-written test double or adapter will no longer compile
until the new members are added to it. Consumers who only *inject* the interfaces, and those using
a mocking framework (Moq, NSubstitute, FakeItEasy) that generates implementations at runtime, are
unaffected. This is accepted deliberately while the package is pre-1.0.

One source-level nuance: on `IHtmlToDocxConverter` and `IHtmlToPdfConverter`, a call written as
`ConvertAsync(html, default)` is now ambiguous between the `CancellationToken` and `Stream`
overloads. Name the argument (`ConvertAsync(html, ct: default)`) to disambiguate. `null` is
unaffected.

## [0.2.2](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.2.1...v0.2.2) (2026-08-03)


### Added

- Runnable sample projects (`samples/ConsoleSample`, `samples/MinimalApiSample`), referencing
  the published packages and built by the existing CI alongside the libraries.
- A DocFX-generated API-reference site (`docfx/`), published to GitHub Pages on every
  successful release via `.github/workflows/docs.yml`.


### Fixed

* add release-published trigger, correct docs from re-review ([9611ee4](https://github.com/Ank-KhoaHo/DocToolkit/commit/9611ee4999ba053f8cd9b9373ae5adc42bc5c81d))
* address final whole-branch review findings ([c86cabb](https://github.com/Ank-KhoaHo/DocToolkit/commit/c86cabb8b2a3d68cc4ee82fb2e6a956809a3da15))


### Changed

* add release-please config, seeded at the current 0.2.1 ([4a4059b](https://github.com/Ank-KhoaHo/DocToolkit/commit/4a4059bc4cff806720bfe314f9837da2820c9cd0))
* add release-please workflow ([d3360e0](https://github.com/Ank-KhoaHo/DocToolkit/commit/d3360e0e0b64f3c0eed43725a1ccdeb54938fe2c))
* correct two inaccuracies in CLAUDE.md's release-please description ([ff3bfc4](https://github.com/Ank-KhoaHo/DocToolkit/commit/ff3bfc44b5714e56c7daefc3c81258ed8bb42b59))
* describe the release-please flow and commit-format requirement ([d0d3c1a](https://github.com/Ank-KhoaHo/DocToolkit/commit/d0d3c1a294d54c71d40fc22cbae721094ae9a231))
* enforce Conventional Commits format on every PR commit ([14683b6](https://github.com/Ank-KhoaHo/DocToolkit/commit/14683b616d52d4cd24566bf0b4e7aa8b80197db9))

## [0.2.1] - 2026-07-31

### Added
- **Extensions:** first release of `Ank.DocToolkit.Extensions.DependencyInjection`.
  `services.AddDocToolkit()` registers six interfaces — `IHtmlToDocxConverter`,
  `IDocxToPdfConverter`, `IHtmlToPdfConverter`, `IDocxEditor`, `IWorkbookEditor`,
  `IPresentationEditor` — mirroring the core static API 1:1, plus `DocToolkitOptions` for
  configuring `AllowRemoteImageDownload` once instead of per call.

### Changed
- **Core:** republished at this version with no functional change, to align version numbers with
  the extensions package. From this release onward, both packages version and release together,
  at the same version, from a single tag.
- Release pipeline merged into one workflow (`release.yml`) that packs, verifies and publishes
  both packages together — replacing the earlier split of `release.yml` /
  `release-extensions.yml` with independent tag prefixes.

## [0.2.0] - 2026-07-31

### Added
- **Core:** `Stream`-based overloads across the public API (`HtmlToDocxConverter`,
  `DocxToPdfConverter`, `HtmlToPdfConverter`, `DocxEditor`, `WorkbookEditor`,
  `PresentationEditor`), alongside the existing `byte[]` API, for large-document efficiency.
  Additive, non-breaking.

## [0.1.0] - 2026-07-29

### Added
- Initial release of `Ank.DocToolkit`: HTML → DOCX, HTML → PDF (via a DOCX pivot — no
  permissively-licensed, NuGet-only, Linux-safe library renders HTML to PDF directly), DOCX
  open/edit (placeholder replacement that preserves per-run formatting and hyperlinks), XLSX
  create/read/edit (ClosedXML), PPTX open/edit (raw `DocumentFormat.OpenXml`).
- Offline guarantee: no default code path performs network I/O, enforced by air-gap guard tests
  against markup naming a loopback listener sixteen different ways.
- CI: Linux + Windows build/test, native-binary guard, banned-package guard, `SixLabors.Fonts`
  pin guard — re-proving the package's licensing/portability premise on every push.
- Tag-driven release to nuget.org via Trusted Publishing (OIDC) — no stored API key.

[0.2.1]: https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.2.0...v0.2.1
[0.2.0]: https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.1.0...v0.2.0
[0.1.0]: https://github.com/Ank-KhoaHo/DocToolkit/releases/tag/v0.1.0
