# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/).

Ank.DocToolkit and Ank.DocToolkit.Extensions.DependencyInjection release together, at the same
version, from a single tag (see README.md > Releasing). Entries below are prefixed **Core:** or
**Extensions:** when they apply to only one package; unprefixed entries apply to both or to
repo-wide tooling (CI, release pipeline).

## [0.12.0](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.11.0...v0.12.0) (2026-08-08)


### Added

* **core:** add XlsxSheet, the typed model for a named sheet ([8af39c8](https://github.com/Ank-KhoaHo/DocToolkit/commit/8af39c8e3a68c9a945e1c9f05d9dc8645ee92543))
* **core:** append rows to an existing sheet ([5d8b2fe](https://github.com/Ank-KhoaHo/DocToolkit/commit/5d8b2fe0a8db19fd531344a0851668866559e9b4))
* **core:** create a workbook with more than one sheet ([f426aaa](https://github.com/Ank-KhoaHo/DocToolkit/commit/f426aaa4b6b09bbb85fb296edb64b45be039638b))
* **core:** let any cell hold a formula via XlsxFormula ([49982ad](https://github.com/Ank-KhoaHo/DocToolkit/commit/49982ad14f29be6ce845845e29d1d20ea496bf19))
* **extensions:** mirror PresentationEditor.Create on IPresentationEditor ([7578053](https://github.com/Ank-KhoaHo/DocToolkit/commit/7578053f9509c85930c4f3bc98860d013a310936))
* mark both packages IsTrimmable, and guard the claim in CI ([83f182f](https://github.com/Ank-KhoaHo/DocToolkit/commit/83f182f593a0ed382600fe0d45b9f8177e9bfc5a))


### Fixed

* claim trimmability by attribute, not by the IsTrimmable property ([46adf28](https://github.com/Ank-KhoaHo/DocToolkit/commit/46adf28c7962a412dddaabfdd82a89fde46137c2))
* **core:** make created slides real title and body placeholders ([a8dcc36](https://github.com/Ank-KhoaHo/DocToolkit/commit/a8dcc3646fc677e585849b546af4d4abc89b1d52))
* **core:** reject an invalid sheet name with ArgumentException, not a wrapped failure ([aa806aa](https://github.com/Ank-KhoaHo/DocToolkit/commit/aa806aa29109c32225ea5219cdf9351fa24bf8f0))
* give the Linux container image the .NET 8 runtime it needs ([fb8c00e](https://github.com/Ank-KhoaHo/DocToolkit/commit/fb8c00e89f523d622a5116441336f3687e68b2b8))
* **samples:** use Path.Join rather than Path.Combine for the output path ([929b1b7](https://github.com/Ank-KhoaHo/DocToolkit/commit/929b1b702d89996385d1abed08fc2b9410de396e))


### Changed

* pin the SDK to 10.0.302 in global.json ([f8c8e6d](https://github.com/Ank-KhoaHo/DocToolkit/commit/f8c8e6d83e0874645a439e942621c4b4c9c7bf78))
* share MSBuild properties and hold the warning line locally ([52e4c9a](https://github.com/Ank-KhoaHo/DocToolkit/commit/52e4c9a67ecdfcb03fcce77721d660c8f1170a4d))
* update OfficeIMO.Word.Pdf to 3.2.0 ([3061462](https://github.com/Ank-KhoaHo/DocToolkit/commit/3061462d4c137921ea5565ae9e41266fbed1838e))

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
