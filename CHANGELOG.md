# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/).

Ank.DocToolkit and Ank.DocToolkit.Extensions.DependencyInjection release together, at the same
version, from a single tag (see README.md > Releasing). Entries below are prefixed **Core:** or
**Extensions:** when they apply to only one package; unprefixed entries apply to both or to
repo-wide tooling (CI, release pipeline).

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
