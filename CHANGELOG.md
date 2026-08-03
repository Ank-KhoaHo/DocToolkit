# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/).

Ank.DocToolkit and Ank.DocToolkit.Extensions.DependencyInjection release together, at the same
version, from a single tag (see README.md > Releasing). Entries below are prefixed **Core:** or
**Extensions:** when they apply to only one package; unprefixed entries apply to both or to
repo-wide tooling (CI, release pipeline).

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
