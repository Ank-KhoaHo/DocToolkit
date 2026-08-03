# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/).

Ank.DocToolkit and Ank.DocToolkit.Extensions.DependencyInjection release together, at the same
version, from a single tag (see README.md > Releasing). Entries below are prefixed **Core:** or
**Extensions:** when they apply to only one package; unprefixed entries apply to both or to
repo-wide tooling (CI, release pipeline).

## Unreleased

### Added
- Runnable sample projects (`samples/ConsoleSample`, `samples/MinimalApiSample`), referencing
  the published packages and built by the existing CI alongside the libraries.
- A DocFX-generated API-reference site (`docfx/`), published to GitHub Pages on every
  successful release via `.github/workflows/docs.yml`.

## [0.3.0](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.2.2...v0.3.0) (2026-08-03)


### Added

* **di-extensions:** add Stream overload to IDocxToPdfConverter ([1bd712c](https://github.com/Ank-KhoaHo/DocToolkit/commit/1bd712c62e1cf4817a4fb6273bfed58f64b36c1e))
* **di-extensions:** add Stream overload to IHtmlToDocxConverter ([f136c97](https://github.com/Ank-KhoaHo/DocToolkit/commit/f136c97f62ab888c0af92b61c63cbc4f1e6d08cf))
* **di-extensions:** add Stream overload to IHtmlToPdfConverter ([4f131b2](https://github.com/Ank-KhoaHo/DocToolkit/commit/4f131b2dadb37e609529dc82f96c8d57fec92fb6))
* **di-extensions:** add Stream/async overloads to IDocxEditor ([4be6e1d](https://github.com/Ank-KhoaHo/DocToolkit/commit/4be6e1d65203543c2a0d6e9738fff9ec2a325e47))
* **di-extensions:** add Stream/async overloads to IPresentationEditor ([c95f958](https://github.com/Ank-KhoaHo/DocToolkit/commit/c95f95866be0cd99593bd4506ffeab1c8bee1ca6))
* **di-extensions:** add Stream/async overloads to IWorkbookEditor ([0592b74](https://github.com/Ank-KhoaHo/DocToolkit/commit/0592b74cbf0872648eee8f54b11972f63f198cae))
* **extensions:** add Stream-based async members to all six DI interfaces ([21890a2](https://github.com/Ank-KhoaHo/DocToolkit/commit/21890a2b9041863ada07a9ac44deba9e98b06c4c))


### Fixed

* **di-extensions:** bump Ank.DocToolkit version floor to 0.2.0 ([51646e2](https://github.com/Ank-KhoaHo/DocToolkit/commit/51646e2c17fecffcb7ae6e410d25a70d11180c58))


### Changed

* add implementation plan for DI extensions Stream/async parity ([bb93272](https://github.com/Ank-KhoaHo/DocToolkit/commit/bb9327265ad72ccac8fdeffaf2292a28c74be442))
* add the API documentation link now that the site is verified live ([812e433](https://github.com/Ank-KhoaHo/DocToolkit/commit/812e4339dd06776ed5eb773e80ec792bdcbbd32a))
* correct version-floor assumption found during Task 1 ([5dd9a26](https://github.com/Ank-KhoaHo/DocToolkit/commit/5dd9a2636526415c3de741878dfe0b5443127f65))
* design Stream/async parity for DI extensions ([a80cd21](https://github.com/Ank-KhoaHo/DocToolkit/commit/a80cd210b1e48ddbcbe7cf8798d09f9e78a5a1d3))
* **di-extensions:** document the Stream overloads ([6a8710c](https://github.com/Ank-KhoaHo/DocToolkit/commit/6a8710cfab22ef418a8fc89e18b724c886cba8a0))
* **di-extensions:** state what the interfaces do not mirror ([a570e2d](https://github.com/Ank-KhoaHo/DocToolkit/commit/a570e2de98c69d3f887d5895b2e85060e67156dc))
* **extensions:** correct stale counts and tighten README claims ([5ed5596](https://github.com/Ank-KhoaHo/DocToolkit/commit/5ed55969fcf03bd092864edb33f5b74eb68f7d1a))
* **extensions:** warn that streamed output is partial on failure ([43ff9ac](https://github.com/Ank-KhoaHo/DocToolkit/commit/43ff9acab2722b09d45f3d8f15eaeda97b329c97))
* fix stale version-floor claim in the plan's Tech Stack line ([1e7f0f2](https://github.com/Ank-KhoaHo/DocToolkit/commit/1e7f0f214b7b65a7d468d9370f663d7072657a54))
* narrow the determinism claim for ClosedXML edits ([5765939](https://github.com/Ank-KhoaHo/DocToolkit/commit/57659394332a9c48edfdf7c7c8252d7a5d5fbf68))
* require a Stream-path option guard in Task 5 ([e5912a1](https://github.com/Ank-KhoaHo/DocToolkit/commit/e5912a19f8049b5bfd1bb67e0d9a2b776857cb04))

## [0.2.2](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.2.1...v0.2.2) (2026-08-03)


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
