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
