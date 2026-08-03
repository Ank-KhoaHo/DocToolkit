# CHANGELOG.md — design

## Why

`Ank.DocToolkit` and `Ank.DocToolkit.Extensions.DependencyInjection` have shipped three releases
(`0.1.0`, `0.2.0`, `0.2.1`) with no changelog. Adopters currently have to read commit history or
GitHub's auto-generated release notes (from PR titles) to know what changed. A maintained
changelog is a standard expectation for a public NuGet package and, per the repo's own guard
philosophy (native-binary/banned-package/SixLabors guards already block a bad release), a missing
entry should block a release the same way a broken premise does.

## Scope

One file: `CHANGELOG.md` at the repo root, covering both packages. Not in scope: automated
changelog generation from commit messages (that's the separate, later "semantic-release
automation" brainstorm) — this is a hand-maintained file with CI enforcement that an entry exists,
not that it's auto-written.

## Format

[Keep a Changelog](https://keepachangelog.com/en/1.1.0/) conventions:

```markdown
# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/).

Ank.DocToolkit and Ank.DocToolkit.Extensions.DependencyInjection release together, at the same
version, from a single tag (see README.md > Releasing). Entries below are prefixed **Core:** or
**Extensions:** when they apply to only one package; unprefixed entries apply to both or to
repo-wide tooling (CI, release pipeline).

## [Unreleased]

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

[Unreleased]: https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.2.1...HEAD
[0.2.1]: https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.2.0...v0.2.1
[0.2.0]: https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.1.0...v0.2.0
[0.1.0]: https://github.com/Ank-KhoaHo/DocToolkit/releases/tag/v0.1.0
```

Note on `0.2.0`: by the time that tag was pushed, the extensions package's source already existed
on `main` (merged via PR #1), but `release.yml` at that point only packed and published the core
project — the extensions package's actual first publish was `0.2.1`. The changelog reflects what
was *published* at each version, not what was in the tree, which is why the extensions package
doesn't appear until `0.2.1`.

## The CI guard

A new step in `.github/workflows/release.yml`, immediately after `Resolve version` and before
`Build` (fail fast, before any expensive work runs):

```yaml
- name: Guard - CHANGELOG.md has an entry for this version
  run: |
    if ! grep -Eq "^## \[${{ steps.v.outputs.version }}\]" CHANGELOG.md; then
      echo "::error::Refusing to publish - CHANGELOG.md has no '## [${{ steps.v.outputs.version }}]' entry. Add one before tagging."
      exit 1
    fi
    echo "CHANGELOG.md has an entry for ${{ steps.v.outputs.version }}"
```

Plain bash grep for the heading, matching the style of the existing guards (no native binaries,
no banned packages, `SixLabors.Fonts` pin) — simple, fast, and consistent with how this workflow
already fails releases that break a premise. Does not validate the entry's *content*, only that
the heading exists; over-validating the prose isn't worth the complexity for a solo-maintainer
changelog.

## Docs updates

- `CLAUDE.md`'s Releasing section gets one line: updating `CHANGELOG.md` (moving `[Unreleased]`
  content under a new version heading) is part of preparing a release, before tagging.
- `README.md` gets a link to `CHANGELOG.md` (e.g. from the badges line or a "Changelog" mention
  near the top), so it's discoverable from the package landing page on nuget.org too (nuget.org
  renders `README.md` but a reader following the GitHub link from there will find it easily).

## Testing / validation

No unit tests apply (this is a markdown file plus a CI grep step). Validation is: run the guard
step logic locally against the backfilled file for each of the three historical versions and
confirm it matches; then prove it fails correctly by grepping for a version number that is
deliberately absent (e.g. `9.9.9`) and confirming the step would exit 1.
