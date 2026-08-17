# Changelog

Fixture: the same 0.30.0 entry after curation, trimmed to the parts that matter to the
check. The load-bearing difference is the Changed section - a behaviour change that
release-please files under Added, because it reads the commit TYPE and #290 was a feat.

## [0.30.0](https://github.com/Ank-KhoaHo/DocToolkit/compare/v0.29.0...v0.30.0) (2026-08-17)


### Added

* **Core: legacy PowerPoint `.ppt` renders to PDF, and is now claimed rather than accidental.**
  Measured across 15 real `.ppt` files: **11 convert (73%)**, the slowest in 1.7 s
  ([#290](https://github.com/Ank-KhoaHo/DocToolkit/issues/290)).


### Changed

* **Core: `XlsxToPdfConverter` now REFUSES a legacy `.xls` immediately, where it used to attempt
  the render.** This is the one behaviour change in this release that can break working code.

  **Migrating:** if you were passing `.xls` bytes to `XlsxToPdfConverter`, you now get a
  `DocumentConversionException` at once, naming the format and telling you to save as `.xlsx`.
