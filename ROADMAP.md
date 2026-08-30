# Roadmap

**No dates.** This is a small project, and a roadmap with dates on it would be a work of fiction.
What follows is direction and, more usefully, the things that are **not** coming and why — so you
can decide whether this package fits before taking the dependency, rather than after.

Anything here may change. Nothing here is a commitment.

## Where it is

All four constraints hold and are re-checked by CI on every push: permissive licences only, NuGet
only with no native binaries, runs on Linux (x64 and arm64), Windows and macOS, and no runtime
network I/O by default.

Current capabilities: HTML → DOCX and PDF; **Markdown → DOCX and PDF**; DOCX → PDF, HTML and
Markdown; XLSX → PDF, **CSV and HTML**; PPTX → PDF; **legacy Word 97-2003 `.doc` → DOCX and legacy
PowerPoint 97-2003 `.ppt` → PDF**, both at published rates rather than as a tick; create and edit
DOCX, XLSX and PPTX; **supplying your own fonts** for characters the PDF renderer cannot otherwise
encode, which takes the host machine out of the answer;
**sheet formatting** — bold header, frozen header, auto-fit, number formats; template filling with
repeating rows and image placeholders; page setup with headers and footers; reading an existing
PDF — page count, **text extraction**, merge, page extraction and metadata; **conversion loss
reporting** on the DOCX text exporters and the Markdown importers; a DI package mirroring the whole
surface, which is now enforced by a derived check rather than asserted.

Both packages are trim-safe and the core package is **native-AOT compatible**, each proved by
publishing a probe application and running it.

Around it: twelve runnable samples, a docs site with eight conceptual guides whose code blocks are
mostly compiled as part of a sample — the handful that are not are marked in place — and a
per-release attested CycloneDX SBOM alongside build provenance. One of those guides is a
**capability matrix generated from the approved API file**, so the list of what this converts can
no longer drift from what it ships; CI fails when it does.

## Under consideration

Roughly in order of how often the gap has actually been hit. None is scheduled.

- **Surfacing conversion warnings from the PDF renderers.** Half of this shipped in 0.26.0: the
  DOCX → HTML/Markdown exporters and the Markdown importers now offer `ConvertWithReport`, which
  returns the output plus what the conversion could not carry across. That half turned out not to
  need designing at all — the underlying library was already computing a structured loss report on
  every call and this package was discarding it.

  The other half is genuinely absent and cannot be done the same way: **the PDF renderers produce
  no report to surface.** Features they cannot represent — charts, conditional formatting, some
  shape effects — are still dropped silently, and the limitation is documented instead.

  **Two of those silent drops are now measured rather than assumed**, and both lose content rather
  than styling: **an unstyled footnote reference loses its text** — the renderer keys on the
  character style Word and `DocxEditor.AddFootnote` both apply to a footnote reference, so a
  footnote authored either way survives; only a reference missing that style is lost — and **a
  table nested inside a table cell loses its content entirely**. Measured 2026-08-25, reading the
  PDF back, with a sibling paragraph in each fixture so a missing token could not be mistaken for an
  empty render. Content controls and text boxes were measured to *survive*, so neither is on that
  list.

  The loss is upstream: `DocxToPdfConverter` is a pass-through to the renderer, so this package
  cannot fix it — only report it and say so.

  **Reporting it is now an API rather than a paragraph.** `DocxToPdfPreflight.Inspect` lists what a
  document CONTAINS that this renderer may not represent, so a caller converting third-party files
  can route those for human review. It deliberately does not claim to know what was dropped — that
  claim is the one this section refuses, and an inventory of the input does not need it. Offering a
  `ConvertWithReport` there that always returned an empty list would be a documented lie.

## Not built, but cheap — ask if you need one

Separate from the list above, which is ordered by how often a gap has been hit. **Nothing here has
been asked for even once.** It is listed because this page exists to help you decide whether the
package fits, and "we could, nobody has" is more useful to you than silence.

Surveyed on 2026-08-15 against what the commercial document libraries sell, then **measured by
running it**: several of those capabilities are already reachable from the dependency graph this
package resolves today, at **no new dependency and therefore no change to the four constraints**.
None is exposed, because the standing rule is that a gap ships when somebody hits it.

The ones that were measured working: **rendering a DOCX, XLSX or PPTX page to PNG/JPEG/SVG**
(thumbnails and previews — no browser involved), **PDF → DOCX**, **PDF redaction**, **PDF
encryption with permissions**, **PDF watermarking**, and **opening or writing password-protected
XLSX and PPTX**.

Two honest caveats, because an unqualified list here would be the sales pitch this page is not.
The page renderer resolves **direct formatting but not paragraph styles** — a `Heading1` renders
like body text — so it suits a recognisable thumbnail, not a faithful page image. And all of it was
measured on Windows only; nothing is proven on Linux or macOS, which is the bar every shipped claim
here has to clear.

If one of these is the thing you are stuck on, say so in a
[feature request](https://github.com/Ank-KhoaHo/DocToolkit/issues/new/choose) — that is what moves
it, and the measurement is already done.

## Not planned, and why

This section is the useful one.

| | |
|---|---|
| **1.0.0** | Never. `0.x` forever, enforced in configuration rather than intended. Under `0.x` semver already says anything may change, which is an honest description of this package. |
| **`net9.0`** | Adds zero reach: a `net9.0` app already consumes the `net8.0` build. It would cost a matrix leg and is the only STS target on offer against two LTS. |
| **`netstandard2.0`** | Not blocked by dependencies — all nine support it. Blocked because the bounded-fetch guarantee on remote images **cannot be expressed** there (no cancellable DNS or stream read), and `DateOnly`/`TimeOnly` would make the public API differ per target. A security guarantee that holds on one target and not another is worse than not offering the target. |
| ~~**Native AOT compatibility**~~ | **Shipped in 0.27.0**, and it sat in this table until it was earned rather than being promised from it. The bar this row set — "not claimed until CI both compiles *and* runs an AOT build" — is exactly the bar that was met: a job now native-AOT-publishes a probe over the real dependency closure and runs it, asserting every capability's result. Left here rather than deleted, because a row that moved out of "not planned" by meeting its own stated condition is the most useful kind of entry this table has. |
| **An input size limit** | This library edits and converts documents; refusing a large one is a defect, not a safeguard. The memory profile is documented instead so you can size a host. |
| **Keyed DI registrations** | Permanent registration surface for a multi-tenant scenario nobody has asked for. Revisit if someone does. |
| **Anything needing a browser, LibreOffice, Office interop, or a native binary** | Out of scope by construction — it is the reason this package exists. |

## What moves an item

Someone hitting the gap and saying so. This project has repeatedly found that a feature assumed hard
was not, once somebody measured the dependency graph instead of reasoning about it — so a request
that names the task you are stuck on is worth more than it might seem.

Open a [feature request](https://github.com/Ank-KhoaHo/DocToolkit/issues/new/choose).
