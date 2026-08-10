# Roadmap

**No dates.** This is a small project, and a roadmap with dates on it would be a work of fiction.
What follows is direction and, more usefully, the things that are **not** coming and why — so you
can decide whether this package fits before taking the dependency, rather than after.

Anything here may change. Nothing here is a commitment.

## Where it is

All four constraints hold and are re-checked by CI on every push: permissive licences only, NuGet
only with no native binaries, runs on Linux (x64 and arm64), Windows and macOS, and no runtime
network I/O by default.

Current capabilities: HTML → DOCX and PDF; DOCX → PDF, HTML and Markdown; XLSX → PDF; PPTX → PDF;
create and edit DOCX, XLSX and PPTX; template filling with repeating rows and image placeholders;
page setup with headers and footers; reading an existing PDF — page count, merge, page extraction
and metadata; a DI package mirroring the whole surface.

Around it: nine runnable samples, a docs site with six conceptual guides whose code blocks are
mostly compiled as part of a sample — the handful that are not are marked in place — and a
per-release attested CycloneDX SBOM alongside build provenance.

## Under consideration

Roughly in order of how often the gap has actually been hit. None is scheduled.

- **Surfacing conversion warnings.** The XLSX and PPTX renderers report features they could not
  represent; today those are dropped silently and the limitation is documented instead. A warning
  channel is purely additive, so it can be added when somebody needs it rather than guessed at now.

## Not planned, and why

This section is the useful one.

| | |
|---|---|
| **1.0.0** | Never. `0.x` forever, enforced in configuration rather than intended. Under `0.x` semver already says anything may change, which is an honest description of this package. |
| **`net9.0`** | Adds zero reach: a `net9.0` app already consumes the `net8.0` build. It would cost a matrix leg and is the only STS target on offer against two LTS. |
| **`netstandard2.0`** | Not blocked by dependencies — all nine support it. Blocked because the bounded-fetch guarantee on remote images **cannot be expressed** there (no cancellable DNS or stream read), and `DateOnly`/`TimeOnly` would make the public API differ per target. A security guarantee that holds on one target and not another is worse than not offering the target. |
| **Native AOT compatibility** | Not claimed until CI both compiles *and* runs an AOT build. An unverified compatibility claim is worse than an absent one. |
| **An input size limit** | This library edits and converts documents; refusing a large one is a defect, not a safeguard. The memory profile is documented instead so you can size a host. |
| **Keyed DI registrations** | Permanent registration surface for a multi-tenant scenario nobody has asked for. Revisit if someone does. |
| **Anything needing a browser, LibreOffice, Office interop, or a native binary** | Out of scope by construction — it is the reason this package exists. |

## What moves an item

Someone hitting the gap and saying so. This project has repeatedly found that a feature assumed hard
was not, once somebody measured the dependency graph instead of reasoning about it — so a request
that names the task you are stuck on is worth more than it might seem.

Open a [feature request](https://github.com/Ank-KhoaHo/DocToolkit/issues/new/choose).
