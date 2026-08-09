# 2. Reject any dependency that ships native binaries

**Status:** accepted

## Context

PPTX support originally used **ShapeCrawler**. Its API was checked carefully — by reflection, to get
the method names right. Its *dependencies* were never checked. Those are different questions, and
the difference cost a rewrite.

ShapeCrawler pulls in SkiaSharp and Magick.NET. Measured in build output:

- **38 native `.so`/`.dylib` files**
- **664 MB** of `runtimes/`
- **26 CVE advisories**

Re-measured 2026-08-09 against a current version: **19 native files and 664 MB** — the payload has
not gone away.

`System.Drawing.Common` is the nastier shape of the same problem: it restores and builds fine, then
throws `PlatformNotSupportedException` at runtime on anything that is not Windows. Nothing fails
until production.

## Decision

No dependency that ships native binaries, and PPTX was rewritten against raw
`DocumentFormat.OpenXml`.

A named list — `EPPlus`, `NPOI`, `Spire.*`, `Syncfusion.*`, `QuestPDF`, `IronPDF`, `ShapeCrawler`,
`SkiaSharp`, `Magick.NET*`, `System.Drawing.Common` — fails the build if it appears in the resolved
graph, for licensing or native-payload reasons.

## Consequences

Some things are simply harder. Image dimensions are read from PNG and JPEG headers by hand rather
than by an imaging library, because the obvious choice (SixLabors.ImageSharp) moved its later majors
onto a revenue-gated licence.

## What would change this

Nothing, for the native-binary half — it is constraint 2 and the package's reason to exist.

The **named list** is a different matter: it is a convenience so a doomed pull request fails fast
with a clear message. If one of those packages changes its licence or drops its native payload, the
entry should be removed after re-measuring. The guard that must never be relaxed is the general one
that scans the resolved graph, not the list of names.
