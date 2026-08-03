# DocToolkit — guidance for Claude Code

A .NET library that converts **HTML → DOCX/PDF** and opens/edits **DOCX, XLSX, PPTX**, shipped
as the `Ank.DocToolkit` NuGet package.

Read `README.md` first for what it does. This file covers what will bite you while changing it.

## The design premise — do not break these

The package exists *only* because it satisfies four constraints at once. Any change that breaks
one makes the package pointless, so all four are enforced by tests and CI:

1. **Permissive licences only** — MIT / Apache-2.0 / BSD. No revenue thresholds, no per-seat fees.
2. **NuGet only** — no browser download, no LibreOffice, **no native binaries**.
3. **Runs on Linux** — verified on `ubuntu-24.04` in CI, not assumed.
4. **No runtime network I/O** — consumers deploy to air-gapped machines with NuGet access only.

All four are properties of the *resolved dependency graph*, not of this code. A single `dotnet add
package` can break every one of them silently. That has already happened once — see below.

### Never add these packages

`EPPlus` · `NPOI` · `Spire.*` · `Syncfusion.*` · `QuestPDF` · `IronPDF` — **not free for
commercial use** (Polyform Noncommercial, paid maintenance fees, revenue-gated community licences,
or outright commercial).

`ShapeCrawler` · `SkiaSharp` · `Magick.NET*` · `System.Drawing.Common` — **drag in native binaries
or break on Linux.** `System.Drawing.Common` is the nastiest: it restores and builds fine, then
throws `PlatformNotSupportedException` at runtime on non-Windows.

`DependencyGuardTests` fails the build if any of these appear. **If that test goes red, remove the
package — never relax the test.**

### Why ShapeCrawler is on that list

PPTX support originally used it. It turned out to depend on SkiaSharp and Magick.NET, which put
**38 native `.so`/`.dylib` files and 664 MB of `runtimes/`** into build output, plus 26 CVE
advisories. It was replaced with raw `DocumentFormat.OpenXml`.

The mistake that let it in: its **API** was checked (by reflection, to get the method names right)
but its **dependencies** never were. Those are different questions. Before adding anything:

```bash
dotnet list package --include-transitive
find . -path '*/bin/*' \( -name '*.so' -o -name '*.so.*' -o -name '*.dylib' \)
```

### `SixLabors.Fonts` is pinned to `[1.0.0]` on purpose

`OfficeIMO.Word` requests `[1.0.0, 3.0.0)`. Version 2.x switches to the **Six Labors Split
License** — Apache-2.0 only under $1M annual revenue, commercial above it. Floating to 2.x would
silently move this package off permissive licensing. CI asserts the pin holds. Do not unpin it.

## Offline guarantee

No default code path may open a socket. `AirGapGuardTests` asserts **zero** connections across the
whole public API, against markup naming a loopback listener sixteen ways (`<img src>`, `srcset`,
`<link rel=stylesheet>`, `@import`, `background-image`, table-cell images, `<iframe>`, `<object>`,
`<script>`).

`HtmlToDocxConverter` enforces this with an `IWebRequest` implementation that refuses everything,
so the fetching component is never constructed. That is deliberately stronger than setting
`ImageProcessingMode` — a rendering-policy knob could be reinterpreted by a future release. It also
blocks `file://` reads, which the default requester would otherwise serve.

The single opt-in (`allowRemoteImageDownload: true`) is documented as failing in air-gapped
environments. **Do not make network access the default, and do not weaken these tests.**

Related: HtmlToOpenXml 3.5.0 has a **non-thread-safe process-wide static `HttpClient`** that
crashes under parallel use. The no-network default avoids that path entirely. Don't go near it.

## Traps in this codebase

**PDF assertions must go through `PdfProbe`.** OfficeIMO writes content streams **uncompressed**
and emits text as **hex-string operators** — `<41636D65> Tj` is `"Acme"`. So searching the raw PDF
bytes for `"Acme"` finds nothing, and inflating the streams finds nothing either. Both fail
*silently* and look exactly like a broken converter. `PdfProbe` decodes correctly, including the
WinAnsi range 0x80–0x9F where `0x97` is an em-dash, not a control character.

**PDFs must stay `binary` in `.gitattributes`.** Because those streams are uncompressed, the files
contain few NUL bytes and git's binary auto-detection guesses "text". It then applies LF→CRLF on
Windows checkout and injects stray carriage returns — `result.pdf` once gained 1,743 bytes that
way, shifting every xref offset. The `.docx`/`.pptx` fixtures are unaffected (ZIPs contain NULs).

**Word and PowerPoint split words across runs.** A single visible `{{placeholder}}` is often
several `w:t` / `a:t` elements. Naive per-run `string.Replace` misses them. `RunTextSplicer` maps
match offsets back onto individual runs so only the runs a match actually overlaps get written —
that is what preserves per-run formatting and leaves `w:hyperlink` children intact. Don't
"simplify" it into a merge-everything-onto-run-0 loop; that silently flattens formatting and guts
hyperlinks.

**`DocxEditor` must not reach into nested paragraphs.** `body.Descendants<Paragraph>()` also yields
paragraphs inside `w:txbxContent` (text boxes). Reaching into them once caused text-box content to
be deleted and relocated into the outer paragraph — schema-valid, no exception, silent data loss.

**HTML → PDF pivots through DOCX by design.** No permissively-licensed, NuGet-only, Linux-safe
library renders HTML to PDF directly: the only free renderers *are* browsers, and a browser is a
native binary. `HtmlToPdfConverter` composes the other two converters — keep it a composition, do
not reimplement conversion inside it.

**Every `Stream` overload follows one shape.** Parameters are `Stream source` wherever the
`byte[]` overload took bytes, then `Stream destination`, then `CancellationToken ct = default`.
`StreamPipeline.RequireReadable`/`RequireWritable` guard by name before anything runs;
`StreamPipeline.DrainAsync` reads a source into a scratch `MemoryStream` (this is where the async
in these overloads is earned — a source may be an HTTP request body, forward-only and
non-seekable); `StreamPipeline.EmitAsync` copies a finished buffer out. **Caller-owned streams are
never disposed, closed or sought.** `DocxToPdfConverter.ConvertAsync` is the one exception that
writes straight through as OfficeIMO's own writer produces the PDF, rather than buffering it —
that is deliberate (see the class's doc comment), not an inconsistency to "fix". Every `*Core`
method (`ReplaceTextCore`, `ExtractTextCore`, `SetCellCore`, …) holds the one real implementation
that both the `byte[]` and `Stream` overloads call, so the two can never drift apart.

`StreamOverloadTests` proves these properties with purpose-built stream doubles, not
`MemoryStream`: `ForwardOnlySource`/`ForwardOnlySink` throw on `Seek`/`Length`/`Position`, so an
implementation that rewinds a destination or reads back what it wrote fails there instead of
against a real socket in production. `TrackingStream` counts sync vs. async read/write calls, so a
`byte[]` round-trip wearing a `Stream` signature gets caught rather than passing by accident. If
you add a new `Stream` overload, add it to the name lists at the top of that file — an overload
missing from those lists is the only way to escape the whole suite.

## Conventions

- **Target frameworks are `net8.0;net10.0`.** Every test runs once per framework, so *N* tests
  report *2N* results. 205 tests (182 core + 23 extensions) → 410 results.
- **Never replace `src/DocToolkit/DocToolkit.csproj` wholesale** — it carries the package metadata
  (`PackageId`, version, licence expression, readme, symbol package). Use `dotnet add package`,
  which edits in place.
- Public API is **static classes**, stateless and safe to call concurrently, with a `byte[]`
  overload and a `Stream` overload for every capability (see above). Failures are wrapped in
  `DocumentConversionException`. Adding overloads is fine; changing existing names or signatures is
  a breaking change for consumers.
- **Async where there is real I/O, sync where the work is CPU-bound.** A `Stream` overload is
  async because draining/emitting a stream genuinely awaits; the document-processing logic inside
  it (OpenXml, ClosedXML) is not wrapped in `Task.Run` to look async when it isn't.
- **Commit messages must not contain a `Co-Authored-By` trailer.**
- The build runs with `-warnaserror` and currently has **0 warnings**. Keep it there.

## The DI extensions package

`src/DocToolkit.Extensions.DependencyInjection/` ships as its own NuGet package,
`Ank.DocToolkit.Extensions.DependencyInjection`, but is **versioned and released together with the
core package** — one `v*` tag packs and publishes both at the same version, in the same
`.github/workflows/release.yml` run (see that file's header comment). This is deliberate: the two
packages are meant to stay in lockstep, and nuget.org Trusted Publishing policies are keyed to an
exact workflow filename, so one workflow file also means only one policy to ever configure.
An earlier design split this into `release.yml`/`release-extensions.yml` with independent
`v*`/`ext-v*` tags; that let the two packages' versions drift apart with no way to enforce they
matched, and required a second Trusted Publishing policy. Don't reintroduce that split.

It references `Ank.DocToolkit` as a real `PackageReference`, never a `ProjectReference` — the
whole point is to prove the extensions package works the way an external consumer's restore
would, against the *published* core package, not against whatever is currently on `main`. Before
changing an interface here, confirm the byte[] method it wraps actually exists in the core
version this project's `Ank.DocToolkit` reference floor requires.

Six interfaces mirror the six static classes 1:1 (`byte[]` in, `byte[]`/`string`/`int` out — no
`Stream` overloads here; that was a deliberate scope decision, not an oversight, since the DI
layer was designed before the static API's `Stream` overloads existed). Service implementations
are `internal sealed` — never `public` — and are pure delegation, one line per method, to the
matching static method. If a service method does anything more than call through, that logic
belongs in the core static method instead.

`DocToolkitOptions.AllowRemoteImageDownload` replaces the static API's per-call
`allowRemoteImageDownload` bool: configured once at `AddDocToolkit(configure)`, not re-decided per
call. `ServiceCollectionExtensionsTests` proves the wiring with a small self-contained loopback
listener (not a copy of the core project's `AirGapGuardTests` — that already proves the
*conversion* behaviour exhaustively; this only has to prove the option value reaches the static
method's parameter).

## Samples and docs site

`samples/ConsoleSample` and `samples/MinimalApiSample` are runnable, added to `DocToolkit.sln`,
and reference the published packages via `PackageReference` (never `ProjectReference`) — same
reasoning as the extensions package itself: they prove the real published artifact works, not
whatever is currently on `main`. They're built by the existing CI `dotnet build` step with no
special handling; a breaking API change fails the next sample build.

`docfx/` holds a DocFX-generated API-reference site — separate from `docs/`, which holds this
project's planning/spec history, not site source. `.github/workflows/docs.yml` builds and deploys
it to GitHub Pages, triggered by `workflow_run` on `release.yml` completing **successfully** — not
independently on the same tag push, so a release that fails its guards never gets a docs site
describing it as shipped. Don't "simplify" this into a direct tag-push trigger; that would break
the guarantee.

**`docfx.json`'s `globalMetadata` must keep `_enableSearch: false` and `pdf: false`.** Without
them, DocFX's default template downloads a ~109 MB headless-Chromium binary via Playwright/Node.js
during the build — verified while adding this pipeline. Re-enabling either without re-verifying
the browser download doesn't come back will make `docs.yml` slow and pull in exactly the kind of
heavyweight native dependency this repo's premise guards otherwise keep out.

`ConsoleSample` reaches into `tests/DocToolkit.Tests/assets/sample.pptx` for its PPTX demo —
there's no "create a PPTX from scratch" method in the public API, so this is a deliberate,
brief-sanctioned trade-off. If that fixture ever moves, the sample fails with an opaque MSBuild
copy error, not a message pointing at the real cause.

## Commands

```bash
dotnet build DocToolkit.sln -c Release -warnaserror
dotnet test  DocToolkit.sln -c Release            # 205 tests x 2 TFMs = 410 results
dotnet pack  src/DocToolkit/DocToolkit.csproj -c Release
dotnet pack  src/DocToolkit.Extensions.DependencyInjection -c Release

# Linux, the way CI checks it
docker build -f Dockerfile.linux-test -t doctoolkit-linux-test .
docker run --rm doctoolkit-linux-test
```

## Branches

Two branches, two jobs.

- **`develop` is the trunk.** All work merges here. It carries everything, including
  `CLAUDE.md`, `docs/` and `spike/`.
- **`main` is release-only.** It carries the shipping library and nothing about the process that
  produced it: `CLAUDE.md`, `docs/` and `spike/` are stripped. It is what release-please watches
  and what `release.yml` tags, packs and publishes. It is also the GitHub default branch, so it
  is the tree a consumer arriving from nuget.org lands on — that is the entire point.

Feature branches (`feat/**`, `fix/**`) branch from `develop` and PR back into `develop`.
`ci.yml`'s `branch-policy` job rejects any PR into `main` that is not a `release/promote-*`
branch or release-please's own Release PR, so a mis-targeted PR fails rather than merging quietly.

**Promote with `scripts/promote-to-main.sh`.** It merges `develop` into a new
`release/promote-*` branch based on `main`, purges the excluded paths, and opens a PR. Promotion
is a *real* `git merge`, not a generated snapshot, so every Conventional Commit subject from
`develop` stays reachable from `main` — that is what keeps release-please working unchanged.
**Merge that PR with a merge commit, never a squash**; squashing collapses those subjects and
release-please would compute the wrong bump, or none.

`scripts/` is deliberately *not* excluded from `main`: it is release tooling, and stripping it
would let the promote script delete itself out from under the running bash process.

### Never merge `main` into `develop`

`main` carries *deletions* of `CLAUDE.md`, `docs/` and `spike/`. A `git merge main` on `develop`
would propagate them and wipe the development record. This is the single most dangerous operation
in this repo.

`CHANGELOG.md` and `.release-please-manifest.json` are **main-owned** — release-please writes
them there. Never edit them on `develop`: as long as `develop` leaves them alone, git resolves
them cleanly on every promote (main's side wins), and the moment `develop` edits either one,
every promote conflicts. Sync them back by content copy, which carries no deletions:

```bash
git switch develop
git checkout main -- CHANGELOG.md .release-please-manifest.json
git commit -m "chore: sync changelog and manifest from the last release"
```

`README.md` must stay byte-identical on both branches: divergence would raise a merge conflict on
every promote, forever, and `main`'s copy is the landing page a consumer arriving from nuget.org
sees. Edit it on `develop` only.

There is no hotfix branch. An urgent fix goes `fix/*` → `develop` → promote; a second path into
`main` would add a way around CI without adding meaningful speed.

## Releasing

Tag-driven: `git tag v1.2.3 && git push origin v1.2.3` runs `.github/workflows/release.yml`,
which packs and publishes **both** `Ank.DocToolkit` and `Ank.DocToolkit.Extensions.DependencyInjection`
at that same version, in the same run. The **tag is the authoritative version**; the csproj
`<Version>` in each project is only a local dev default, so do not expect them to match. There is
no separate tag prefix for the extensions package — see "The DI extensions package" above for why.

**The tag is normally created by release-please, not by hand.**
`.github/workflows/release-please.yml` watches every push to `main`, computes the version bump
from Conventional Commits (`feat:` → minor, `fix:` → patch, `!`/`BREAKING CHANGE:` → major — this
mapping is release-please's built-in Conventional Commits strategy, not something configured in
this repo; `release-please-config.json` only maps commit types to changelog sections), and
maintains a single persistent Release PR with the computed `CHANGELOG.md` entry already written.
**Merging that PR is the human release decision** this project has always required — the
automation only replaces the manual "pick a version, write the changelog, tag it" bookkeeping
upstream of that decision, not the decision itself. Both packages are tracked as one component
(`"."` in `release-please-config.json`) so they can never version independently.
`release-please-config.json`'s `last-release-sha` seeds where the very first automated Release PR
starts counting commits from (main's tip right before this automation was added) — it exists only
to keep that first PR from sweeping in unrelated older history, goes inert once a real GitHub
Release exists to anchor from instead, and can be deleted from the config once that's happened.
**Pre-1.0, a breaking change (`!`/`BREAKING CHANGE:`) jumps straight to `1.0.0`**, not the next minor —
`bump-minor-pre-major` is deliberately left at its default (`false`).

**Any commit matching a recognized Conventional Commit type opens or updates the Release PR** —
not just `feat:`/`fix:`, and this is not gated by `changelog-sections` visibility (that only
controls what shows up *in* the changelog, not whether a release is proposed at all). Even a
`chore:`-only or `test:`-only merge (both `hidden: true`) proposes a release — it just produces a
changelog entry with a bare version heading and no visible body. **Check the Release PR's actual
diff before merging** — a version bump with an empty-looking changelog entry is a real signal to
hold off, not necessarily a release worth shipping; nothing stops it from otherwise merging and
publishing an empty version. Expect the Release PR to be open most of the time — that's normal,
not a bug; it just accumulates until you choose to merge it.

Merging the Release PR creates **both** the tag and a GitHub Release (release-please's own
generated notes, derived from the same `CHANGELOG.md` entry) — `release.yml`'s own
`Create or update GitHub Release` step detects that a release already exists and attaches the
`.nupkg`/`.snupkg` assets to it instead of creating a second one, rather than failing. **An
earlier version of this pipeline set `skip-github-release: true` to avoid that collision — don't
reintroduce it.** release-please's own documentation is explicit that skipping the GitHub Release
also skips creating the tag unless you build separate tagging infrastructure, which would have
meant a merged Release PR never actually triggers `release.yml` at all, silently. Verified against
release-please's own documentation directly, not assumed.

Commit messages must follow Conventional Commits (`type(scope)?: description`) going forward —
`ci.yml`'s `commit-format` job enforces the `type(scope)?: description` shape on every PR,
checking every commit in the PR's range (this repo true-merges, so every commit lands on `main`
and matters to the bump calculation, not just the PR title). By convention, not CI-enforced,
`scope` is `core`, `extensions`, or omitted — matching `CHANGELOG.md`'s own `Core:`/`Extensions:`
prefixes; the regex itself accepts any lowercase/digit/hyphen scope, so a different scope won't
fail CI, it just won't match that convention. Get the *type* wrong and release-please either
miscategorizes a change or silently drops it from the changelog — the CI guard exists so that's
caught at PR time, not discovered in a Release PR that's already wrong. **A `Merge branch 'develop'
into <feature>` commit fails this guard** (no type prefix) — rebase onto `develop` instead of
merging it into a long-lived feature branch. **Never rebase a feature branch onto `main`:** `main`
carries deletions of `CLAUDE.md`, `docs/` and `spike/`, so replaying your commits on top of it
drags those deletions into your branch, and from there into `develop` — the same loss
"Never merge `main` into `develop`" above warns about, reached by a detour.

`release-please.yml` needs its own PAT (not the default `GITHUB_TOKEN`) stored as the
`RELEASE_PLEASE_TOKEN` repository secret — GitHub Actions doesn't let a workflow's default token
trigger other workflows when it pushes, so a release-please-authored tag push would otherwise never
reach `release.yml`. Fine-grained PAT, scoped to this repo only, `Contents: read and write` +
`Pull requests: read and write`. Fine-grained PATs expire — rotate it before it does, or releases
will silently stop triggering. `release-please.yml` fails fast with a clear message if the secret
is missing, matching `release.yml`'s own `NUGET_USER` guard.

A manual `git tag v1.2.3 && git push origin v1.2.3` still works as a fallback — `release.yml` only
cares that a `v*` tag arrived, not how — but tag `main`, never `develop`, and write the
`## [X.Y.Z] - YYYY-MM-DD` heading into `CHANGELOG.md` **on `main`** first (`CHANGELOG.md` is
main-owned; see "Branches");
release-please only writes that entry when it's the one creating the tag, and never touches
`## Unreleased` at all (that heading is deliberately bracket-free — `## [Unreleased]` would
collide with release-please's own version-heading detection; don't put the brackets back). The
release workflow greps for the `## [X.Y.Z]` heading and refuses to publish if it's missing — the
same fail-fast treatment as the other premise guards. Once release-please starts handling most
releases, don't expect `## Unreleased` to keep accumulating content the way it did by hand —
anything added there manually won't automatically fold into the next automated Release PR, since
release-please computes its own entry from commit history, not from that section.

Publishing to nuget.org is **irreversible** — a version can be unlisted, never deleted or
replaced. The workflow therefore runs the full suite *and* all four premise guards (checked
against both projects, where applicable) before it pushes, and fails rather than shipping a
package that broke them. `--skip-duplicate` on the push step means a version already published
for one package but not the other is never an error. **Do not add a `continue-on-error` or bypass
to those steps.**

Authentication is **Trusted Publishing (OIDC)** — no long-lived API key exists. The job needs
`permissions: id-token: write`; without it the token request fails *silently* and `NuGet/login`
returns nothing. The temporary key lasts one hour, so the login step sits immediately before the
push, after every check has passed — do not hoist it to the top of the job.

Config lives on nuget.org (policy: owner `Ank-KhoaHo`, repo `DocToolkit`, workflow `release.yml`)
plus a `NUGET_USER` variable holding the nuget.org profile name, not an email. Never reintroduce a
stored API key for CI.

## Layout

```
src/DocToolkit/                                         the library
tests/DocToolkit.Tests/                                 182 tests, including StreamOverloadTests, AirGapGuardTests, DependencyGuardTests
src/DocToolkit.Extensions.DependencyInjection/          DI extensions package (services.AddDocToolkit())
tests/DocToolkit.Extensions.DependencyInjection.Tests/  42 tests, including ServiceCollectionExtensionsTests
samples/ConsoleSample/                                  core package, all five capabilities
samples/MinimalApiSample/                               DI extensions package, one endpoint per interface
docfx/                                                  DocFX site source, published to GitHub Pages on release
scripts/                                                promote-to-main.sh and its test — on main too, see Branches
spike/                                                  original proof-of-concept, kept as reference — do not modify (develop only)
docs/                                                   design docs and implementation plans this was built from (develop only)
```

The research behind the library selection lives in a separate, private knowledge base; the public
summary is in `README.md` under *Design notes*.
