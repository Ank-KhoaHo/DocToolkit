# DocToolkit enhancement backlog

**Status:** open · **Compiled:** 2026-08-03 against `develop` at v0.3.2 ·
**Last refreshed:** 2026-08-03 against `main` at v0.3.9

A survey of everything worth improving in DocToolkit, across features, testing, CI/CD,
documentation and repo structure. This is a *backlog*, not a design — each item is a candidate
that still needs its own design doc and implementation plan before it gets built.

Items carry stable IDs (`A1`, `C7`, …). Refer to them by ID in later sessions; **don't renumber**,
and don't delete completed ones — mark them, so the record of what was tried survives.

Every claim below was checked against the working tree, not assumed. Where a gap is already
acknowledged in the project's own docs, the reference is given — those are the strongest
candidates, because the project has already conceded the point.

---

## Closed on 2026-08-03

Seven items, in two tracks — dependency automation and the single-branch collapse. Both have design
docs and implementation plans alongside this file.

| ID | Outcome |
|---|---|
| **C1** | **Done.** `.github/dependabot.yml` covers src, tests and GitHub Actions. Majors are excluded from groups so they arrive individually — the first grouped run bundled eight, including the action that drives publishing. |
| **C2** | **Done, with a caveat that became C18.** `packages.lock.json` on both packable projects, verified by a `--locked-mode` guard in `ci.yml` and `release.yml`. |
| **B6** | **Done.** xunit 2.9.3, `Microsoft.NET.Test.Sdk` 18.8.1, coverlet 10.0.1, `xunit.runner.visualstudio` 3.1.5. Coverage output re-verified, since `ci.yml` depends on cobertura files a plain `dotnet test` never produces. |
| **B9** | **Done, better than proposed.** The fix was not to bump the floor but to remove it: the samples now use `Version="*"`, so they track the newest release with nothing proposing a bump. Bumping the floor worked for exactly one release and then created a weekly release loop. |
| **C11** | **Moot, not fixed.** Publishing an empty release is now the *specified* behaviour — every merge to `main` publishes, by explicit decision. There is nothing left to guard against. |
| **C16** | **Moot.** The Codecov badge pointed at `branch/main` while development happened on `develop`. There is only `main` now. |
| **D1** | **Reduced, still open.** `CLAUDE.md` and `docs/` now live on `main`, so the process is no longer invisible to contributors. A `CONTRIBUTING.md` is still worth writing, but it is now a summary rather than the only copy. |

### Also fixed along the way

- **A flaky test that had been unsound since it was written.** `WorkbookEditorServiceTests` asserted
  byte-equality between two ClosedXML saves. ClosedXML re-stamps ZIP timestamps on save, so it
  passed only when both calls landed in the same two-second tick. It failed CI on a docs-only PR.
  A probe across all four formats found **only** ClosedXML non-deterministic — PPTX, DOCX and
  DOCX→PDF are byte-stable — so the byte-equality assertions in the other five service tests were
  confirmed sound and left alone.
- **`main`'s required status checks named a job that the single-branch collapse deletes.** Left
  unfixed it would have blocked every future PR permanently. Replaced with the seven jobs that
  survive, which also closed a pre-existing gap: the Windows build, the docs site and the extensions
  package verification had never gated `main`.
- **`strict` (require branches up to date) removed from `main`.** Under auto-release `main` moves
  twice per change, so every open PR went stale within minutes. It cost ten rebases across three
  PRs in one session, broke two PRs when the update was done by merge (`Merge branch 'main'` fails
  `commit-format`), and **silently stalled a Release PR** — release-please green, auto-merge
  enabled, nothing published.

---

## A. Library features

| ID | Gap | Evidence |
|---|---|---|
| A1 | **No "create a PPTX from scratch."** `PresentationEditor` has only `SlideCount`/`ExtractText`/`ReplaceText`. | Acknowledged in `CLAUDE.md` ("Samples and docs site"): `ConsoleSample` borrows `tests/DocToolkit.Tests/assets/sample.pptx` because the public API cannot produce one. |
| A2 | **No "create a DOCX from scratch."** The only way to obtain a DOCX is `HtmlToDocxConverter`. | `src/DocToolkit/DocxEditor.cs` is edit-only. |
| A3 | **XLSX surface is thin.** `Create` (single sheet), `ReadCell`, `SetCell` only. No multi-sheet, read-range / read-sheet, append-rows, list-sheet-names, formulas, or CSV import/export. | `src/DocToolkit/WorkbookEditor.cs` |
| A4 | **No repeating-table-row templating.** `ReplaceText` substitutes scalars only; there is no way to expand a table row per record — the most common real-world Word-template need (invoice line items). | `src/DocToolkit/DocxEditor.cs:34` |
| A5 | **No image insertion into a placeholder** (logo, signature, QR code). | Absent from the whole public API. |
| A6 | **No XLSX→PDF or PPTX→PDF.** Only DOCX→PDF exists, despite the package being named DocToolkit. Genuinely hard under the four constraints — needs an honest written decision either way, not silence. | `src/DocToolkit/DocxToPdfConverter.cs` is the only PDF producer. |
| A7 | **No DOCX→HTML / DOCX→Markdown.** `ExtractText` returns flat text with no structure. | `src/DocToolkit/DocxEditor.cs:144` |
| A8 | **No PDF utilities** — page count, merge, split, metadata. Requires a licence check first; PdfPig (Apache-2.0) and PDFsharp (MIT) are both plausible under the premise. | — |
| A9 | **No conversion options.** No page size, orientation, margins, header/footer injection, or external stylesheet on `HtmlToDocxConverter`. Conversion is fire-and-forget. | `src/DocToolkit/HtmlToDocxConverter.cs:25` |
| A10 | **`allowRemoteImageDownload: true` is an unguarded bool** — no timeout, no host allowlist, no maximum response size. A consumer who flips it hands their service an SSRF reach, which is precisely what the rest of the design exists to prevent. An options object would make the one sanctioned network path safe to use. | `src/DocToolkit/HtmlToDocxConverter.cs:46` |
| A11 | **File-path overloads are inconsistent.** The converters have `ConvertToFileAsync` / `ConvertFile`; the three editors have none. | Public API survey of `src/`. |
| A12 | **Zero observability.** No `ILogger`, no `ActivitySource`, no metrics anywhere — not even in the DI package, where a logger is already in the container. | `src/DocToolkit.Extensions.DependencyInjection/` |
| A13 | **DI options are frozen at registration.** No `IOptionsSnapshot`/`IOptionsMonitor`, no keyed registrations, singleton-only. | `ServiceCollectionExtensions.cs:20` |
| A14 | **No input size limits.** A `byte[]` API on a 200 MB document makes several full copies with nothing to stop it. | — |
| A15 | **No trim / AOT annotations** (`IsTrimmable`, `IsAotCompatible`). A pure-managed library is exactly the kind that *can* claim these, and it is a real differentiator against the native-binary alternatives. | `src/DocToolkit/DocToolkit.csproj` |

---

## B. Testing and quality

| ID | Gap |
|---|---|
| B1 | **No public-API approval test.** `CLAUDE.md` states that changing existing names or signatures is a breaking change for consumers, but nothing catches it mechanically. `PublicApiGenerator` + `Verify` would. The 0.3.0 changelog already documents one accidental source-breaking change (adding members to shipped interfaces), so this is not hypothetical. |
| B2 | **No coverage gate.** `ci.yml` uploads to Codecov with `fail_ci_if_error: false` and there is no `codecov.yml` target. Coverage is *reported*, never *enforced*. (`ci.yml:77-82`) |
| B3 | **No mutation testing.** `README.md` claims the air-gap guard "is proved by mutation" — done by hand, once, at authoring time. Stryker.NET would keep that claim true as the code changes. |
| B4 | **No property or fuzz tests on `RunTextSplicer`.** It is the trickiest code in the repo (mapping match offsets back onto individual runs) and has the quietest failure mode — silent formatting loss, no exception. |
| B5 | **No concurrency stress test.** The public API is documented as safe to call concurrently, and HtmlToOpenXml 3.5.0 is known to carry a non-thread-safe process-wide static `HttpClient`. Nothing proves the claim holds. |
| B6 | **Test stack is behind.** xunit 2.5.3, `Microsoft.NET.Test.Sdk` 17.8.0, coverlet 6.0.0 — on a repo that targets net10. xunit v3 exists. (`tests/DocToolkit.Tests/DocToolkit.Tests.csproj:12-17`) |
| B7 | **No benchmarks.** No perf baseline, and no regression detection for the `Stream`-vs-`byte[]` efficiency claim that headlined 0.2.0. |
| B8 | **No clean-machine restore smoke test.** Nothing proves the *published* `.nupkg` restores from nuget.org into an empty project and runs. The samples approximate this at build time, but only against whatever is already cached. |
| B9 | **The samples are a disarmed canary — same min-version bug 0.3.0 fixed for the extensions package.** Both samples reference a floor: `Ank.DocToolkit [0.2.1, )` (`ConsoleSample.csproj:12`) and `Ank.DocToolkit.Extensions.DependencyInjection [0.2.1, )` (`MinimalApiSample.csproj:11`). NuGet resolves a minimum-version range to the *lowest* satisfying version, so both restore **0.2.1**, not the current 0.3.2. `CLAUDE.md` states the samples exist to prove the published artifact works and that "a breaking API change fails the next sample build" — neither holds while they are pinned to the floor. Identical in kind to the extensions-package floor bug documented in the 0.3.0 changelog, never applied here. Found 2026-08-03 while designing C1. |

---

## C. CI/CD

| ID | Gap | Why it matters here specifically |
|---|---|---|
| C1 | **No Dependabot or Renovate.** | The largest single gap. The project's entire premise is that a single upstream bump can silently break the resolved graph — and there is no automated bump flow at all. It pairs exactly with what already exists: Dependabot proposes, the premise guards adjudicate. |
| C2 | **No lockfiles, no central package management.** No `packages.lock.json`, no `RestoreLockedMode` in CI, no `Directory.Packages.props`. | Same theme as C1. The graph this repo obsesses over is *re-resolved on every restore*, and the extensions package floats its references (`[8.0.0, )`). (`DocToolkit.Extensions.DependencyInjection.csproj:37-39`) |
| C3 | **Actions pinned by tag, not commit SHA.** `actions/checkout@v4`, `codecov/codecov-action@v5`, `NuGet/login@v1`. | For a repo that publishes to nuget.org via OIDC, a compromised action tag is the realistic attack path — there is no stored API key left to steal, so the workflow itself is the target. |
| C4 | **No `permissions:` block in `ci.yml`.** | `docs.yml` and `release.yml` both set one; `ci.yml` does not, so it inherits the repo default rather than least privilege. |
| C5 | **No `concurrency:` group on `release.yml`.** | Two tags pushed in quick succession race on an irreversible operation. |
| C6 | **`Dockerfile.linux-test` is never run in CI.** | It is documented as *the* Linux check, but CI uses ubuntu runners with the SDK preinstalled — which does not prove the container story most consumers actually deploy into. |
| C7 | **`ubuntu-latest` vs the documented `ubuntu-24.04`.** | `README.md` and `CLAUDE.md` both promise verification on `ubuntu-24.04`; `ci.yml:37` says `ubuntu-latest`. These diverge silently the day GitHub moves the alias. |
| C8 | **No macOS, no arm64 in the matrix.** | "Pure managed" implies both work. Nothing proves it. |
| C9 | **No CodeQL / code scanning** on a public repository. |
| C10 | **No SBOM, no package signing, no build-provenance attestation.** | The natural next step after Trusted Publishing, and consistent with the supply-chain posture the repo already takes. |
| C11 | **No empty-release guard.** | `CLAUDE.md` ("Releasing") warns that a `chore:`-only Release PR proposes a version with an empty changelog body, and that *"nothing stops it from otherwise merging and publishing an empty version."* A documented hazard with no automated check. |
| C12 | **`THIRD-PARTY-NOTICES.txt` is hand-maintained.** | Nothing regenerates or diffs it against the resolved graph, so the project's central licensing claim can go stale silently. The same applies to `README.md`'s hand-counted "19 packages — 18 MIT, 1 Apache-2.0". |
| C13 | **No `dotnet format --verify-no-changes`,** and no `.editorconfig` to verify against. |
| C14 | **DocFX version duplicated** across `ci.yml` and `docs.yml` (both `2.78.5`), with no `.config/dotnet-tools.json` to keep them in step. |
| C15 | **No `global.json`.** SDK version floats per contributor, on a repo that multi-targets net8 and net10. |
| C16 | **Codecov badge points at `branch/main`,** but development happens on `develop`, so the badge lags until a promote. |
| C17 | **No auto-merge for green patch-level dependency PRs.** Follows C1. |
| C18 | **Dependabot cannot write a correct lockfile for a multi-targeted project.** Found 2026-08-03, twice: it rewrote `src/DocToolkit/packages.lock.json` from 256 lines to 6, deleting both the `net8.0` and `net10.0` sections, with `NU1004: The project target frameworks are different than the lock file's target frameworks`. This is *not* [dependabot-core#13950](https://github.com/dependabot/dependabot-core/issues/13950), which is about `ProjectReference` consumers — it is a separate defect affecting any multi-target project. **Every future bump of the five shipped dependencies will arrive corrupt.** The locked-mode guard catches it every time, so this is noise rather than risk, and manual bumps take about ten minutes. The documented fallback is Renovate for the two `src/` projects, which regenerates lockfiles by running `dotnet restore` rather than modelling the graph — weighed and declined once on supply-chain grounds (a third-party app with write access to a repo publishing via OIDC). Revisit when the noise outweighs that. |
| C20 | **The Release PR should not auto-merge — versions are climbing too fast.** Requested 2026-08-03 after eleven versions shipped in a day (0.3.1 → 0.3.12), most containing no library change. `release-please.yml`'s `Auto-merge the Release PR` step makes every merge to `main` publish, which was the explicit decision when the single-branch model went in — but the churn it produces in practice is worse than the bookkeeping it removed. **Wanted:** release-please keeps computing the bump and writing the changelog and keeps its Release PR open, but a human merges it, so several merges batch into one release. This is a straight reversal of that decision, now informed by a day of watching it: remove or gate the auto-merge step, leaving everything else — the branch-name lookup, the corroborator check that fails loudly when the lookup finds nothing — exactly as it is. Note that the loud-failure hardening was added *because* releases were unattended; with a human gate it becomes belt-and-braces rather than essential, and is worth keeping either way. Also revisit backlog **C11**, which this un-moots: with batching restored, an empty-changelog release becomes a thing worth guarding against again. |
| C19 | **Dependabot `ignore` rules are scoped to the block that declares them, and grouped updates cross directories.** Found 2026-08-03: a grouped run from the `/tests/**` block proposed `SixLabors.Fonts [1.0.0] → [3.0.0]` — past the revenue-gated 2.x line the pin exists to prevent — plus three shipped floors ignored elsewhere. The test projects reach `src/` through `ProjectReference`, and Dependabot follows those and edits csproj files belonging to other blocks. Mitigated by repeating all four guards in the tests block, but the underlying trap remains: **any new update block must repeat every rule that protects `src/`**. A test asserting that would be better than a comment. |

---

## D. Documentation

| ID | Gap |
|---|---|
| D1 | **No `CONTRIBUTING.md` — the sharpest documentation gap in the repo.** The whole branching model (develop-as-trunk, main release-only, *never merge main into develop*, the promote script, Conventional Commits enforced by CI) lives **only in `CLAUDE.md`**, which `scripts/promote-to-main.sh` strips from `main` — the branch every consumer arriving from nuget.org lands on. An outside contributor sees one blockquote saying "target develop", then hits the `commit-format` and `branch-policy` CI failures with no document anywhere explaining them. |
| D2 | **No `SECURITY.md`.** Public repo, no disclosure path — despite the explicit SSRF-adjacent opt-in described in A10. |
| D3 | **No `CODE_OF_CONDUCT.md`, no issue templates, no PR template.** |
| D4 | **Docs site is API reference only.** No conceptual guides, no "fill a Word template" tutorial, no recipes. `_enableSearch: false` is correct (it avoids the ~109 MB Chromium download), but it means the site is not searchable — which raises, rather than lowers, the value of hand-written navigation and guides. |
| D5 | **No `<example>` / `<code>` blocks in XML doc comments,** so the generated site shows signatures with no usage snippets. |
| D6 | **No public ADRs.** The *why* — ShapeCrawler removed, SixLabors pinned, HTML→PDF pivoting through DOCX, one release workflow instead of two — is genuinely strong reasoning that exists only on `develop`, in `CLAUDE.md` and `docs/`, both stripped from `main`. |
| D7 | **No "known limitations" page.** A consumer cannot discover that XLSX→PDF does not exist, or that CSS layout fidelity is bounded, without reading source. |
| D8 | **No roadmap.** |
| D9 | **Three READMEs kept in sync by hand** (root, core package, extensions package), under a hard constraint that the root README stay byte-identical across `main` and `develop`. |
| D10 | **Sample gaps.** No worker service, no MVC/Razor, no Docker, and **no large-file streaming sample** — even though the `Stream` overloads were the headline feature of 0.2.0. |

---

## E. Structure and workflow

| ID | Gap |
|---|---|
| E1 | **No `Directory.Build.props`.** The two `src/` project files duplicate roughly fifteen properties (TFMs, nullable, LangVersion, Authors, licence expression, package output path, symbols, deterministic). They will drift, and the drift will be invisible until a package ships wrong. |
| E2 | **`ContinuousIntegrationBuild` is never set.** `<Deterministic>true</Deterministic>` alone does not normalize source paths in CI, so the published builds are not actually reproducible. |
| E3 | **No `EnableNETAnalyzers` / `AnalysisLevel` / `TreatWarningsAsErrors` in the project files.** `-warnaserror` is passed on the command line only, so a plain local `dotnet build` does not hold the line CI holds. |
| E4 | **csproj `<Version>` says 0.1.0 while 0.3.2 ships.** Documented as intentional (the tag is authoritative), but every local build and sample run reports 0.1.0. release-please's `extra-files` could sync it and remove the confusion. |
| E5 | **`net9.0` / `netstandard2.0` reach.** Currently net8 + net10 only. Widening is plausible but constrained by the dependency graph — worth a written decision either way rather than an unexamined default. |

---

## Checked and found sound

Recorded so a later pass does not re-investigate them:

- `.superpowers/sdd/` inside the repo is self-ignoring (its `.gitignore` contains `*`) — no tooling
  scratch is tracked in the public repo.
- SourceLink is enabled implicitly via the .NET 8+ SDK; `sourcelink.json` is present in build output.
- `docs.yml` and `release.yml` both set least-privilege `permissions:` blocks correctly. Only
  `ci.yml` is missing one (C4).

---

## Suggested order — revised 2026-08-03, after the first two tracks shipped

1. **A4 — repeating table rows in DOCX templates.** Two complete infrastructure tracks have now
   shipped and the library has gained nothing a *user* would notice. This is the most-requested
   real-world Word-template need and the first item here a consumer would see. Self-contained
   enough for a single spec.
2. **B1 — public-API approval tests.** Unchanged in priority. One accidental source-breaking change
   has already shipped, and pre-1.0 is the cheap moment to install the guard.
3. **A1–A3, A5 — the rest of the create-and-template gap**, each as its own spec.
4. **C3 — SHA-pin the actions.** Newly practical: all seven action majors are current as of today,
   so pinning no longer freezes them at stale versions.
5. **C12 — generate and diff `THIRD-PARTY-NOTICES.txt`.** Sharper than when written: the
   SixLabors 1.0.1 bump required hand-editing it in three separate places, which is precisely the
   drift this item predicts.

### The original ranking, kept for the record

C1+C2 first, then D1, B1, A1–A5, and C11+C12. C1 and C2 shipped; C11 turned out to be moot rather
than worth building; D1 shrank once `CLAUDE.md` reached `main`. The infrastructure-first ordering
was right — the guards it installed caught a licensing breach, two corrupted lockfiles and a
long-standing flaky test within hours of going live — but it is spent now, and the next return is
in the library itself.
