# DocToolkit enhancement backlog

**Status:** open · **Compiled:** 2026-08-03 · **Against:** `develop` at v0.3.2

A survey of everything worth improving in DocToolkit, across features, testing, CI/CD,
documentation and repo structure. This is a *backlog*, not a design — each item is a candidate
that still needs its own design doc and implementation plan before it gets built.

Items carry stable IDs (`A1`, `C7`, …). Refer to them by ID in later sessions; don't renumber.

Every claim below was checked against the working tree, not assumed. Where a gap is already
acknowledged in the project's own docs, the reference is given — those are the strongest
candidates, because the project has already conceded the point.

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

## Suggested order

Ranked by leverage against the project's own stated premise, not by size.

1. **C1 + C2 — Dependabot plus lockfiles and central package management.** Directly serves the
   premise that the resolved graph *is* the product. Everything needed to adjudicate the resulting
   PRs already exists in the premise guards. Self-contained: no public API change, no version bump.
2. **D1 — `CONTRIBUTING.md`, authored to live on `main`.** The process knowledge is currently
   invisible to exactly the audience that needs it.
3. **B1 — public-API approval tests.** One accidental source-breaking change has already shipped;
   pre-1.0 is the cheap moment to install the guard.
4. **A1–A5 — the create-and-template feature gap.** What turns a converter into a toolkit. Largest
   track by a wide margin and the only one that changes the public API surface — needs decomposing
   into several specs before any of it is built.
5. **C11 + C12 — the empty-release guard and generated third-party notices.** Both close hazards
   the project has already written down as hazards.
