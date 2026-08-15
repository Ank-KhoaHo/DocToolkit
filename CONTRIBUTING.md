# Contributing to DocToolkit

This project has a [Code of Conduct](CODE_OF_CONDUCT.md). By participating you are expected
to uphold it; report unacceptable behaviour through the private channel it names.

Thanks for considering it. This package has a few constraints that are stricter than most, and
they are the reason it exists — so this document explains what will get a pull request rejected,
not just the day-to-day workflow of building and testing it.

## Getting set up

You need **.NET SDK 10.0.302**, and the .NET 8 SDK alongside it.

`global.json` pins the 10.x SDK to `10.0.302` with `rollForward: latestPatch`, so a newer patch in
the same feature band is fine but a different band is not — `dotnet` will refuse to run rather than
quietly build with something else. Before this pin, local builds and CI were on genuinely different
toolchains: `10.0.101` against `10.0.302`, two feature bands apart, with different Roslyn and
analyzers.

The .NET 8 SDK is still needed even though the 10.x SDK builds every target: the tests also *run*
the `net8.0` half of the suite, which needs the .NET 8 runtime. CI installs both for the same
reason.

```bash
git clone https://github.com/Ank-KhoaHo/DocToolkit.git
cd DocToolkit

dotnet build DocToolkit.sln -c Release
dotnet test  DocToolkit.sln -c Release
```

Every test runs once per target framework, so the result count is twice the test count.

### Checking warnings the way CI does

The build runs with warnings as errors and currently has none. Verify like this:

```bash
dotnet build DocToolkit.sln -c Release -warnaserror --no-incremental
```

**`--no-incremental` is not optional.** MSBuild skips projects whose inputs have not changed, and a
skipped project emits no diagnostics — so on an already-built tree the command cheerfully reports
`0 Warning(s)` while warnings exist. You will pass locally and fail CI, twice, before you work out
why. CI is unaffected because a fresh runner has nothing to skip.

### Checking cross-platform support

Platform support is verified, not assumed. CI runs the whole suite four times - Linux x64,
Windows, macOS (Apple Silicon) and Linux arm64 - because "pure managed" is a claim about every
platform .NET runs on, and a claim nothing exercises is a hope.

Linux is the one you can reproduce locally, and the container check is stricter than the CI
runner: it starts from a bare image, so it also proves the image's `COPY` list is complete.

```bash
docker build -f Dockerfile.linux-test -t doctoolkit-linux-test .
docker run --rm doctoolkit-linux-test
```

## Branching and pull requests

There is one branch: `main`. It cannot be pushed directly — every change arrives by pull request.

One exception, and it is not a development branch: `nuget-stats` is an orphan data branch written
once a day by `.github/workflows/nuget-stats.yml`. It shares no history with `main`, is never
merged into it, and carries no source. Nothing is developed there and no pull request targets it.

**If you have push access to this repository:**

```bash
git switch -c feat/your-change main
# ... work, commit ...
git push -u origin feat/your-change
```

**If you do not** — the normal case for an outside contributor, since this repo is public — fork it
on GitHub first, clone your fork, and add the original repository as `upstream` so you can branch
from its `main` and stay current with it:

```bash
git clone https://github.com/<your-username>/DocToolkit.git
cd DocToolkit
git remote add upstream https://github.com/Ank-KhoaHo/DocToolkit.git
git fetch upstream
git switch -c feat/your-change upstream/main
# ... work, commit ...
git push -u origin feat/your-change
```

Then open the pull request from your fork's branch into `Ank-KhoaHo/DocToolkit`'s `main`.

Branch names are not enforced by CI — use whichever commit type fits your change (`feat/**`,
`fix/**`, `chore/**`, `refactor/**`, and so on). What *is* enforced is which pushes get a CI run before
you open the pull request: only `main`, `feat/**` and `fix/**` trigger it on push, so a branch
under any other type gets its first CI run when the pull request opens, not before. There is no
hotfix path and no second way in — every change reaches `main` through a pull request.

Rebasing onto `main` is preferred over merging it in, for readable history:

```bash
git pull --rebase origin main
```

(From a fork, that's `git pull --rebase upstream main`.)

## Commit messages

Commits must follow [Conventional Commits](https://www.conventionalcommits.org/):

```
type(scope)?: description
```

**Title the pull request as a Conventional Commit — it is the only subject that reaches `main`.**
This repository **squash-merges, and merge commits are disabled**, so your branch becomes exactly
one commit whose subject is the pull request title and whose body is the pull request description.
The title is not decoration; it is the commit message `release-please` reads to build the changelog
and decide the version bump.

CI also checks **every commit in your branch**, not just the title. Under squash-merging those
individual subjects do not reach `main`, so this is a lower-stakes check than the title — it is kept
because a branch whose commits are already well-formed makes a good title obvious, and because the
setting could change.

**A title it cannot parse costs you the whole pull request.** release-please discards an
unparseable commit entirely, body included, so every `feat:` and `fix:` line on your branch
disappears from the changelog with it. Nothing fails: the checks stay green, the release succeeds,
and the code ships with nothing announcing it.

> Measured on 2026-08-10. PR #189 shipped the entire headers-and-footers feature under the title
> *"Headers and footers on generated documents"*, and 0.19.0's changelog records only an unrelated
> pull request that happened to share the release. A published release's notes cannot be rewritten.

**This advice used to say the opposite, and the whole reversal is worth knowing because the
repository spent a while in the worst of both states.** Under true merge commits GitHub copies the
pull request title into the merge commit's *body*, beside your real commits — so a prefixed title
produced **two identical changelog lines**, which is what happened to 0.15.0 and 0.16.0. Omitting
the prefix was the fix for that.

This page then claimed squash-merging had inverted it, and required a prefixed title again — while
merge commits were still the mode actually in use. Both statements could not be true, and the
duplicate came straight back: 0.27.2's entry was proposed twice, from the real commit and from the
merge commit's body. **Settled on 2026-08-15 by changing the repository rather than the prose**:
merge commits are disabled, squash is the only mode, and the squash body is the pull request
description — which, unlike a merge commit's body, cannot contain a second copy of the subject.

The `commit message format` check now asserts this on the title as well as on every commit, because
advice alone had already failed twice.

| Type | Use for | Appears in the public changelog |
|---|---|---|
| `feat` | a new capability | yes, under **Added** |
| `fix` | a bug fix | yes, under **Fixed** |
| `perf` | a performance change | yes |
| `build` | packaging or target frameworks | yes |
| `revert` | reverting a previous commit | yes |
| `docs` | documentation | no |
| `ci` | workflows and automation | no |
| `refactor` | internal restructuring, no behaviour change | no |
| `style` | formatting only, no behaviour change | no |
| `test` | tests only | no |
| `chore` | anything else | no |

`scope` is `core` or `extensions` by convention — the two packages — or omitted. The convention is
not enforced; any lowercase scope passes CI.

**Getting the type wrong has consequences beyond style.** The type decides the version bump and
which changelog section your change appears in. A `fix:` on something that is not a bug fix tells
consumers a library bug was fixed. A `feat:` on a docs change consumes a minor version.

### Marking a breaking change

Mark one with a `!` after the type and optional scope — `feat(core)!: ...` — or with a
`BREAKING CHANGE:` footer in the commit body. Either tells release-please to bump the version
beyond a patch.

It will not jump to `1.0.0`. `release-please-config.json` sets `bump-minor-pre-major: true`, so on
this package a breaking change bumps the **minor** version instead (0.5.0 → 0.6.0), because staying
below `1.0.0` is a deliberate, standing decision here, not an oversight. Do not assume `!` means
`1.0.0` — on this repository it does not.

## Four things that will get a pull request rejected

This package exists only because it satisfies four constraints at once. A change that breaks any
one of them makes the package pointless, so all four are enforced by tests rather than by review.

1. **Permissive licences only** — MIT, Apache-2.0, BSD. No revenue thresholds, no per-seat fees, no
   "free under $1M" tiers.
2. **NuGet only** — no browser download, no LibreOffice, no Office interop, and **no native
   binaries**. `dotnet restore` must be the only setup step.
3. **Runs everywhere .NET does** — the suite runs on Linux, Windows, macOS and arm64 in CI.
   "Pure managed" implies all of them, so all of them are measured rather than inferred.
4. **No runtime network I/O** — the library is used on air-gapped machines. No default code path
   may open a socket.

All four are properties of the *resolved dependency graph*, not of the code in this repository. A
single `dotnet add package` can break every one of them silently, which is why they are tested.

### Packages that cannot be added

`DependencyGuardTests` fails the build if any of these appear anywhere in the resolved graph,
matched by exact assembly name or namespace prefix — so `Spire` also catches `Spire.Doc`,
`Spire.Xls` and the rest of the family:

| Package | Why |
|---|---|
| `System.Drawing.Common` | throws `PlatformNotSupportedException` on Linux |
| `SkiaSharp` | pulls native binaries |
| `Magick.NET-Q16-AnyCPU` | pulls native binaries |
| `ShapeCrawler` | pulls both of the above |
| `EPPlus` | Polyform Noncommercial — not free for commercial use |
| `NPOI` | 2.8.0 and later require a paid maintenance fee |
| `Spire` | feature-capped free editions |
| `Syncfusion` | revenue-gated community licence |
| `QuestPDF` | revenue-gated community licence |
| `IronPdf` | commercial |

`System.Drawing.Common` is the nastiest of these: it restores and builds perfectly, then throws at
runtime on any non-Windows machine.

**ShapeCrawler is on the list from experience, not suspicion.** PPTX support originally used it.
It turned out to depend on SkiaSharp and Magick.NET, which put 38 native `.so`/`.dylib` files and
664 MB of `runtimes/` into build output. It was replaced with raw `DocumentFormat.OpenXml`. The
mistake that let it in was checking the library's *API* but never its *dependencies* — those are
different questions.

Before proposing any new dependency:

```bash
dotnet list package --include-transitive
find . -path '*/bin/*' \( -name '*.so' -o -name '*.so.*' -o -name '*.dylib' \)
```

If it's a direct dependency of `src/DocToolkit/` or `src/DocToolkit.Extensions.DependencyInjection/`,
also regenerate that project's lockfile and commit the result — otherwise CI's `--locked-mode`
restore fails on a dependency that was never actually banned:

```bash
dotnet restore src/DocToolkit/DocToolkit.csproj --force-evaluate
dotnet restore src/DocToolkit.Extensions.DependencyInjection/DocToolkit.Extensions.DependencyInjection.csproj --force-evaluate
```

## When a check goes red

| Check | What it means |
|---|---|
| `build & test (linux)` / `(windows)` / `(macos)` / `(linux-arm64)` | All four must pass. These are supported platforms, not nice-to-haves. **The check name carries the platform, not the runner image** - branch protection requires checks by name, so naming the image would make a required check a function of the image version, and repinning it would rename that check to one protection had never heard of, blocking every PR forever. The images themselves are pinned where a doc names a version (`ubuntu-24.04`, so the README stays checkable) and left as `latest` where none does (`macos-latest`). |
| `no native binaries / no banned packages` | This job runs several checks in order, so the failing step tells you the cause. If it fails on the **first** step (`dotnet restore --locked-mode`), `packages.lock.json` is just stale — run `dotnet restore <project> --force-evaluate` and commit the regenerated lockfile; nothing is banned. If it fails **later** — a native binary in build output, a banned package in the resolved graph, or `SixLabors.Fonts` drifting off the `1.x` line (2.x moves to a revenue-gated licence) — a dependency actually broke one of the four constraints. **Remove or re-pin it. Never relax the test.** If it fails on the **last** step (`THIRD-PARTY-NOTICES.txt matches the resolved graph`), nothing is wrong with the dependency — the attribution files just need regenerating: run `python scripts/gen-third-party-notices.py` and commit the result. |
| `build & test (linux)` — *Assert coverage has not regressed* | Line or branch coverage fell below the floor in `scripts/check-coverage.py`, which prints the assemblies that failed and the files with the most uncovered lines. **Write the test.** Lowering the floor is a deliberate decision to be argued for in the pull request, not the default repair. The floors carry the measurement that justifies them; if coverage has climbed well above one, the script says so and raising it is welcome. |
| `arm auto-merge if eligible` | Dependabot PRs only, and it never fails a human's PR — it is skipped. On a Dependabot PR it either arms auto-merge or prints why it held off: the update was not patch-level, or it touches a GitHub Action that `ci.yml` never runs (the publish, attestation and docs-deploy actions), where a green check proves nothing about the workflow that uses it. Held PRs are merged by hand after review. |
| `mutation score` | Not a pull-request check — it runs weekly and on demand (Actions → *Mutation testing* → *Run workflow*). It fails when the mutation score of the guard-critical files drops below the `break` threshold in `stryker-config.json`, which means a test stopped discriminating. The uploaded HTML report names every surviving mutant. Run it locally with `dotnet tool restore && dotnet stryker` (~4 minutes). |
| `linux container build & test` | The build failed inside a clean `mcr.microsoft.com/dotnet/sdk` image, which copies in only what `Dockerfile.linux-test` lists. It is not a duplicate of `build & test (linux)`: that job runs on a runner with the whole working tree present, so **only this one proves the image's `COPY` list is complete**. A new file the build needs, not added to that list, fails here and nowhere else. |
| `trimmed app publishes and runs` | `tests/TrimProbe` is trim-published over the real dependency closure and **executed**, with every capability asserting on its result. It fails either because a trim warning names `DocToolkit` — the `IsTrimmable` claim would no longer be true — or because the trimmed binary produced a wrong or empty document at runtime. A warning from a dependency (ClosedXML emits one) is reported, not fatal. |
| `AOT app publishes and runs` | The same shape, one step stronger: `tests/AotProbe` is native-AOT-published and executed. It backs the `IsAotCompatible` claim, which was set only once this job was green. **If it goes red, remove that attribute rather than weakening the job** — AOT breaks at runtime, when a type resolved by name turns out not to be there, so a publish that merely links proves nothing. |
| `formatting` | `dotnet format --verify-no-changes`, plus several derived guards — the failing step names which. The most common are `check-readme-coverage.py` (a shipped public type not named in the README its package publishes) and `gen-capability-matrix.py --check` (the docs capability table no longer matches the approved API; regenerate and commit). The full list is [below](#which-guard-runs-in-which-check). Note the third-party notices check is **not** here — it runs in `no native binaries / no banned packages`. |
| `commit message format` | A commit in your branch is not Conventional Commits — **or the pull request title is not**, which is checked separately because this repository squash-merges, so that title **is** the commit subject `release-please` parses. Amend, rebase, or retitle. Note that editing the title only re-runs CI because `edited` is in the workflow's `pull_request` types; re-running the job alone replays the old title. |
| `pack & verify .nupkg (core)` / `(extensions)` | The NuGet package no longer builds or verifies. |
| `build docs site` | The API documentation site failed to build. |
| `analyze (csharp)` | CodeQL static analysis. A failure here is the job breaking (usually the build step); a *finding* surfaces as a code-scanning alert on the pull request rather than as a red check. |

There is also a public-API approval test. It generates the shipped public surface and compares it
to a checked-in approved file. If it fails and **your change to the public API was intended**,
update the approved file in the source tree — `tests/DocToolkit.Tests/PublicApi/` and
`tests/DocToolkit.Extensions.DependencyInjection.Tests/PublicApi/` — and that diff becomes part of
the review. Editing the copy in the build output directory does nothing.

If it fails and you did **not** mean to change the public API, that is the test doing its job.

### Which guard runs in which check

Several checks run small Python guards from `scripts/`. When one goes red, this says which check
it belongs to, so you can find the failing step without reading the workflow.

**This table is generated** by `scripts/gen-guard-inventory.py` from the workflows, and CI fails
when it drifts. That is not ceremony: the `formatting` row above used to name one of its six
guards and credit it with a check that runs in a different job — which is the worst possible time
to be wrong, because you are reading it precisely when something has already failed.

<!-- BEGIN GENERATED (scripts/gen-guard-inventory.py) - do not edit by hand -->

| Check | Workflow | Guards it runs |
|---|---|---|
| `build & test (…)` | `ci.yml` | `check-coverage.py` |
| `formatting` | `ci.yml` | `check-configureawait.py`<br>`check-core-sharing.py`<br>`check-dependabot-scoping.py`<br>`check-doc-snippets.py`<br>`check-readme-coverage.py`<br>`check-workflow-tools.py`<br>`gen-capability-matrix.py`<br>`gen-guard-inventory.py` |
| `no native binaries / no banned packages` | `ci.yml` | `gen-third-party-notices.py`<br>`repair-lockfiles.py` |
| `arm auto-merge if eligible` | `dependabot-automerge.yml` | `automerge-eligible.py` |
| `outdated shipped dependencies` | `dependency-report.yml` | `check-dependabot-scoping.py`<br>`gen-third-party-notices.py` |
| `release-please` | `release-please.yml` | `extract-changelog-section.py` |

Generated from `.github/workflows/*.yml`. A guard added to a job appears here automatically; one moved between jobs moves with it. What each guard means, and what to do when it fails, is the table above — that part is written by hand because it cannot be derived.

<!-- END GENERATED -->

Run any of them locally the same way CI does — `python scripts/<name>` — before pushing.

## One thing that surprises people

**A brand-new public API cannot be demonstrated in `samples/` in the same pull request.**

Every sample references the *published* NuGet package with `Version="*"`, deliberately: the samples
exist to prove that the artifact a consumer actually restores works, not that this working tree
compiles. A method you merged but have not released does not exist from a sample's point of view,
and the build will fail on it.

So if you add a capability, its sample is follow-up work for after the next release. This is the
guarantee working, not a gap to route around — please do not switch a sample to a `ProjectReference`
to make it compile.

## Releases

Releases are cut by a maintainer, by hand. There is nothing for a contributor to do, and two things
to avoid:

- **Do not edit `CHANGELOG.md`.** It is generated from commit messages; your change would be
  overwritten by the next release.
- **Do not edit `.release-please-manifest.json`.** Same reason.

Both packages — `Ank.DocToolkit` and `Ank.DocToolkit.Extensions.DependencyInjection` — always ship
together at the same version.

### Why the release pull request is not merged automatically

Every other pull request here arms auto-merge. The release PR deliberately does not, and that is a
decision rather than an omission.

Merging it **publishes to nuget.org**, and a published version can be unlisted but never edited or
replaced. Everything upstream of that point is reversible: a bad commit can be reverted, a bad
merge can be undone, a broken `main` can be fixed forward. The moment the release PR merges, the
artifact is permanent and other people's builds can resolve it.

So it is the one gate where a human reads the generated changelog before it becomes public - which
is exactly where the duplicate-entry problem above gets caught, and where a change filed under the
wrong heading gets noticed while it can still be described properly in **Migrating**.

## Questions

Open an issue. If you are unsure whether a change fits the four constraints, ask before writing it
— that is a cheaper conversation than a rejected pull request.
