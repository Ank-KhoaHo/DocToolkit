# Contributing to DocToolkit

Thanks for considering it. This package has a few constraints that are stricter than most, and
they are the reason it exists — so this document leads with what will get a pull request rejected,
rather than burying it.

## Getting set up

You need the .NET SDK. The projects target `net8.0` and `net10.0`, so you need a `net10.0`-capable
SDK to build everything.

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

### Checking Linux support

Linux support is verified, not assumed. CI runs the suite on Ubuntu; you can run the same check
locally:

```bash
docker build -f Dockerfile.linux-test -t doctoolkit-linux-test .
docker run --rm doctoolkit-linux-test
```

## Branching and pull requests

There is one branch: `main`. It cannot be pushed directly.

```bash
git switch -c feat/your-change main
# ... work, commit ...
git push -u origin feat/your-change
```

Then open a pull request into `main`. Branch names are `feat/**` for new capability and `fix/**`
for bug fixes. There is no hotfix path and no second way in. CI runs on your branch and on the
pull request.

Rebasing onto `main` is preferred over merging it in, for readable history:

```bash
git pull --rebase origin main
```

## Commit messages

Commits must follow [Conventional Commits](https://www.conventionalcommits.org/):

```
type(scope)?: description
```

CI checks **every commit in your pull request**, not just the title. This repository merges pull
requests with a real merge commit, so every one of your commits lands on `main` and is read by the
release tooling. Merge commits themselves are exempt — git writes those subjects and no prefix is
possible.

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
3. **Runs on Linux** — verified on Ubuntu in CI.
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

## When a check goes red

| Check | What it means |
|---|---|
| `build & test (ubuntu-latest)` / `(windows-latest)` | Both must pass. Linux is a supported platform, not a nice-to-have. |
| `no native binaries / no banned packages` | A dependency broke one of the four constraints. **Remove the package. Never relax the test.** |
| `commit message format` | A commit in your branch is not Conventional Commits. Amend or rebase. |
| `pack & verify .nupkg (core)` / `(extensions)` | The NuGet package no longer builds or verifies. |
| `build docs site` | The API documentation site failed to build. |

There is also a public-API approval test. It generates the shipped public surface and compares it
to a checked-in approved file. If it fails and **your change to the public API was intended**,
update the approved file in the source tree — `tests/DocToolkit.Tests/PublicApi/` and
`tests/DocToolkit.Extensions.DependencyInjection.Tests/PublicApi/` — and that diff becomes part of
the review. Editing the copy in the build output directory does nothing.

If it fails and you did **not** mean to change the public API, that is the test doing its job.

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

## Questions

Open an issue. If you are unsure whether a change fits the four constraints, ask before writing it
— that is a cheaper conversation than a rejected pull request.
