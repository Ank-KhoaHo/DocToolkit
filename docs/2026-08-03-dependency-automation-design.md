# Dependency automation — design

Backlog items **C1** (automated dependency updates) and **C2** (lockfiles), from
`2026-08-03-enhancement-backlog.md`.

## Why

This package exists because it satisfies four constraints at once — permissive licences, NuGet
only, runs on Linux, no runtime network I/O — and every one of them is a property of the
**resolved dependency graph**, not of this code. `CLAUDE.md` and `README.md` both say so, and CI
already re-proves all four on every push.

What is missing is the other half of that posture. The guards can only adjudicate a change once
someone makes it, and nothing proposes changes. Dependencies go stale silently until a human
happens to look, and — worse — the graph the guards inspect is **re-resolved from scratch on every
restore**. Two restores a month apart can legitimately produce different closures from identical
source, and nothing records or reviews the difference.

So: something to propose bumps (C1), and something to make the resolved graph a committed,
reviewable artifact rather than a re-computation (C2). The guards that already exist do the
adjudicating; this design only supplies what they adjudicate.

## Scope

**In:** Dependabot config; `packages.lock.json` on the two packable projects; a locked-mode guard
in `ci.yml` and `release.yml`.

**Out:** Central Package Management (`Directory.Packages.props`) and `Directory.Build.props` —
see *Rejected alternatives*. Nothing about the public API, the branching model, or the release
pipeline's decision points changes.

## The constraint that shapes everything

[`dependabot-core#13950`](https://github.com/dependabot/dependabot-core/issues/13950) — opened
2026-01-15, still open, unassigned, no documented workaround. When a project holding a
`packages.lock.json` sits **downstream of a `ProjectReference`** whose graph changed, Dependabot
updates the upstream lockfile and misses the downstream one. Restore in locked mode then fails:

```
NU1004: The project references [X] whose dependencies has changed. The packages lock file is
inconsistent with the project dependencies so restore can't be run in locked mode.
```

That is not a corner case for this repo. The `ProjectReference` edges are:

```
src/DocToolkit                       PackageReference only   ← no ProjectReference in or out
src/…Extensions.DependencyInjection  PackageReference only   (Ank.DocToolkit is a *package* ref,
                                                              deliberately — see CLAUDE.md)
tests/DocToolkit.Tests                       ProjectReference → src/DocToolkit
tests/…Extensions.DependencyInjection.Tests  ProjectReference → src/…DependencyInjection
```

Both edges run tests → src. Lockfiles on **both sides** of those edges is precisely the broken
configuration, and would put every Dependabot PR into a red build.

## Lockfiles on `src/` only

Neither packable project consumes another project. So confining lockfiles to those two makes the
bug unreachable — there is no locked project downstream of a project reference.

That boundary is not a workaround dressed up as a principle. It is the boundary the premise guards
already use: `ci.yml` runs `dotnet list "$PROJECT" package --include-transitive` against exactly
those two projects and no others. The test projects' graph does not ship and carries none of the
licensing or native-binary promise. Locking precisely what the guards inspect, and nothing else,
makes the two mechanisms describe the same object.

**Accepted cost:** a test-only dependency (xunit, coverlet, Test.Sdk) can still resolve differently
between machines. That has never been what this repo's reproducibility claim is about, and if it
ever needs to be, Renovate is the escape hatch (below).

## Tool: Dependabot

Native to GitHub. No third-party app with write access to a repository that publishes to nuget.org
via OIDC, no PAT beyond the one release-please already needs, no privileged workflow writing to PR
branches next to an irreversible publish pipeline.

**Renovate was the serious alternative** and is better at this specific job: it delegates lockfile
regeneration to `dotnet restore` run as an external command
([docs](https://docs.renovatebot.com/modules/manager/nuget/)), so `#13950` does not apply to it at
all, and it groups and schedules more flexibly. It was rejected on supply-chain surface, not
capability — either the Mend GitHub App gets write access to this repo, or Renovate is self-hosted
as a scheduled workflow and maintained. For a repo whose entire security story is "nothing
long-lived exists to leak", adding a third-party write-scoped app is the wrong trade for the
convenience gained.

**If the src-only lockfile boundary ever needs to widen to the test projects, switch to Renovate
rather than adding a lockfile-fixer workflow.** That is the documented escape hatch.

## `.github/dependabot.yml`

```yaml
version: 2

updates:
  # ── The core library ────────────────────────────────────────────────────────
  # Its own block: SixLabors.Fonts is a direct reference only here, so the ignore
  # rule has to attach to the project it actually applies to.
  - package-ecosystem: nuget
    directory: "/src/DocToolkit"
    schedule:
      interval: weekly
    target-branch: develop
    commit-message:
      prefix: build
    open-pull-requests-limit: 10
    ignore:
      # 2.x moves to the Six Labors Split License — Apache-2.0 only under $1M
      # annual revenue. The [1.0.0] pin is the licensing wall. Blocking the major
      # here stops a doomed PR reopening every week; ci.yml's premise guard stays
      # the backstop if anything ever slips past.
      - dependency-name: "SixLabors.Fonts"
        update-types: ["version-update:semver-major"]

  # ── The extensions package ──────────────────────────────────────────────────
  # Every reference here is a FLOOR, and raising a floor narrows who can consume
  # the package. That is a compatibility decision, never routine maintenance, so
  # none of them are automated.
  - package-ecosystem: nuget
    directory: "/src/DocToolkit.Extensions.DependencyInjection"
    schedule:
      interval: weekly
    target-branch: develop
    commit-message:
      prefix: build
    ignore:
      - dependency-name: "Ank.DocToolkit"
      - dependency-name: "Microsoft.Extensions.DependencyInjection.Abstractions"
      - dependency-name: "Microsoft.Extensions.Options"

  # ── Test projects ───────────────────────────────────────────────────────────
  # No lockfiles here (see above), so #13950 cannot fire. Grouped into one PR:
  # these are noise, not risk.
  - package-ecosystem: nuget
    directories:
      - "/tests/DocToolkit.Tests"
      - "/tests/DocToolkit.Extensions.DependencyInjection.Tests"
    schedule:
      interval: weekly
    target-branch: develop
    commit-message:
      prefix: chore
    groups:
      test-dependencies:
        patterns: ["*"]

  # ── Samples: the external-consumer canary ───────────────────────────────────
  # Unlike the extensions package above, these floors SHOULD be bumped — that is
  # what keeps the canary armed. See "The samples finding" below.
  - package-ecosystem: nuget
    directories:
      - "/samples/ConsoleSample"
      - "/samples/MinimalApiSample"
    schedule:
      interval: weekly
    target-branch: develop
    commit-message:
      prefix: build

  # ── GitHub Actions ──────────────────────────────────────────────────────────
  # Feeds backlog item C3: pinning actions by SHA is only practical if something
  # keeps the pins current. Grouped — action bumps are reviewed as a set.
  - package-ecosystem: github-actions
    directory: "/"
    schedule:
      interval: weekly
    target-branch: develop
    commit-message:
      prefix: ci
    groups:
      actions:
        patterns: ["*"]
```

### Why each of those settings is not a default

**`target-branch: develop` is mandatory, not stylistic.** `ci.yml`'s `branch-policy` job rejects
any PR into `main` that is not a `release/promote-*` branch or release-please's own Release PR.
Omit this and *every* Dependabot PR fails CI on arrival.

**`commit-message.prefix` exists to satisfy `commit-format`.** That job enforces
`type(scope)?: description` on every commit in a PR's range. Dependabot's default subject
(`Bump X from A to B`) has no type prefix and fails it.

**The choice of prefix is a changelog decision.** `release-please-config.json` maps `build` →
*Changed* (visible) and `chore` → hidden. Production and sample dependency bumps use `build:`
because for this package the resolved graph **is** the product — a dependency change is exactly the
kind of thing a consumer needs to see. Test-only bumps use `chore:` because they do not ship.

The scope is omitted rather than `(deps)`, matching `CLAUDE.md`'s stated convention that scope is
`core`, `extensions`, or nothing. A `(deps)` scope would pass CI but read as off-convention.

## `packages.lock.json`

One line added in place to each packable project — never a wholesale csproj replacement, per
`CLAUDE.md`'s warning about the package metadata those files carry:

```xml
<RestorePackagesWithLockFile>true</RestorePackagesWithLockFile>
```

Generated with `dotnet restore <project> --force-evaluate`, and committed.

**A useful side effect worth stating plainly.** The extensions lockfile will record
`Ank.DocToolkit` at **0.2.0**, because NuGet resolves a minimum-version range to the *lowest*
satisfying version. That is precisely the behaviour the 0.3.0 changelog documents as having caused
a real shipped bug — the package built against a core release predating the `Stream` API it wrapped.
The lockfile turns that invisible resolution rule into a committed fact that shows up in a diff.

## The CI guard

```yaml
# The other premise guards check WHAT the resolved graph contains. This one checks
# it is the graph we agreed to. --locked-mode fails rather than silently
# re-resolving, so an upstream bump cannot enter a release without a lockfile
# change that someone reviewed.
#
# Restores the two packable projects individually, not the solution: tests/ and
# samples/ deliberately carry no lockfile, and a solution-wide --locked-mode would
# be ambiguous about them.
- name: Assert the resolved graph matches the committed lockfiles
  run: |
    for project in "$PROJECT" "$EXTENSIONS_PROJECT"; do
      echo "::group::$project"
      if ! dotnet restore "$project" --locked-mode; then
        echo "::error::$project restored differently from its committed packages.lock.json. If the dependency change is intended, run 'dotnet restore $project --force-evaluate' and commit the updated lockfile."
        exit 1
      fi
      echo "::endgroup::"
    done
```

**In `ci.yml`, this goes in the `premise-guard` job**, as the first step after `setup-dotnet` —
before that job's `dotnet build`, whose implicit restore would otherwise re-resolve ahead of the
check. `premise-guard` is the right home on the merits: this *is* a premise guard, and that job's
header already promises a failure there "is never a test bug."

**In `release.yml`, the same step joins the other three guards**, before pack. That follows the
existing rule that the irreversible path re-proves every guard rather than trusting that CI ran
earlier.

## The samples finding

Found while writing this design; logged as backlog item **B9**.

Both samples reference a floor — `Ank.DocToolkit [0.2.1, )` and
`Ank.DocToolkit.Extensions.DependencyInjection [0.2.1, )`. By the same min-version rule above, they
restore **0.2.1**, not the current 0.3.2.

`CLAUDE.md` says the samples exist to "prove the real published artifact works, not whatever is
currently on `main`", and that "a breaking API change fails the next sample build." Neither holds
while they are pinned to a floor that never moves. It is the same bug 0.3.0 fixed for the
extensions package, never applied here.

This design does not fix it directly — it makes it self-correcting. The samples block above lets
Dependabot raise those floors as new versions publish, which re-arms the canary. A one-off bump to
0.3.2 as part of rollout is worth doing anyway so the fix does not wait a week.

## Consequences accepted

- **More PRs than usual in the first week.** xunit 2.5.3, `Microsoft.NET.Test.Sdk` 17.8.0 and
  coverlet 6.0.0 are all well behind (backlog B6). That is debt being paid down, not a malfunction.
  Grouping keeps it modest: the whole test-dependency backlog arrives as one PR, and the actions
  backlog as another, so the burst is roughly one PR per production dependency plus two.
- **This makes backlog item C11 fire more often.** A `chore:`-only test-dependency bump proposes a
  release with an empty changelog body — the hazard `CLAUDE.md` already names and nothing currently
  guards. This design does not solve C11; it strengthens the case for it. Using `build:` for
  production bumps deliberately keeps the *shipped-graph* changes out of that category.
- **Test-project dependency resolution stays unpinned.** Stated above; revisit via Renovate if it
  ever matters.
- **Dependabot does not activate until the config reaches `main`.** Found during implementation,
  2026-08-03. GitHub reads `.github/dependabot.yml` from the repository's **default branch** only,
  and this repo's default branch is `main` — deliberately, since that is the tree a consumer
  arriving from nuget.org lands on. Merging to `develop` therefore changes nothing; the first
  Dependabot run happens after `scripts/promote-to-main.sh` carries the file to `main`. This does
  not argue for moving the file or changing the branching model: `target-branch: develop` is
  precisely the mechanism for this situation, letting Dependabot read config from `main` while
  opening every PR against `develop`, where `branch-policy` allows it.
- **Setting `target-branch` suppresses Dependabot *security* updates.** GitHub raises security
  updates for vulnerable manifests on the default branch only, "except where `target-branch` is
  used". Version updates — everything this design is actually about — are unaffected, but this is
  not a free setting, and it means Dependabot alerts remain the security signal rather than
  automatic security PRs. Accepted: the alternative is retargeting PRs at `main`, which
  `branch-policy` rejects by design.

## Rejected alternatives

**Central Package Management (`Directory.Packages.props`) in this change.** CPM combined with
lockfiles is the exact configuration `#13950` reports against. Folding it in here would reintroduce
the bug the whole design is shaped to avoid. It stays backlog item E1, decided separately.

**Dependabot plus a lockfile-fixer workflow** (regenerate lockfiles on Dependabot's branches, as in
the commonly-cited GitHub Actions workaround). It would preserve the wider src+tests lockfile scope,
but requires a workflow with write access to PR branches — a privilege-escalation surface sitting
beside an OIDC publish pipeline. The most careful work of the options, for the least gain.

**Lockfiles on `src/DocToolkit` alone.** Leaves the extensions package's shipped graph unpinned
even though it also publishes to nuget.org. No meaningful cost saved over doing both.

**Doing C1 without C2.** Gets proposal automation but leaves the graph re-resolved on every
restore, which is the specific thing C2 exists to fix.

## Rollout

1. Add `<RestorePackagesWithLockFile>true</…>` to the two `src/` project files.
2. `dotnet restore <project> --force-evaluate` for each; commit both lockfiles.
3. Add the guard step to `ci.yml` (`premise-guard`) and `release.yml`.
4. Add `.github/dependabot.yml`.
5. Bump both samples' floors to the current published version — `[0.3.2, )`, keeping the range form
   rather than switching to an exact pin, so the existing style holds and Dependabot carries it
   forward from there. Fixes B9 immediately rather than waiting for Dependabot's first run.
6. **Prove the guard discriminates.** Bump a version in a csproj *without* regenerating its
   lockfile, push to a scratch branch, confirm CI goes red with `NU1004`, revert. This repo holds
   its guards to that standard already — the air-gap suite is proved by mutation "so it
   discriminates rather than passing vacuously." A guard nobody has watched fail is not a guard.
7. On Dependabot's first run, confirm PRs target `develop`, subjects clear `commit-format`, and a
   src-project bump carries its lockfile in the same PR.

## To verify at implementation time

Two details that should be confirmed against current behaviour rather than trusted from this
document:

- **`directories` (plural) in a Dependabot block.** Used above for the test and sample blocks. If
  the configuration schema rejects it, split those into one block per directory — no other part of
  the design depends on it.
- **Dependabot's NuGet updater regenerating `packages.lock.json` for a project with no
  `ProjectReference` consumers.** This is the case `#13950` does *not* cover, and the whole
  src-only boundary rests on it. Rollout step 7 is the check; if it fails, the fallback is
  Renovate, per *Tool* above.

## Success criteria

- A Dependabot PR that bumps a core dependency updates `src/DocToolkit/packages.lock.json` in the
  same PR, targets `develop`, and passes `commit-format`, `premise-guard` and the new lockfile
  guard without human intervention.
- A hand-edited csproj version with a stale lockfile fails CI with an actionable message.
- A `SixLabors.Fonts` 2.x PR is never opened.
- No PR proposes raising a shipped floor on the extensions package.
