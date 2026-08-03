# Dependency automation — implementation plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development
> (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use
> checkbox (`- [ ]`) syntax for tracking.

**Goal:** Give DocToolkit automated dependency-update PRs and make the resolved dependency graph of
the two published packages a committed, CI-verified artifact instead of a per-restore recomputation.

**Architecture:** `packages.lock.json` on the two packable projects only — never on the test or
sample projects, because a lockfile downstream of a `ProjectReference` triggers an open Dependabot
bug. A new CI step restores those two projects with `--locked-mode` so a graph change cannot enter
a release without a reviewed lockfile diff. `.github/dependabot.yml` proposes the changes that the
existing premise guards then adjudicate.

**Tech Stack:** .NET SDK 8 + 10, NuGet lock files, GitHub Actions, Dependabot.

Design doc: `docs/2026-08-03-dependency-automation-design.md`. Read it before starting — it
explains *why* the lockfile boundary sits where it does, and that reasoning is load-bearing.

## Global Constraints

- **Branch:** all work happens on `develop` or a `feat/**`/`fix/**` branch cut from it. Never
  target `main`; `ci.yml`'s `branch-policy` job rejects it.
- **Never merge or rebase onto `main`** — it carries deletions of `CLAUDE.md`, `docs/` and `spike/`.
- **Commit messages must follow Conventional Commits** (`type(scope)?: description`). `ci.yml`'s
  `commit-format` job checks every commit in a PR's range.
- **Never add a `Co-Authored-By` trailer to any commit.**
- **Never edit `CHANGELOG.md` or `.release-please-manifest.json`** — they are main-owned.
- **Lockfiles go on `src/DocToolkit` and `src/DocToolkit.Extensions.DependencyInjection` only.**
  Adding one to a `tests/` or `samples/` project reintroduces the bug this design avoids.
- **Never relax a premise guard.** If one goes red, the dependency is wrong, not the guard.
- **The build runs at 0 warnings under `-warnaserror`.** Keep it there.
- Target frameworks are `net8.0;net10.0`; 224 tests × 2 TFMs = 448 results.

---

### Task 1: Generate and commit the lockfiles

**Files:**
- Modify: `src/DocToolkit/DocToolkit.csproj` (PropertyGroup at lines 3–9)
- Modify: `src/DocToolkit.Extensions.DependencyInjection/DocToolkit.Extensions.DependencyInjection.csproj` (PropertyGroup at lines 3–9)
- Create: `src/DocToolkit/packages.lock.json` (generated)
- Create: `src/DocToolkit.Extensions.DependencyInjection/packages.lock.json` (generated)

**Interfaces:**
- Consumes: nothing.
- Produces: two committed `packages.lock.json` files, and the `--locked-mode` restore behaviour that
  Task 2's CI step depends on.

Edit these csproj files **in place**. Never replace either wholesale — they carry the NuGet package
identity (`PackageId`, licence expression, readme, symbol package settings).

- [ ] **Step 1: Confirm locked mode currently proves nothing**

Run:

```bash
dotnet restore src/DocToolkit/DocToolkit.csproj --locked-mode
```

Expected: **succeeds**. There is no lockfile yet, so locked mode has nothing to compare against and
passes vacuously. This is the "test fails first" baseline — it shows the guard added in Task 2 would
be worthless without the rest of this task.

- [ ] **Step 2: Turn on lockfile generation for the core project**

In `src/DocToolkit/DocToolkit.csproj`, add one line to the **first** `PropertyGroup` (the one
containing `TargetFrameworks`), immediately after `<GenerateDocumentationFile>`:

```xml
    <GenerateDocumentationFile>true</GenerateDocumentationFile>
    <RestorePackagesWithLockFile>true</RestorePackagesWithLockFile>
```

- [ ] **Step 3: Turn on lockfile generation for the extensions project**

In `src/DocToolkit.Extensions.DependencyInjection/DocToolkit.Extensions.DependencyInjection.csproj`,
make the identical addition to its first `PropertyGroup`:

```xml
    <GenerateDocumentationFile>true</GenerateDocumentationFile>
    <RestorePackagesWithLockFile>true</RestorePackagesWithLockFile>
```

- [ ] **Step 4: Generate both lockfiles**

Run:

```bash
dotnet restore src/DocToolkit/DocToolkit.csproj --force-evaluate
dotnet restore src/DocToolkit.Extensions.DependencyInjection/DocToolkit.Extensions.DependencyInjection.csproj --force-evaluate
```

Expected: both succeed, and both files now exist:

```bash
ls src/DocToolkit/packages.lock.json
ls src/DocToolkit.Extensions.DependencyInjection/packages.lock.json
```

- [ ] **Step 5: Sanity-check what landed in the core lockfile**

Run:

```bash
grep -c '"net8.0"\|"net10.0"' src/DocToolkit/packages.lock.json
grep 'SixLabors.Fonts' src/DocToolkit/packages.lock.json
```

Expected: both target frameworks present, and `SixLabors.Fonts` resolved at **1.0.0**. If it shows
2.x, stop — the `[1.0.0]` pin has broken and that is a licensing problem, not a lockfile problem.

- [ ] **Step 6: Sanity-check the extensions lockfile records the floor, not the latest**

Run:

```bash
grep -A2 '"Ank.DocToolkit"' src/DocToolkit.Extensions.DependencyInjection/packages.lock.json
```

Expected: `"resolved": "0.2.0"` — **not** the latest published version. This is correct and is the
point: NuGet resolves a minimum-version range to the lowest satisfying version, and the lockfile
makes that normally-invisible rule a committed fact. Do not "fix" it.

- [ ] **Step 7: Verify locked mode now passes**

Run:

```bash
dotnet restore src/DocToolkit/DocToolkit.csproj --locked-mode
dotnet restore src/DocToolkit.Extensions.DependencyInjection/DocToolkit.Extensions.DependencyInjection.csproj --locked-mode
```

Expected: both succeed.

- [ ] **Step 8: Prove the guard discriminates — deliberately break it**

This repo holds its guards to the standard that they must be *seen* failing, not assumed to work.
Temporarily edit `src/DocToolkit/DocToolkit.csproj` and change the ClosedXML version:

```xml
    <PackageReference Include="ClosedXML" Version="0.105.0" />
```

Run:

```bash
dotnet restore src/DocToolkit/DocToolkit.csproj --locked-mode
```

Expected: **FAIL**, with an error resembling:

```
error NU1004: The packages lock file is inconsistent with the project dependencies so restore can't
be run in locked mode. Disable the RestoreLockedMode MSBuild property or pass an explicit
--force-evaluate option to run restore to update the lock file.
```

**If it succeeds, do not conclude the guard is broken — check for an IDE first.** Verified on
2026-08-03: with the repo open in VS Code, the C# extension watches the project files and runs its
own restore, *without* locked mode, within about two seconds of any csproj change. That rewrites
`packages.lock.json` to match the mutation, so the check that follows finds them in agreement and
passes legitimately. It looks exactly like a broken guard and is not one.

To tell the two apart, mutate the csproj and read the lockfile **without running restore yourself**:

```bash
git diff --stat src/DocToolkit/packages.lock.json
```

If the lockfile is already modified, an IDE restored behind you. Close the editor (or use a
checkout the editor does not have open) and repeat. CI has no file watcher, so this cannot happen
there.

Once the race is excluded, a genuine failure looks like:

```
error NU1004: The package reference ClosedXML version has changed from [0.105.1, ) to [0.105.0, ).
The packages lock file is inconsistent with the project dependencies so restore can't be run in
locked mode.
```

with exit code 1, and `packages.lock.json` **unmodified** — locked mode fails before writing.

If it still passes with no IDE running, re-check that `RestorePackagesWithLockFile` landed in the
PropertyGroup that also holds `TargetFrameworks`, and that `packages.lock.json` sits beside the
csproj.

- [ ] **Step 9: Revert the deliberate break**

Restore the version to `0.105.1`:

```xml
    <PackageReference Include="ClosedXML" Version="0.105.1" />
```

Confirm the working tree holds only the intended changes:

```bash
git status --short
git diff src/DocToolkit/DocToolkit.csproj
```

Expected: `git diff` shows **only** the added `RestorePackagesWithLockFile` line. If ClosedXML still
shows as modified, the revert did not take.

- [ ] **Step 10: Verify the full build and test suite still pass**

Run:

```bash
dotnet build DocToolkit.sln -c Release -warnaserror
dotnet test DocToolkit.sln -c Release --no-build
```

Expected: build succeeds with 0 warnings; 448 test results, all passing.

- [ ] **Step 11: Commit**

```bash
git add src/DocToolkit/DocToolkit.csproj \
        src/DocToolkit/packages.lock.json \
        src/DocToolkit.Extensions.DependencyInjection/DocToolkit.Extensions.DependencyInjection.csproj \
        src/DocToolkit.Extensions.DependencyInjection/packages.lock.json
git commit -m "build: lock the resolved dependency graph of both packages

Adds packages.lock.json to the two packable projects, so the graph that
carries the licensing and no-native-binary promise is a committed artifact
rather than something re-resolved on every restore.

Confined to src/ deliberately: dependabot-core#13950 breaks when a locked
project sits downstream of a ProjectReference, which is exactly the
tests -> src edge here. Neither packable project consumes another project,
so this boundary makes the bug unreachable, and it matches the boundary
the premise guards already inspect."
```

---

### Task 2: Add the lockfile guard to both workflows

**Files:**
- Modify: `.github/workflows/ci.yml` (`premise-guard` job, around lines 86–99)
- Modify: `.github/workflows/release.yml` (around lines 113–125)

**Interfaces:**
- Consumes: the committed lockfiles from Task 1, and the existing `PROJECT` / `EXTENSIONS_PROJECT`
  env vars already defined at the top of both workflows.
- Produces: a CI failure mode that Task 3's Dependabot PRs must satisfy.

**Placement note — this refines the design doc.** The design doc placed the release guard
"alongside the other three guards". It goes **before the `Build` step** in both workflows instead,
because `dotnet build` performs an implicit restore that re-resolves without locked mode. Checking
after that would still compare lockfile against csproj correctly, but the build would already have
run against a graph nobody verified. Guard first, then build.

- [ ] **Step 1: Add the guard to `ci.yml`'s `premise-guard` job**

In `.github/workflows/ci.yml`, inside the `premise-guard` job, insert this step **between** the
`actions/setup-dotnet@v4` step and the `- name: Build` step:

```yaml
      # The other premise guards check WHAT the resolved graph contains. This one
      # checks it is the graph we agreed to. --locked-mode fails rather than
      # silently re-resolving, so an upstream bump cannot enter a release without
      # a lockfile change that someone reviewed.
      #
      # Runs before Build because Build's implicit restore would otherwise
      # re-resolve first. Restores the two packable projects individually rather
      # than the solution: tests/ and samples/ deliberately carry no lockfile
      # (see docs/2026-08-03-dependency-automation-design.md), and a
      # solution-wide --locked-mode would be ambiguous about them.
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

- [ ] **Step 2: Add the same guard to `release.yml`**

In `.github/workflows/release.yml`, insert the identical step **between** the
`- name: Fail early if the nuget.org username is missing` step and the `- name: Build` step.

Use the same YAML as Step 1, but change the two `$PROJECT` / `$EXTENSIONS_PROJECT` shell references
to match that file's existing style, which interpolates them from the workflow `env` block:

```yaml
      - name: Assert the resolved graph matches the committed lockfiles
        run: |
          for project in "${{ env.PROJECT }}" "${{ env.EXTENSIONS_PROJECT }}"; do
            echo "::group::$project"
            if ! dotnet restore "$project" --locked-mode; then
              echo "::error::Refusing to publish - $project restored differently from its committed packages.lock.json. Run 'dotnet restore $project --force-evaluate' and commit the updated lockfile."
              exit 1
            fi
            echo "::endgroup::"
          done
```

Note the different error text: `release.yml`'s other guards all say "Refusing to publish", and this
one should match its neighbours.

- [ ] **Step 3: Verify both workflow files are valid YAML**

Run:

```bash
python3 -c "import yaml,sys; [yaml.safe_load(open(f)) for f in ['.github/workflows/ci.yml','.github/workflows/release.yml']]; print('both parse')"
```

Expected: `both parse`.

- [ ] **Step 4: Verify the guard command works exactly as CI will run it**

Run the loop locally with the same env vars CI sets:

```bash
PROJECT=src/DocToolkit/DocToolkit.csproj \
EXTENSIONS_PROJECT=src/DocToolkit.Extensions.DependencyInjection/DocToolkit.Extensions.DependencyInjection.csproj \
bash -c 'for project in "$PROJECT" "$EXTENSIONS_PROJECT"; do dotnet restore "$project" --locked-mode || exit 1; done; echo GUARD-PASSED'
```

Expected: ends with `GUARD-PASSED`.

- [ ] **Step 5: Commit**

```bash
git add .github/workflows/ci.yml .github/workflows/release.yml
git commit -m "ci: assert the resolved graph matches the committed lockfiles

Restores both packable projects with --locked-mode in premise-guard and
again in release.yml before packing, so a dependency change cannot reach
nuget.org without a lockfile diff someone reviewed.

Placed before Build in both workflows: Build's implicit restore would
otherwise re-resolve before the check ran."
```

---

### Task 3: Add the Dependabot configuration

**Files:**
- Create: `.github/dependabot.yml`

**Interfaces:**
- Consumes: the lockfiles from Task 1 and the guard from Task 2 — Dependabot's PRs must satisfy both.
- Produces: automated update PRs targeting `develop`.

- [ ] **Step 1: Create `.github/dependabot.yml`**

```yaml
version: 2

# Automated dependency updates. This repo's whole premise is that the RESOLVED
# dependency graph satisfies four constraints at once - permissive licences,
# NuGet only, Linux-safe, no runtime network I/O - and that a single upstream
# bump can break any of them silently. ci.yml's premise guards already
# adjudicate a change once someone makes it; this file is what proposes them.
#
# See docs/2026-08-03-dependency-automation-design.md for why lockfiles exist on
# src/ only, and why each ignore rule below is where it is.

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
      # 2.x moves to the Six Labors Split License - Apache-2.0 only under $1M
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
  # No lockfiles here (see the design doc), so dependabot-core#13950 cannot fire.
  # Grouped into one PR and prefixed chore: these do not ship, so they stay out
  # of the public changelog.
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
  # Unlike the extensions package above, these floors SHOULD be bumped - that is
  # what keeps the canary armed. A sample pinned to an old floor proves an old
  # release works, forever, and a breaking API change sails past it.
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
  # Actions are currently pinned by tag. Keeping them current is the thing that
  # makes SHA-pinning practical later rather than a maintenance burden.
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

- [ ] **Step 2: Verify it is valid YAML**

Run:

```bash
python3 -c "import yaml; d=yaml.safe_load(open('.github/dependabot.yml')); print('version', d['version'], '-', len(d['updates']), 'update blocks')"
```

Expected: `version 2 - 5 update blocks`.

- [ ] **Step 3: Verify every block targets `develop`**

This is the single setting that, if wrong, makes every Dependabot PR fail CI on arrival —
`branch-policy` rejects any PR into `main` that is not a promote or release-please branch.

```bash
python3 - <<'PY'
import yaml
d = yaml.safe_load(open('.github/dependabot.yml'))
bad = [u for u in d['updates'] if u.get('target-branch') != 'develop']
print('blocks missing target-branch develop:', len(bad))
assert not bad, bad
print('all blocks target develop')
PY
```

The quoted heredoc (`<<'PY'`) matters: it stops the shell expanding anything inside the program.
`ci.yml`'s own `package` job uses the same form for the same reason.

Expected: `all blocks target develop`.

- [ ] **Step 4: Verify every commit prefix clears the `commit-format` regex**

`ci.yml`'s `commit-format` job enforces
`^(feat|fix|docs|style|refactor|perf|test|build|ci|chore|revert)(\([a-z0-9_-]+\))?!?: .+$` on every
commit in a PR's range. Check each configured prefix produces a subject that matches:

```bash
python3 - <<'PY'
import re, yaml
regex = re.compile(r'^(feat|fix|docs|style|refactor|perf|test|build|ci|chore|revert)(\([a-z0-9_-]+\))?!?: .+$')
d = yaml.safe_load(open('.github/dependabot.yml'))
for u in d['updates']:
    prefix = u['commit-message']['prefix']
    subject = f'{prefix}: bump SomePackage from 1.0.0 to 1.0.1'
    assert regex.match(subject), f'FAILS commit-format: {subject}'
    print('OK  ' + subject)
PY
```

A quoted heredoc is required here, not a convenience: the regex contains `!` and `$`, both of which
an interactive shell would mangle inside a double-quoted `python3 -c` string.

Expected: five `OK` lines, no assertion error.

- [ ] **Step 5: Commit**

```bash
git add .github/dependabot.yml
git commit -m "ci: propose dependency updates automatically

The premise guards can only adjudicate a graph change once someone makes
it, and nothing proposed them. Weekly NuGet updates across src, tests and
samples, plus GitHub Actions.

Targets develop because branch-policy rejects PRs into main. Uses build:
for shipped and sample dependencies so they reach the changelog - for this
package the resolved graph is the product - and chore: for test-only
dependencies, which do not ship. SixLabors.Fonts majors and the extensions
package's version floors are not automated; both are decisions, not bumps."
```

---

### Task 4: Re-arm the samples canary (backlog B9)

**Files:**
- Modify: `samples/ConsoleSample/ConsoleSample.csproj:12`
- Modify: `samples/MinimalApiSample/MinimalApiSample.csproj:11`

**Interfaces:**
- Consumes: nothing from earlier tasks — independent, but grouped here because Task 3 is what keeps
  it fixed.
- Produces: samples that restore the current published packages.

Both samples reference a floor (`[0.2.1, )`). NuGet resolves a minimum-version range to the lowest
satisfying version, so both restore **0.2.1** regardless of what has shipped since. `CLAUDE.md`
claims the samples prove the published artifact works and that a breaking API change fails the next
sample build — neither holds while they are pinned to a floor that never moves.

- [ ] **Step 1: Find the actual latest published version — do not assume**

```bash
curl -s https://api.nuget.org/v3-flatcontainer/ank.doctoolkit/index.json
curl -s https://api.nuget.org/v3-flatcontainer/ank.doctoolkit.extensions.dependencyinjection/index.json
```

Expected: a JSON array of versions for each. Take the highest **stable** version present in *both*
listings — the two packages ship in lockstep, so they should match.

At the time this plan was written the newest entry in `CHANGELOG.md` is **0.3.2**, and the two code
blocks below use it. **If the API reports something newer, use that instead** — substitute it
everywhere `0.3.2` appears in this task.

If the two lists disagree on their highest version, stop and investigate: that means a release
published one package and not the other, which is a real problem worth understanding before
papering over it.

- [ ] **Step 2: Bump the ConsoleSample floor**

In `samples/ConsoleSample/ConsoleSample.csproj`, replace line 12:

```xml
    <PackageReference Include="Ank.DocToolkit" Version="[0.3.2, )" />
```

Keep the range form rather than switching to an exact pin — it matches the existing style, and
Dependabot carries it forward from here.

- [ ] **Step 3: Bump the MinimalApiSample floor**

In `samples/MinimalApiSample/MinimalApiSample.csproj`, replace line 11:

```xml
    <PackageReference Include="Ank.DocToolkit.Extensions.DependencyInjection" Version="[0.3.2, )" />
```

- [ ] **Step 4: Verify both samples still build against the newer packages**

This is the point of the change — if a real breaking change landed between 0.2.1 and now, it
surfaces here, which is exactly the canary doing its job.

```bash
dotnet build samples/ConsoleSample/ConsoleSample.csproj -c Release
dotnet build samples/MinimalApiSample/MinimalApiSample.csproj -c Release
```

Expected: both succeed. **If either fails to compile**, do not work around it — that is a genuine
breaking-change finding. Record it and stop for review.

- [ ] **Step 5: Verify the ConsoleSample actually runs**

```bash
dotnet run --project samples/ConsoleSample
```

Expected: completes without an unhandled exception.

- [ ] **Step 6: Commit**

```bash
git add samples/ConsoleSample/ConsoleSample.csproj samples/MinimalApiSample/MinimalApiSample.csproj
git commit -m "build: bump the sample package floors so the canary re-arms

Both samples referenced [0.2.1, ), and NuGet resolves a minimum-version
range to the LOWEST satisfying version - so they restored 0.2.1 no matter
what had shipped since. They proved an old release worked, and a breaking
API change would have sailed past them.

Same trap 0.3.0 fixed for the extensions package's own floor. Dependabot
now carries these forward, so it cannot silently recur."
```

---

### Task 5: Post-merge verification

**Files:** none — this is an observation checklist, run after the branch merges to `develop`.

**Interfaces:**
- Consumes: everything above, live on `develop`.
- Produces: confirmation, or the trigger to fall back to Renovate.

The design rests on one behaviour that cannot be verified locally: that Dependabot's NuGet updater
regenerates `packages.lock.json` for a project with **no** `ProjectReference` consumers. That is the
case `dependabot-core#13950` does not cover, and the whole src-only boundary depends on it.

- [ ] **Step 1: Confirm CI is green on `develop` with the new guard**

Check the Actions tab: the `premise-guard` job should now show the
`Assert the resolved graph matches the committed lockfiles` step passing.

- [ ] **Step 2: Promote to `main` — Dependabot cannot run until you do**

GitHub reads `.github/dependabot.yml` from the **default branch** only, and this repo's default
branch is `main`. Until `scripts/promote-to-main.sh` carries the file there, nothing happens, and
the absence of PRs means nothing. Run the promote and merge its PR with a **merge commit, never a
squash**, before treating any of the checks below as signal.

- [ ] **Step 3: Trigger Dependabot rather than waiting a week**

In the repository: **Insights → Dependency graph → Dependabot → Check for updates**.

- [ ] **Step 3: Verify the PRs that appear**

For each opened PR, confirm:

- it targets **`develop`**, not `main`;
- the commit subject matches `commit-format` (e.g. `build: bump ClosedXML from 0.105.1 to 0.106.0`);
- `commit-format`, `premise-guard`, `build-test`, `docs-build`, `package` and `promote-script` all pass.

- [ ] **Step 4: The load-bearing check — does a core-dependency PR carry its lockfile?**

Find a PR bumping a dependency of `src/DocToolkit` (ClosedXML, DocumentFormat.OpenXml,
HtmlToOpenXml.dll or OfficeIMO.Word.Pdf) and inspect its **Files changed**.

Expected: both `src/DocToolkit/DocToolkit.csproj` **and** `src/DocToolkit/packages.lock.json` are
modified in the same PR.

**If the lockfile is missing** and the new guard fails with `NU1004`: the src-only boundary is not
sufficient after all. Do not add a lockfile-fixer workflow. Switch to Renovate, per the *Tool*
section of the design doc — it regenerates lockfiles by running `dotnet restore` itself, so the
failure mode does not apply. Record the finding in the design doc before changing anything.

- [ ] **Step 5: Confirm no PR proposes something it should not**

Expected across the whole set: **no** `SixLabors.Fonts` 2.x PR, and **no** PR raising
`Ank.DocToolkit`, `Microsoft.Extensions.DependencyInjection.Abstractions` or
`Microsoft.Extensions.Options` in the extensions package.

If one appears, the corresponding `ignore` block is not matching — check the `dependency-name`
spelling against the csproj exactly.

- [ ] **Step 6: Record the outcome**

Append a short "Verified on <date>" note to the design doc's *To verify at implementation time*
section, stating what the first Dependabot run actually did. Commit as
`docs: record the dependency-automation verification outcome`.

---

## Notes for the reviewer

- **Tasks 1–4 are independently reviewable and independently revertable.** Task 4 (the samples fix)
  is the one that could surface a genuine breaking-change finding; if it does, that is a result, not
  a blocker to the rest.
- **Two things in this plan were deliberately not verified in advance** and are checked by running
  it: that `directories` (plural) is accepted in a Dependabot block (Task 3, Step 2 — if it is
  rejected, split those into one block per directory; nothing else depends on it), and the Task 5
  Step 4 lockfile behaviour.
- **The plan does not touch `CHANGELOG.md`.** release-please computes the entry from the commit
  subjects above, and `CHANGELOG.md` is main-owned.
