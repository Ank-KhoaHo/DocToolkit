# Semantic-Release Automation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Automate the version-number-and-changelog bookkeeping that currently precedes every
manual release, via release-please, without changing `release.yml`, the CHANGELOG guard, or the
principle that a human deliberately decides to publish.

**Architecture:** `release-please-config.json` + `.release-please-manifest.json` configure
`googleapis/release-please-action@v4` (new `.github/workflows/release-please.yml`, triggered on
every push to `main`) to track both packages as one root component, maintaining a persistent
"Release PR" with a computed version bump and generated `CHANGELOG.md` entry. Merging that PR
pushes the `vX.Y.Z` tag, which `release.yml` picks up completely unchanged. A new `commit-format`
job in `ci.yml` enforces Conventional Commits on every PR so the bump/changelog inference is never
silently wrong.

**Tech Stack:** GitHub Actions (`googleapis/release-please-action@v4`), plain bash (no new
language runtime — deliberately not `commitlint`/Node.js).

## Global Constraints

- **`release-type: "simple"`, no `extra-files` configured** — release-please must never edit any
  `.csproj`. `<Version>` in both projects stays a local dev default, untouched, exactly as today.
- **Both packages tracked as one component (`"."`)** — they must never version independently.
  This is what makes the existing lockstep guarantee hold with zero special-casing.
- **`include-component-in-tag: false`** — required, not optional (defaults to `true`). Without it,
  tags come out as `component-v1.2.3` instead of the plain `v1.2.3` that `release.yml`'s
  `push: tags: ["v*"]` trigger expects.
- **`skip-github-release: true`** — required. release-please creates a GitHub Release by default;
  `release.yml` already creates one (`Create GitHub Release`, with `--generate-notes`, after tests
  and every guard pass). Without this flag, both would try to create a release for the same tag and
  the second one would fail outright, on every release.
- **`release-please.yml` must use `secrets.RELEASE_PLEASE_TOKEN`, never the default
  `GITHUB_TOKEN`.** GitHub Actions does not let a workflow's default token trigger *other*
  workflows when it pushes (loop prevention) — using the default token would mean release-please's
  tag push silently never reaches `release.yml`. Creating the actual PAT and adding it as a
  repository secret is a manual step outside what any task in this plan can do (see the note after
  Task 2) — the workflow file must be written to expect it regardless.
- **`changelog-sections` maps `feat`→`Added`, `fix`→`Fixed`, `docs`/`ci`/`build`/`refactor`→`Changed`,
  `chore`/`test` hidden** — matches this repo's existing Keep-a-Changelog categories and its
  established convention of documenting CI/pipeline changes (not release-please's stock
  feat/fix-only defaults).
- **Commit format regex, exact:**
  `^(feat|fix|docs|style|refactor|perf|test|build|ci|chore|revert)(\([a-z0-9_-]+\))?!?: .+$`
  — type, optional `(scope)`, optional `!` (breaking-change marker), `: `, non-empty description.
- **No `Co-Authored-By` trailer in any commit message.** Repo-wide convention, no exceptions.
- **Manifest seed version is `0.2.1`** — the currently published version as of this plan (verified
  against nuget.org, not assumed, immediately before writing this plan).

---

### Task 1: release-please config files

**Files:**
- Create: `release-please-config.json`
- Create: `.release-please-manifest.json`

**Interfaces:**
- Produces: the config `release-please-action` (Task 2) reads by default from the repo root — no
  path needs passing to the action, it looks for these two exact filenames there.

- [ ] **Step 1: Create `release-please-config.json`**

```json
{
  "$schema": "https://raw.githubusercontent.com/googleapis/release-please/main/schemas/config.json",
  "include-component-in-tag": false,
  "skip-github-release": true,
  "packages": {
    ".": {
      "release-type": "simple",
      "changelog-path": "CHANGELOG.md",
      "changelog-sections": [
        { "type": "feat", "section": "Added" },
        { "type": "fix", "section": "Fixed" },
        { "type": "docs", "section": "Changed" },
        { "type": "ci", "section": "Changed" },
        { "type": "build", "section": "Changed" },
        { "type": "refactor", "section": "Changed" },
        { "type": "chore", "section": "Changed", "hidden": true },
        { "type": "test", "section": "Changed", "hidden": true }
      ]
    }
  }
}
```

- [ ] **Step 2: Create `.release-please-manifest.json`**

```json
{
  ".": "0.2.1"
}
```

- [ ] **Step 3: Verify both files are valid JSON**

Run: `python3 -c "import json; json.load(open('release-please-config.json')); print('config ok')"`
Expected: `config ok`

Run: `python3 -c "import json; json.load(open('.release-please-manifest.json')); print('manifest ok')"`
Expected: `manifest ok`

- [ ] **Step 4: Commit**

```bash
git add release-please-config.json .release-please-manifest.json
git commit -m "ci: add release-please config, seeded at the current 0.2.1"
```

---

### Task 2: release-please workflow

**Files:**
- Create: `.github/workflows/release-please.yml`

**Interfaces:**
- Consumes: `release-please-config.json` / `.release-please-manifest.json` (Task 1), read
  implicitly by `release-please-action` from the repo root.
- Consumes: `secrets.RELEASE_PLEASE_TOKEN` — does not exist yet; see the note after this task.

- [ ] **Step 1: Create the workflow**

```yaml
name: release-please

# Maintains a persistent "Release PR" on main: as commits land in Conventional Commits
# format, this computes the next version and writes the CHANGELOG.md entry. Merging that
# PR is the human release decision this project has always required - the automation only
# replaces the manual "pick a version, write the changelog, tag it" bookkeeping upstream of
# it, not the decision itself. See release-please-config.json for the version/changelog
# rules, and CLAUDE.md's "Releasing" section for the full flow.
#
# Requires a Personal Access Token, not the default GITHUB_TOKEN: GitHub Actions does not
# let a workflow's default token trigger OTHER workflows when it pushes a commit or tag
# (loop prevention). Using the default token here would mean the tag this action pushes on
# Release-PR merge never reaches release.yml's `push: tags: ["v*"]` trigger - no error,
# the tag would just sit there unpublished.
#
# Setup, once: create a fine-grained PAT scoped to only this repo, with
# "Contents: read and write" and "Pull requests: read and write" permissions, and add it as
# a repository secret named RELEASE_PLEASE_TOKEN. Fine-grained PATs expire - rotate it
# before it does, or releases silently stop triggering.

on:
  push:
    branches: [main]

permissions:
  contents: write
  pull-requests: write

jobs:
  release-please:
    runs-on: ubuntu-latest
    steps:
      - uses: googleapis/release-please-action@v4
        with:
          token: ${{ secrets.RELEASE_PLEASE_TOKEN }}
```

- [ ] **Step 2: Verify the YAML is well-formed**

Run: `python3 -c "import yaml; yaml.safe_load(open('.github/workflows/release-please.yml')); print('valid')"`
Expected: `valid`

- [ ] **Step 3: Commit**

```bash
git add .github/workflows/release-please.yml
git commit -m "ci: add release-please workflow"
```

**Manual step required after this task, before release-please can do anything (not automatable
by any task in this plan):** create a fine-grained GitHub PAT scoped to this repo only, with
`Contents: read and write` and `Pull requests: read and write` permissions, and add it as a
repository secret named `RELEASE_PLEASE_TOKEN`. Until this exists, `release-please.yml` will run
on every push to `main` and fail at the action step with an authentication error — a clear,
diagnosable failure, not a silent no-op.

---

### Task 3: commit message format guard

**Files:**
- Modify: `.github/workflows/ci.yml`

**Interfaces:**
- Consumes: the exact regex from Global Constraints above.
- Produces: nothing consumed by later tasks — this is a leaf job.

**Context:** `ci.yml` already has three jobs following one pattern — `build-test`,
`premise-guard`, `docs-build` — each a separate top-level job under `jobs:`, `runs-on:
ubuntu-latest`, plain bash `run:` steps with `::error::` annotations on failure (see
`premise-guard`'s "Assert zero native binaries" step for the exact style to match). This task adds
a fourth job in that same style, `commit-format`, gated to `pull_request` events only — on a
`push` to `main`, every commit already passed this check while it was in its PR, so there's
nothing new to check.

- [ ] **Step 1: Add the `commit-format` job**

In `.github/workflows/ci.yml`, add this job after `premise-guard`'s closing (after its "Assert
SixLabors.Fonts stayed on 1.x" step) and before the `docs-build:` job definition:

```yaml
  # Enforces Conventional Commits so release-please's version-bump and changelog inference
  # (release-please-config.json) is never silently wrong. Checks every commit in the PR's
  # range, not just the PR title, because this repo true-merges - every individual commit
  # lands on main and matters to release-please's bump calculation.
  commit-format:
    name: commit message format
    runs-on: ubuntu-latest
    if: github.event_name == 'pull_request'
    steps:
      - uses: actions/checkout@v4
        with:
          fetch-depth: 0

      - name: Assert every commit in this PR follows Conventional Commits
        run: |
          regex='^(feat|fix|docs|style|refactor|perf|test|build|ci|chore|revert)(\([a-z0-9_-]+\))?!?: .+$'
          base="${{ github.event.pull_request.base.sha }}"
          head="${{ github.event.pull_request.head.sha }}"
          bad=0
          while IFS= read -r line; do
            sha="${line%% *}"
            subject="${line#* }"
            if ! [[ "$subject" =~ $regex ]]; then
              echo "::error::Commit $sha does not match Conventional Commits format (type(scope)?: description): \"$subject\""
              bad=1
            fi
          done < <(git log --format='%H %s' "$base..$head")
          if [ "$bad" -ne 0 ]; then
            echo "::error::One or more commits do not follow Conventional Commits format."
            exit 1
          fi
          echo "all commits in this PR match Conventional Commits format"
```

- [ ] **Step 2: Verify the YAML is well-formed**

Run: `python3 -c "import yaml; yaml.safe_load(open('.github/workflows/ci.yml')); print('valid')"`
Expected: `valid`

- [ ] **Step 3: Test the regex logic locally against known good and bad examples**

The job's `if [[ ... =~ $regex ]]` line is the part that must discriminate correctly. Test it
standalone before trusting it in CI — this exact loop, run locally against a crafted list rather
than real git history, so both a pass and a fail case are proven in one run:

```bash
regex='^(feat|fix|docs|style|refactor|perf|test|build|ci|chore|revert)(\([a-z0-9_-]+\))?!?: .+$'
test_subjects=(
  "feat(core): add HtmlToDocxConverter overload|PASS"
  "fix(extensions): correct AddDocToolkit registration|PASS"
  "docs: update README|PASS"
  "ci!: breaking change to release pipeline|PASS"
  "samples: add ConsoleSample exercising all five core capabilities|FAIL"
  "add ConsoleSample|FAIL"
  "Merge pull request #2 from Ank-KhoaHo/feat/docs-samples-site|FAIL"
)
for entry in "${test_subjects[@]}"; do
  subject="${entry%%|*}"
  expected="${entry##*|}"
  if [[ "$subject" =~ $regex ]]; then actual="PASS"; else actual="FAIL"; fi
  if [ "$actual" = "$expected" ]; then
    echo "ok:   \"$subject\" -> $actual"
  else
    echo "WRONG: \"$subject\" -> $actual, expected $expected"
  fi
done
```

Expected: every line prints `ok:`, none print `WRONG:`. Note the `samples: add ConsoleSample...`
case is a real commit subject from this repo's own history (`04bf9bc`) — it correctly fails, which
is exactly why this guard is being added now: that commit predates this convention and would not
have passed it.

If any line prints `WRONG:`, the regex itself needs fixing before proceeding — do not adjust the
test cases to match a broken regex.

- [ ] **Step 4: Commit**

```bash
git add .github/workflows/ci.yml
git commit -m "ci: enforce Conventional Commits format on every PR commit"
```

---

### Task 4: documentation updates

**Files:**
- Modify: `README.md`
- Modify: `CLAUDE.md`

**Interfaces:**
- Consumes: nothing new — describes Tasks 1-3's outputs in prose.

**Context:** Both files currently describe releasing as "tag manually, `release.yml` does the
rest." Both need a new explanation of how the tag now actually gets created, without touching the
existing description of what `release.yml` itself does (unchanged). Read each file's current
`## Releasing` section in full before editing — the exact current text is reproduced below so the
edit is unambiguous, but confirm it still matches before replacing (if it doesn't, something
changed since this plan was written; stop and report rather than guessing which version is
current).

- [ ] **Step 1: Update `README.md`'s Releasing section**

Find this exact block in `README.md`:

```markdown
## Releasing

Publishing is driven by tags, and **one tag releases both packages at the same version.**
`Ank.DocToolkit` and `Ank.DocToolkit.Extensions.DependencyInjection` are meant to stay in
lockstep — the tag is the single source of truth for the version; the `<Version>` in each csproj
is only a local dev default.

```bash
git tag v1.0.0
git push origin v1.0.0
```

That runs [`.github/workflows/release.yml`](.github/workflows/release.yml), which **re-proves
```

Replace it with (inserting a new paragraph between the intro and the `git tag` example, and
changing the introduction to the tag example):

```markdown
## Releasing

Publishing is driven by tags, and **one tag releases both packages at the same version.**
`Ank.DocToolkit` and `Ank.DocToolkit.Extensions.DependencyInjection` are meant to stay in
lockstep — the tag is the single source of truth for the version; the `<Version>` in each csproj
is only a local dev default.

**The tag is normally created automatically, not by hand.**
[release-please](https://github.com/googleapis/release-please) watches commits on `main` (in
[Conventional Commits](https://www.conventionalcommits.org/) format — `feat:`/`fix:`/etc.) and
maintains a single "Release PR" with the computed version bump and a generated `CHANGELOG.md`
entry. **Merging that PR is the release decision** — review it like any other PR, hand-edit the
changelog prose if you want, then merge; release-please pushes the `vX.Y.Z` tag as part of that
merge, which is all `release.yml` (below) actually needs to fire.

A manual tag still works too — `release.yml` only cares that a `v*` tag arrived, not how:

```bash
git tag v1.0.0
git push origin v1.0.0
```

That runs [`.github/workflows/release.yml`](.github/workflows/release.yml), which **re-proves
```

- [ ] **Step 2: Verify the README edit landed correctly**

Run: `grep -n "release-please\|git tag v1.0.0" README.md`
Expected: both the new `release-please` paragraph and the (still-present, now second) `git tag
v1.0.0` example appear, with the new paragraph appearing first in the file.

- [ ] **Step 3: Update `CLAUDE.md`'s Releasing section**

Find this exact block in `CLAUDE.md`:

```markdown
## Releasing

Tag-driven: `git tag v1.2.3 && git push origin v1.2.3` runs `.github/workflows/release.yml`,
which packs and publishes **both** `Ank.DocToolkit` and `Ank.DocToolkit.Extensions.DependencyInjection`
at that same version, in the same run. The **tag is the authoritative version**; the csproj
`<Version>` in each project is only a local dev default, so do not expect them to match. There is
no separate tag prefix for the extensions package — see "The DI extensions package" above for why.

**Before tagging, move `[Unreleased]` in `CHANGELOG.md` under a new `## [X.Y.Z] - YYYY-MM-DD`
heading for the version you're about to release.** The release workflow greps for that heading
and refuses to publish if it's missing — the same fail-fast treatment as the other premise guards,
because a changelog gap is easy to forget under release pressure and easy to fix beforehand.
```

Replace it with:

```markdown
## Releasing

Tag-driven: `git tag v1.2.3 && git push origin v1.2.3` runs `.github/workflows/release.yml`,
which packs and publishes **both** `Ank.DocToolkit` and `Ank.DocToolkit.Extensions.DependencyInjection`
at that same version, in the same run. The **tag is the authoritative version**; the csproj
`<Version>` in each project is only a local dev default, so do not expect them to match. There is
no separate tag prefix for the extensions package — see "The DI extensions package" above for why.

**The tag is normally created by release-please, not by hand.**
`.github/workflows/release-please.yml` watches every push to `main`, computes the version bump
from Conventional Commits (`feat:` → minor, `fix:` → patch, `!`/`BREAKING CHANGE:` → major — see
`release-please-config.json`), and maintains a single persistent Release PR with the computed
`CHANGELOG.md` entry already written. **Merging that PR is the human release decision** this
project has always required — the automation only replaces the manual "pick a version, write the
changelog, tag it" bookkeeping upstream of that decision, not the decision itself. Both packages
are tracked as one component (`"."` in `release-please-config.json`) so they can never version
independently.

Commit messages must follow Conventional Commits (`type(scope)?: description`, `scope` one of
`core`, `extensions`, or omitted) going forward — `ci.yml`'s `commit-format` job enforces this on
every PR, checking every commit in the PR's range (this repo true-merges, so every commit lands on
`main` and matters to the bump calculation, not just the PR title). Get this wrong and
release-please either miscategorizes a change or silently drops it from the changelog — the CI
guard exists so that's caught at PR time, not discovered in a Release PR that's already wrong.

`release-please.yml` needs its own PAT (not the default `GITHUB_TOKEN`) stored as the
`RELEASE_PLEASE_TOKEN` repository secret — GitHub Actions doesn't let a workflow's default token
trigger other workflows when it pushes, so a release-please-authored tag push would otherwise never
reach `release.yml`. Fine-grained PAT, scoped to this repo only, `Contents: read and write` +
`Pull requests: read and write`. Fine-grained PATs expire — rotate it before it does, or releases
will silently stop triggering.

A manual `git tag v1.2.3 && git push origin v1.2.3` still works as a fallback — `release.yml` only
cares that a `v*` tag arrived, not how — but if you tag manually, **move `[Unreleased]` in
`CHANGELOG.md` under a new `## [X.Y.Z] - YYYY-MM-DD` heading yourself first**; release-please only
writes that entry when it's the one creating the tag. The release workflow greps for that heading
and refuses to publish if it's missing — the same fail-fast treatment as the other premise guards.
```

- [ ] **Step 4: Verify the CLAUDE.md edit landed correctly**

Run: `grep -n "release-please\|RELEASE_PLEASE_TOKEN\|commit-format" CLAUDE.md`
Expected: all three terms appear at least once.

- [ ] **Step 5: Commit**

```bash
git add README.md CLAUDE.md
git commit -m "docs: describe the release-please flow and commit-format requirement"
```

---

## Manual step required (not automatable by any task above)

Before the first push to `main` after this merges, create a fine-grained GitHub PAT scoped to only
the `DocToolkit` repo, with `Contents: read and write` and `Pull requests: read and write`
permissions, and add it as a repository secret named `RELEASE_PLEASE_TOKEN`. Until this exists,
`release-please.yml` runs on every push and fails at the action step with a clear authentication
error — diagnosable, not silent. Once added, validate end-to-end per the design doc's Testing /
validation section: merge a `feat:` commit, confirm a Release PR opens with a correct version bump
and changelog entry, merge that PR, and confirm via `gh run list --workflow=release.yml` that the
tag push actually triggered it — the PAT-vs-default-token distinction is exactly the kind of thing
that fails silently if skipped.
