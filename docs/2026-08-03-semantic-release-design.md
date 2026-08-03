# Semantic-release automation — design

## Why

Every release so far has been fully manual: a human picks a version number, hand-writes a
`CHANGELOG.md` entry, and pushes a `v*` tag. That's deliberate — a publish to nuget.org is
irreversible, so `release.yml` re-proves everything (tests, guards, the CHANGELOG-entry check)
before it pushes. But the bookkeeping *upstream* of that decision — figuring out whether the next
version is a patch or minor bump, and writing the changelog prose — is pure toil with a
well-established automated answer: [Conventional Commits](https://www.conventionalcommits.org/)
plus a tool that computes the bump and drafts the changelog from commit messages.

This design automates that upstream bookkeeping only. It does not touch `release.yml`, the
CHANGELOG-guard, or the "a human decides to publish" principle — it just removes the manual
version-number-and-changelog-prose step that currently sits before that decision.

## Tool: release-please

[`googleapis/release-please-action`](https://github.com/googleapis/release-please-action),
`release-type: "simple"` — verified against release-please's actual schema and a real-world
`release-type: simple` config while writing this design (not assumed from memory, given two
earlier mistakes this session from trusting memory over verification). "simple" tracks the
version as a plain string in `.release-please-manifest.json` and never edits any `.csproj` —
there's no `extra-files` version-marker configured, so `<Version>` in both csprojs stays exactly
what it already is: a local dev default the release workflow overrides at pack time, untouched by
this tool. That separation, already documented in `CLAUDE.md`, is preserved rather than broken.

Both packages are tracked as **one component at the repo root** (`"."`), not two independently
versioned packages — release-please's manifest mode supports multiple independently-versioned
packages, but that's not what we want; we want the existing lockstep-version guarantee, and a
single root component gives it for free. Any qualifying commit anywhere in the repo bumps the one
shared version; `release.yml`'s existing `--skip-duplicate` publish already handles "one package's
content didn't actually change" with zero special-casing.

## Config files

`release-please-config.json`:

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

`include-component-in-tag: false` is required, not optional — it defaults to `true` (verified
against the schema), which would produce tags like `component-v1.2.3` instead of the plain `v1.2.3`
`release.yml` already expects.

`changelog-sections` maps commit types onto this repo's existing Keep-a-Changelog categories
(`Added`/`Fixed`/`Changed`) rather than release-please's own stock section names, and — per the
approved scope — makes `docs`/`ci`/`build`/`refactor` visible under `Changed` rather than hidden,
matching how this repo has already been documenting CI/pipeline changes by hand (e.g. the 0.2.1
entry about merging the release workflows). `chore`/`test` stay hidden: pure maintenance with no
externally-visible effect.

`skip-github-release: true` resolves a real collision found while writing this design:
release-please, by default, creates a GitHub Release object (not just a git tag) when its PR
merges — but `release.yml` *already* creates one (`Create GitHub Release`, with
`--generate-notes`), after tests and every guard have passed. Without this flag, both would try
to create a release for the same tag and the second one (`release.yml`'s) would fail outright, on
every single release. Setting it means release-please only ever handles the Release PR and the tag
push; `release.yml`'s existing release-creation step is untouched and remains the one actually
backed by test/guard evidence.

`.release-please-manifest.json`:

```json
{
  ".": "0.2.1"
}
```

Seeded at the current published version — release-please computes every future bump from this
baseline forward, reading actual commit history since the matching tag.

## Workflow: `.github/workflows/release-please.yml`

```yaml
name: release-please

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

**Requires a Personal Access Token, not the default `GITHUB_TOKEN` — this is the one detail in
this design most likely to silently break if skipped.** GitHub Actions deliberately does not let
a workflow's default `GITHUB_TOKEN` trigger *other* workflows when it pushes a commit or tag (loop
prevention). If release-please pushed the release tag using the default token, `release.yml`'s
`push: tags: ["v*"]` trigger would simply never fire — no error, the tag would just sit there
unpublished. Verified against release-please-action's own documented example, which uses a custom
token for exactly this reason.

**One manual step, once** (same shape as the nuget Trusted Publishing policy and the Codecov
token earlier in this project): create a fine-grained PAT scoped to only this repo, with
`Contents: read and write` and `Pull requests: read and write` permissions, and add it as a
repository secret named `RELEASE_PLEASE_TOKEN`. Fine-grained PATs expire (GitHub caps them at a
maximum lifetime) — this needs periodic rotation, unlike the OIDC-based nuget publishing, which is
a real ongoing cost worth accepting consciously, not discovering when a release silently stops
triggering months from now.

## End-to-end flow

1. PRs merge to `main` as they do today (true merges, not squashes — every individual commit
   lands, which matters below).
2. `release-please.yml` runs on every push to `main`. It reads commits since the last release tag,
   computes the version bump from Conventional Commit types, and opens or updates a single
   persistent "Release PR" — title like `chore(main): release 0.3.0`, body containing the
   generated `CHANGELOG.md` diff.
3. **You review that PR like any other PR** — the generated changelog prose can be hand-edited
   before merging, same as today's manual entries, just starting from a machine-written draft
   instead of a blank page.
4. Merging it is the deliberate publish decision this repo has always required. release-please
   pushes the `vX.Y.Z` tag as part of that merge.
5. `release.yml` fires on that tag push, completely unchanged — full test suite, all three premise
   guards, and the CHANGELOG-guard (already satisfied, since release-please wrote that entry as
   part of the merged PR) — then publishes both packages exactly as it does today.

## Commit message enforcement

Conventional Commits format (`type(scope)?: description`) becomes a real requirement, not a loose
habit, once a tool is inferring version bumps from it. `scope` is one of `core`, `extensions`, or
omitted for repo-wide/tooling changes — matching the `Core:`/`Extensions:`/unprefixed convention
`CHANGELOG.md` already uses for hand-written entries.

Enforced by a new step in `ci.yml`'s existing `pull_request` trigger, checking **every commit in
the PR's range** — not just the PR title — against the format, since this repo true-merges (every
commit lands on `main` individually, and each one matters to release-please's bump calculation, not
just an overall PR summary). A plain bash regex guard, matching this repo's existing guard style
(native-binary/banned-package/SixLabors/CHANGELOG checks are all simple bash, not third-party
actions) — deliberately not `commitlint`, which would pull a Node.js toolchain into a pure .NET
repo's CI for a single regex check:

```
^(feat|fix|docs|style|refactor|perf|test|build|ci|chore|revert)(\([a-z0-9_-]+\))?!?: .+$
```

Type, optional `(scope)`, optional `!` (breaking-change marker), then `: ` and a non-empty
description. The merge commit GitHub itself creates for a merged PR (`Merge pull request #N
from ...`) is exempt — it isn't authored against this convention and isn't one of the commits
release-please parses for the *next* PR's range (it reads the true commits between release tags,
which this repo's merge commits sit alongside but don't replace).

## Explicitly out of scope

- Full auto-publish (every qualifying merge releasing with no human step) — rejected per the
  approved automation-level decision.
- Rewriting any `.csproj`'s `<Version>` — `release-type: simple` with no `extra-files` config
  means release-please never touches those files.
- Retroactively reformatting this repo's existing commit history to Conventional Commits — the
  enforcement guard applies going forward only; release-please's bump calculation starts fresh
  from the `.release-please-manifest.json` baseline (`0.2.1`), not by re-parsing old history.

## Testing / validation

No unit tests apply — this is CI/tooling config. Validation is: after implementation, open a
throwaway PR with a `feat:` commit, merge it, confirm `release-please.yml` actually opens a Release
PR with a correct version bump and changelog entry; merge that PR and confirm the tag gets pushed
and `release.yml` actually fires (not silently skipped) — this last check is the one most likely to
fail silently if the PAT setup is wrong, so it must be observed via `gh run list`, not assumed.
