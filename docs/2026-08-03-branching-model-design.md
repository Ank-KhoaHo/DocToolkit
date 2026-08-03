# Two-branch model: `main` for releases, `develop` for work — design

## Why

Everything in this repo currently lives on one branch. `main` is simultaneously the trunk that
feature branches merge into, the branch release-please watches, the tree that gets tagged and
packed, *and* the landing page a package consumer sees on GitHub. That last role is the one it
serves badly: alongside `src/`, `tests/` and `samples/`, a visitor also lands on `CLAUDE.md`
(21 KB of agent instructions and internal build traps), `docs/` (eleven SDD design docs and
implementation plans), and `spike/` (a dead proof-of-concept explicitly marked "do not modify").

None of that is wrong to *keep* — it is the record of how the package came to exist. It is wrong
to put it in front of someone who arrived from nuget.org wanting to know what
`Ank.DocToolkit` does.

This design splits the two jobs. `develop` becomes the trunk where all work happens and where the
full development record lives. `main` becomes release-only: tags, publishes, and a tree that
contains the shipping library and nothing about the process that produced it.

This **reverses an earlier decision**, recorded in `CLAUDE.md`'s Layout section and
`README.md`'s, that `docs/` and `spike/` are permanent visible parts of the repo. They remain
permanent — on `develop`. They are no longer visible on the branch consumers land on.

## What is excluded from `main`

| Path | Why it is develop-only |
|---|---|
| `CLAUDE.md` | Agent working instructions, internal build traps, maintainer release mechanics. Zero value to a package consumer. |
| `docs/` | SDD design docs and implementation plans — process history, not product documentation. Product docs are `docfx/`, which stays. |
| `spike/` | The original proof-of-concept, kept only as reference and never built. |

Everything else stays on `main`. Three of those are worth stating explicitly, because stripping
them would look superficially reasonable and would break something:

- **`tests/` must stay.** `release.yml` runs `dotnet test` against the tag on `main` before it
  publishes. That is one of the guards standing between a bug and an irreversible nuget.org
  push. Stripping tests from `main` silently removes it.
- **`samples/` must stay.** Both sample projects are in `DocToolkit.sln` and are built by the
  existing CI `dotnet build` step; that build is what catches a breaking API change.
- **`Dockerfile.linux-test` stays**, even though it is a pure dev tool. It is referenced from
  `README.md`, and README must be byte-identical on both branches (see below).

### Verified: stripping these breaks no build

Checked against the actual tree, not assumed:

- `DocToolkit.sln` contains **no reference to `spike/`** — `HtmlPipelineSpike.csproj` is a
  standalone project outside the solution. `dotnet build DocToolkit.sln` on a stripped `main` is
  unaffected.
- `docfx/docfx.json` sources metadata from `../src` only, and content from `docfx/**/*.{md,yml}`
  only. It never reads `docs/` or `spike/`, so `ci.yml`'s `docs-build` job and `docs.yml`'s
  deploy both work on a stripped `main`.
- `.superpowers/` is untracked in this repo (0 files in `git ls-files`), so it needs no exclusion.

## Mechanism: true merge, then scripted strip

`develop` is the real trunk. Promoting to `main` is a **genuine `git merge`**, so every
individual commit from `develop` — with its original Conventional Commit subject — stays
reachable from `main`.

That reachability is the whole reason for choosing a merge over a generated snapshot.
release-please computes its version bump and changelog by walking individual commits reachable
from the branch it watches; `CLAUDE.md` already records that this repo true-merges rather than
squashes for exactly that reason. Preserving a real merge means **release-please keeps watching
`main` with no configuration change at all**. It simply sees a batch of new commits when you
promote, instead of a trickle after every feature merge.

The alternative considered — move release-please to `develop` and regenerate `main` as a filtered
synthetic snapshot at release time — was rejected: it discards per-commit history for source
files on `main` (`git log`/`git blame` there would show only snapshot commits) and requires
rewiring release-please's target branch and the true-merge policy, for no gain over the merge
approach.

### Why the strip must be scripted, not one-time

A one-time deletion commit on `main` is not enough. Once `main` has deleted `docs/` and `develop`
subsequently *modifies* it — which happens every time a new design doc lands — the next merge
produces a **modify/delete conflict** on every affected path. Left to a human, that is a manual
conflict resolution on every single promote, forever.

`scripts/promote-to-main.sh` makes it deterministic by purging the excluded paths
unconditionally, whatever state the merge left them in:

See `scripts/promote-to-main.sh` for the implementation, and
`scripts/test-promote.sh` for the test that proves it. The essential moves are:
fetch, create a throwaway git worktree on a new `release/promote-*` branch based
on `origin/main`, `git merge --no-ff --no-commit origin/develop` tolerating an
exit of exactly 1, purge every excluded path from index and worktree, abort if
any path is *still* unmerged, then commit as `chore: promote develop to main`.

The work happens in a throwaway worktree rather than by switching the current
checkout, because the script lives in `scripts/` and switching the current
worktree to a `main`-based branch could delete the file bash is still reading.
For the same reason `scripts/` is deliberately **not** in the excluded set - it
is release tooling and belongs on `main`.

The script captures `git merge`'s exit code instead of letting `set -e` act on it directly, and
only tolerates exactly 1: a modify/delete conflict (or any other merge conflict) makes `git merge`
exit 1, and that is expected — the purge below is what resolves it, so the script must not abort
first. Anything **above** 1 (128 for unrelated histories, an unreadable object, a merge already in
progress, ...) is not a conflict at all — nothing merged — so the script hard-fails immediately
with that same exit code instead of falling through to the purge and reporting "nothing to
promote" as if it were a clean no-op. The `--diff-filter=U` check afterwards is what keeps a
tolerated exit-1 from swallowing a genuine conflict in `src/` — excluded paths are resolved by the
purge, so anything still unmerged after it is real and stops the script.

The commit subject `chore: promote develop to main` satisfies `ci.yml`'s `commit-format` guard.
A bare `Merge branch ...` subject would fail it, which `CLAUDE.md` already warns about.

### Promotion goes through a PR

The script pushes `release/promote-<timestamp>` and opens a PR into `main`, rather than pushing
to `main` directly, for two reasons:

1. It runs the full CI suite against **main's stripped tree** — the only place that exact
   combination of files is ever exercised. A promote that accidentally strips something the build
   needs fails here, not during a release.
2. It preserves a human gate, consistent with how this repo already treats releases.

**The PR must be merged with a merge commit, never squashed or rebased.** Squashing collapses
every develop commit into one, destroying the Conventional Commit subjects release-please needs
to compute the bump.

## Two failure modes this design must avoid

### Never merge `main` back into `develop`

`main` carries *deletions* of `CLAUDE.md`, `docs/` and `spike/`. A `git merge main` on `develop`
would propagate those deletions and wipe the development record. This is the single most
dangerous operation in this model and must be documented as forbidden in `CLAUDE.md`.

After a release, the two files release-please owns are synced to `develop` by **content copy**,
which carries no deletions:

```bash
git switch develop
git checkout main -- CHANGELOG.md .release-please-manifest.json
git commit -m "chore: sync changelog and manifest from the last release"
```

### `CHANGELOG.md` is main-owned; never edit it on `develop`

release-please writes `CHANGELOG.md` and `.release-please-manifest.json` on `main` only. So long
as `develop` never touches them, git resolves them cleanly on every promote: the merge base has
the old content, `develop` left it alone, `main` changed it — `main` wins, no conflict. The
moment `develop` edits either file, every promote conflicts.

This retires the hand-maintained `## Unreleased` section that `CLAUDE.md` currently describes as
a fallback for manual tags. A manual tag is still possible, but its changelog entry is written on
`main`, not `develop`.

### `README.md` must stay byte-identical on both branches

`README.md` is packed into both `.nupkg`s and its presence is asserted by CI. If it diverged
between branches it would conflict on every promote. So it is fixed **once, on `develop`**: the
Layout tree drops its `spike/` and `docs/` rows, and that information moves into `CLAUDE.md`,
which is develop-only anyway. The `Dockerfile.linux-test` line stays, which is why that file is
not excluded from `main`.

## CI changes

| Workflow | Change |
|---|---|
| `ci.yml` | `push: [main, develop, "feat/**", "fix/**"]`; `pull_request: [main, develop]` |
| `ci.yml` | **New job `branch-policy`**: a PR into `main` must originate from `release/promote-*` or `release-please--branches--main`. Fails with a message pointing at `develop` otherwise. |
| `release-please.yml` | unchanged — still `push: branches: [main]` |
| `release.yml` | unchanged — still `push: tags: ["v*"]` plus `release: published` |
| `docs.yml` | unchanged — still `workflow_run` on Release |

`branch-policy` is what mechanically enforces "main is release-only". Without it, the rule is
convention that a single mis-targeted PR quietly breaks. It is gated on
`if: github.event_name == 'pull_request'` — like the existing `commit-format` job — so it never
runs, and never fails, on a plain push or a `workflow_dispatch`. Its allow-list must include
`release-please--branches--main`, or the Release PR itself would be blocked.

## Hotfixes

There is no separate hotfix branch. An urgent fix branches from `develop` as `fix/*`, merges to
`develop`, and is promoted immediately — the promote script plus a PR merge is fast enough that a
second path into `main` would add a way to bypass CI without adding meaningful speed. The
trade-off is that a hotfix carries along whatever else is already sitting on `develop`; if that is
ever unacceptable, the answer is to promote more often, not to open a side door into `main`.

## Default branch

`main` remains the GitHub default branch. GitHub uses one setting for both the repository landing
page and the default PR base, and they cannot be separated; the landing page is the reason this
design exists, so it wins.

The cost is accepted deliberately: a fresh `git clone` checks out `main`, a tree with no
`CLAUDE.md` and no `docs/`, which is not a tree to develop in. Contributors must
`git switch develop` explicitly, and `README.md` gains a short line saying so. Mis-targeted PRs
are caught by `branch-policy` with an actionable message rather than silently merged.

## Rollout

1. Push the three commits currently unpushed on `main`, so `origin/main` and local agree before
   anything branches.
2. Create `develop` from `main` and push it.
3. On `develop`: add `scripts/promote-to-main.sh`, update `ci.yml` triggers and add
   `branch-policy`, fix `README.md`'s Layout tree, and rewrite `CLAUDE.md`'s Releasing section to
   document the two-branch model, the forbidden back-merge, and main-ownership of `CHANGELOG.md`.
4. Run the promote script once. Its PR produces `main`'s initial strip commit, and CI on that PR
   is the first proof that a stripped `main` builds, tests, packs and builds docs.
5. Retarget the in-flight `di-extensions-parity` worktree onto `develop`.
6. Configure branch protection: `main` requires a PR and green CI; `develop` requires green CI.

Step 4 is the verification step for the whole design — it is where a wrong exclusion surfaces, and
it happens before any tag is pushed, so a mistake there costs a re-run rather than an irreversible
publish.

## What is deliberately not changed

- Release-please's configuration, the version-bump rules, and the Conventional Commits
  requirement.
- `release.yml` in full: the tag-is-the-version rule, all four premise guards, the CHANGELOG
  guard, and Trusted Publishing.
- The lockstep versioning of the core and DI extensions packages.
- `docfx/` and the GitHub Pages deploy chain.
- AutoLnD, which has no CI, no packaging and no release concept, and which this design does not
  touch.
