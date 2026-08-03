# Single-branch model — design

**Supersedes `2026-08-03-branching-model-design.md`**, which introduced the two-branch
`develop`/`main` split this document removes. That design is kept for the record, not as guidance.

## Why

The two-branch model bought one thing: a `main` carrying the shipping library and nothing about the
process that produced it, so a consumer arriving from nuget.org lands on a clean tree. Everything
else it cost. Four of those costs were observed directly, not predicted:

1. **It suppressed Dependabot security updates.** `.github/dependabot.yml` is read from the default
   branch only, so with `main` as default and PRs required to land on `develop`, the config needed
   `target-branch: develop` — and GitHub raises security updates on the default branch only,
   "except where `target-branch` is used". The security-update loss was a pure side effect of the
   branch split.
2. **It hid the contributor documentation from contributors.** `CLAUDE.md` — which holds the entire
   branching model, the commit-format requirement, and every trap in the codebase — was stripped
   from `main`, the branch every newcomer lands on. Backlog item D1 exists only because of this.
3. **It forced `README.md` to stay byte-identical on both branches** forever, or every promote
   conflicts.
4. **It created the repo's most dangerous operation.** `CLAUDE.md` describes "never merge `main`
   into `develop`" as "the single most dangerous operation in this repo" — a hazard that exists
   purely because `main` carries deletions. Remove the split and the hazard stops existing.

Plus ongoing maintenance: a promote script, a test for the promote script, and a CI job to run that
test, none of which serve the library.

The benefit was cosmetic and a good README delivers it.

## The target model

1. One branch: `main`. `develop` is deleted.
2. `main` is the default branch (already true).
3. `main` cannot be pushed directly — every change arrives by pull request.
4. Every merge to `main` produces a release. No human gate.
5. `CLAUDE.md`, `docs/` and `spike/` live on `main` like everything else.

## What is deleted

- `scripts/promote-to-main.sh` and `scripts/test-promote.sh` — the whole `scripts/` directory.
- `ci.yml`'s `branch-policy` job. It rejects PRs into `main` that are not promote or release-please
  branches. Under a single-branch model every PR targets `main`, so the job's entire purpose is
  gone.
- `ci.yml`'s `promote-script` job, whose subject no longer exists.

## What changes

**`ci.yml` triggers** drop `develop` from both `push` and `pull_request`.

**`.github/dependabot.yml`** drops `target-branch: develop` from all five update blocks. PRs then
target the default branch, which is `main`. **Dependabot security updates return** — see cost 1
above. This is the clearest evidence the collapse is the right call: a capability lost to the
branching model comes back by deleting the branching model, not by adding anything.

**`CLAUDE.md`** rewrites *Branches* and *Releasing*, and deletes outright:
- "Never merge `main` into `develop`"
- the `CHANGELOG.md` / `.release-please-manifest.json` main-ownership rules and the content-copy
  sync procedure
- the README byte-identical constraint
- the `(develop only)` annotations in *Layout*

**`README.md`** drops the "development happens on `develop`" contributor note and corrects
"Maintainer procedure lives in `CLAUDE.md` on `develop`".

`CLAUDE.md`, `docs/` and `spike/` reach `main` because nothing strips them any more, not because
anything is added.

## The release mechanism

Requirement 4 (every merge releases) collides with requirement 3 (`main` is PR-only): auto-release
normally means a bot writing a changelog and tag back to `main`, which is a direct push.

**Resolution: keep release-please, and auto-merge its Release PR.** One step appended to
`release-please.yml`:

```yaml
      - name: Auto-merge the Release PR
        if: steps.release.outputs.pr
        env:
          GH_TOKEN: ${{ secrets.RELEASE_PLEASE_TOKEN }}
        run: |
          number=$(echo '${{ steps.release.outputs.pr }}' | jq -r '.number')
          gh pr merge "$number" --merge --auto
```

Flow: merge any PR → release-please opens or updates the Release PR with the computed bump and
changelog entry → auto-merge fires once checks pass → release-please's next run creates the tag and
GitHub Release → `release.yml` packs, re-proves every guard, and publishes.

The bot merges a **pull request**; it never pushes. Branch protection holds with no bypass actor,
so requirement 3 stays literally true rather than nominally true.

No loop: release-please ignores its own release commits, so the Release PR merging produces no
further Release PR.

`release.yml` needs no change, including its `CHANGELOG.md`-entry guard — release-please still
writes that entry before tagging.

### Repository settings this requires

Not expressible in files. All three must be configured by hand:

1. Branch protection on `main`: require a pull request, require status checks to pass.
2. Repository setting **Allow auto-merge** enabled.
3. **Do not require approving reviews on `main`**, or add release-please's identity as a bypass
   actor.

Item 3 is a trap worth stating plainly: a required-approvals rule can never be satisfied by an
unattended bot, so the Release PR would sit open indefinitely and **releases would stop with no
error anywhere**. It fails silently, in the direction of nothing happening.

## Rejected alternatives

**Tag on push, drop release-please.** A workflow on `push: main` computes the version and pushes a
tag; simplest runtime story, one merge to one publish. Rejected because writing `CHANGELOG.md` back
to `main` would be a direct push, so the committed changelog would have to be dropped in favour of
GitHub Release notes — losing a real artifact that `README.md` links and that `release.yml` guards.
Simplicity paid for with a deletion.

**Bot pushes directly using a bypass PAT.** Fewest moving parts, keeps everything. Rejected because
it makes requirement 3 false: a long-lived token holds open a direct-push path to `main`, in a repo
whose security posture is Trusted Publishing precisely so that no long-lived credential exists.

**Keeping the two-branch model.** Rejected on the four observed costs above.

## Migration

A plain `git merge develop` into `main` does **not** restore `CLAUDE.md`, `spike/` or the older
`docs/` files. `main` carries commits deleting them, `develop` has not modified them since the merge
base, and git resolves that combination by keeping the deletion. The restore must be explicit.

Verified before writing this: `CHANGELOG.md` and `.release-please-manifest.json` do not appear in
the `main`↔`develop` diff, so they are already identical and the migration has no conflict there.

1. Branch from `main`, named `release/promote-single-branch` — matching `release/promote-*` so the
   still-live `branch-policy` guard admits the PR regardless of whether the job runs.
2. `git merge origin/develop`, carrying all ten commits **with their subjects intact**, which
   release-please needs to compute the bump.
3. `git checkout origin/develop -- CLAUDE.md docs/ spike/` and commit — the explicit restore.
4. Apply the *What is deleted* and *What changes* edits above.
5. **Scan for secrets before opening the PR.** `docs/` and `spike/` are about to become
   world-readable and permanent; roughly 10,500 lines arrive at once. The workspace `CLAUDE.md` is
   explicit that publishing is a one-way door.
6. Open the PR against `main` and merge it with a **merge commit, never a squash**.
7. Delete `develop`, local and remote, and the merged `feat/dependency-automation`.
8. **Then** configure the three repository settings.

### Why the settings come last

If auto-merge is live before the migration lands, the migration PR itself publishes a release
unattended — a large, unusual change would be the first thing to ship through a pipeline nobody has
yet watched work. Landing it under the current manual gate, verifying, then enabling auto-release
makes the *first* auto-release a small deliberate test instead.

## Consequences accepted

- **Every merge to `main` publishes both packages to nuget.org, irreversibly.** A version can be
  unlisted, never deleted or replaced. This includes documentation-only merges, CI tweaks and
  test-only changes, each consuming a version number and producing a release whose changelog entry
  may be empty. Chosen deliberately after the alternative — restricting auto-release to merges that
  touch `src/` — was raised and declined.
- **Version numbers will move quickly.** Dependabot's weekly PRs each become a release once merged.
- **Backlog item C11 (the empty-release guard) is now moot rather than pending.** Publishing an
  empty release is the specified behaviour, not a hazard to guard against.
- **`docs/` becomes public**, including `2026-08-03-enhancement-backlog.md`, which catalogues 56
  open weaknesses in this project in candid terms.

## Success criteria

- No `develop` branch exists; every workflow and document refers only to `main`.
- A PR merged to `main` results in a published version with no human action after the merge.
- `main` rejects a direct push.
- `CLAUDE.md`, `docs/` and `spike/` are present on `main`.
- `.github/dependabot.yml` contains no `target-branch`, and Dependabot security updates are active.
- `dotnet build -warnaserror` is clean and all 448 test results pass on `main` after migration.

## Verified 2026-08-03

All criteria met, with one deliberate exception recorded below.

Merging PR #15 produced a release with no further human action: release-please opened
`chore(main): release 0.3.3`, auto-merge took it, `v0.3.3` was tagged, and `release.yml` published
both packages and their symbol packages to nuget.org. The lockfile guard added the same day ran
inside that release and passed, so the resolved graph was checked against its committed lockfile
immediately before an irreversible publish.

Dependabot began work immediately, opening `ci: bump the actions group with 8 updates` (grouped as
configured) and `build: Bump SixLabors.Fonts from 1.0.0 to 1.0.1`. The second is the `ignore` rule
behaving exactly as designed — 1.0.1 was checked and is still Apache-2.0, so a safe patch was
admitted while the revenue-gated 2.x line stays blocked.

**Exception: `main` does not reject a direct push from a repository admin.** `enforce_admins` is
deliberately left `false`, so branch protection binds contributors while the owner retains an
emergency override. Requirement 3 therefore holds by enforcement for everyone else and by
convention for the owner. This was decided explicitly, not inherited — see the implementation
plan's Task 7 for the alternative and its cost.

### Three things the implementation corrected

1. **`main`'s required status checks named `main is release-only`** — the `branch-policy` job the
   migration deletes. Had the PR merged unchanged, every future PR into `main` would have waited
   forever on a check nothing could report. The list was replaced with the seven jobs that survive,
   which also closed a pre-existing gap: the Windows build, the docs site and the extensions
   package verification had never gated `main` at all.
2. **`git merge --no-edit` produces a subject with no Conventional Commit type,** which
   `commit-format` correctly rejects — it exempts GitHub's own PR-merge commits but deliberately
   not hand-made merges. The migration branch was rebuilt with an explicit `chore:` merge message,
   and the rebuilt tree verified byte-identical to the tested one before force-pushing.
3. **The `Auto-merge the Release PR` step failed loudly** on the first run, with
   `GraphQL: Auto merge is not allowed for this repository`, because *Allow auto-merge* was still
   off. This was the failure mode of most concern in the design, and it behaved correctly: visible,
   not silent.
