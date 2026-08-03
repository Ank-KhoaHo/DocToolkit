# Two-branch model (`main` = releases, `develop` = work) — implementation plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Split DocToolkit into a release-only `main` that contains no development-phase
material, and a `develop` trunk that contains everything, without changing how release-please,
`release.yml` or the nuget.org publish work.

**Architecture:** `develop` is the real trunk. Promotion to `main` is a genuine `git merge`
performed in a throwaway git worktree, followed by an unconditional purge of the excluded paths,
raised as a PR into `main`. Because it is a real merge, every Conventional Commit subject from
`develop` stays reachable from `main`, so release-please keeps watching `main` with zero
configuration change.

**Tech Stack:** git 2.53, bash, GitHub Actions, `gh` CLI 2.96 (already authenticated as
`Ank-KhoaHo`), release-please v4.

**Design doc:** `docs/2026-08-03-branching-model-design.md`

**Every command in this plan runs from the repository root**, `e:/PJ/LnDPrj/DocToolkit`, in bash
(Git Bash on Windows). Paths are relative to it.

## Global Constraints

- **Excluded from `main` (develop-only): `CLAUDE.md`, `docs/`, `spike/`.** Exactly these three.
- **`tests/`, `samples/`, `.github/`, `.gitattributes`, `CHANGELOG.md`, `docfx/`,
  `Dockerfile.linux-test`, `release-please-config.json`, `.release-please-manifest.json` all stay
  on `main`.** `release.yml` runs `dotnet test` from the tag on `main` before publishing;
  stripping `tests/` would silently remove that guard.
- **`scripts/` stays on `main`** — it is release tooling, and excluding it would let the promote
  script delete itself out from under the running bash process.
- **Never `git merge main` into `develop`.** `main` carries deletions; merging it back wipes the
  development record. Sync only by content copy: `git checkout main -- <file>`.
- **`CHANGELOG.md` and `.release-please-manifest.json` are main-owned.** Never edit them on
  `develop`.
- **`README.md` must stay byte-identical on both branches** — it is packed into both `.nupkg`s
  and asserted by CI; divergence would conflict on every promote.
- **Promote PRs must be merged with a merge commit, never squashed.** Squashing destroys the
  Conventional Commit subjects release-please needs.
- **Commit messages follow Conventional Commits** (`type(scope)?: description`) and must **never**
  contain a `Co-Authored-By` trailer.
- Do not touch `release.yml`, `release-please.yml`, `docs.yml`, `release-please-config.json`, or
  any `.csproj`.

---

### Task 1: Push pending work and create `develop`

`main` is currently 4 commits ahead of `origin/main`. Everything downstream branches from `main`,
so local and remote must agree first.

**Files:** none — git refs only.

**Interfaces:**
- Produces: remote branch `origin/develop`, identical to `origin/main`. Every later task commits
  to `develop`.

- [ ] **Step 1: Confirm the working tree is clean and see what is unpushed**

```bash
git status --short
git log --oneline origin/main..main
```

Expected: `status --short` prints nothing. `log` prints exactly 4 commits, ending with
`docs: design a two-branch model separating releases from development`.

If `status --short` prints anything, stop and resolve it before continuing.

- [ ] **Step 2: Push `main`**

```bash
git push origin main
```

Expected: push succeeds. Re-running `git log --oneline origin/main..main` now prints nothing.

- [ ] **Step 3: Create and push `develop`**

```bash
git switch -c develop main
git push -u origin develop
```

Expected: `Branch 'develop' set up to track 'origin/develop'`.

- [ ] **Step 4: Verify both branches point at the same commit**

```bash
git rev-parse origin/main origin/develop
```

Expected: two identical SHAs.

- [ ] **Step 5: No commit**

This task creates refs only; there is nothing to commit.

---

### Task 2: The promote script and its test

The riskiest piece of the whole design. Its merge/purge/conflict logic gets a real test that
builds a throwaway repository, so a regression is caught here rather than on a release branch.

**Files:**
- Create: `scripts/promote-to-main.sh`
- Create: `scripts/test-promote.sh`
- Modify: `.gitattributes` (add an `*.sh text eol=lf` rule)
- Modify: `docs/2026-08-03-branching-model-design.md` (correct the script sketch)

**Interfaces:**
- Produces: `scripts/promote-to-main.sh`, honouring env overrides `PROMOTE_REMOTE`
  (default `origin`), `PROMOTE_PUSH` (default `1`), `PROMOTE_OPEN_PR` (default `1`). Creates a
  branch named `release/promote-<YYYYMMDD-HHMMSS>`. Exits `0` on success, `1` on a real conflict.
- Produces: `scripts/test-promote.sh`, run by CI in Task 3.

- [ ] **Step 1: Add the line-ending rule first**

`.gitattributes` currently sets `* text=auto`, which gives `.sh` files CRLF on Windows checkout.
CI runs these scripts on `ubuntu-latest`, where a `\r` produces
`bash: line 2: $'\r': command not found`. Append to `.gitattributes`:

```
# Shell scripts must keep LF even when checked out on Windows: `* text=auto`
# above would give them CRLF, and bash on Linux then fails with
# `$'\r': command not found`. scripts/ runs in CI on ubuntu-latest.
*.sh             text eol=lf
```

- [ ] **Step 2: Write the failing test**

Create `scripts/test-promote.sh`:

```bash
#!/usr/bin/env bash
# Proves the merge-and-strip logic in promote-to-main.sh.
#
# Builds a throwaway repository with a local bare "origin" under $TMPDIR, so this
# never touches the real repository or its remote. Three properties matter:
#   1. excluded paths are stripped while source changes carry over;
#   2. a modify/delete conflict on an excluded path resolves itself (this is the
#      one that recurs on every promote once main has deleted docs/);
#   3. a real conflict in shipping code stops the promote instead of being
#      silently swallowed by the purge.
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROMOTE="$SCRIPT_DIR/promote-to-main.sh"

TMP="$(mktemp -d)"
trap 'cd /; chmod -R u+w "$TMP" 2>/dev/null || true; rm -rf "$TMP"' EXIT

pass=0
fail=0
check() { # check <description> <0-for-pass>
  if [ "$2" -eq 0 ]; then
    echo "  ok   - $1"; pass=$((pass + 1))
  else
    echo "  FAIL - $1"; fail=$((fail + 1))
  fi
}
assert_has()    { if grep -qx "$1" "$TMP/tree.txt"; then check "$1 present" 0; else check "$1 present" 1; fi; }
assert_absent() { if grep -qx "$1" "$TMP/tree.txt"; then check "$1 absent"  1; else check "$1 absent"  0; fi; }

tree_of() { git -C "$TMP/work" ls-tree -r --name-only "$1" > "$TMP/tree.txt"; }

promoted_branch() {
  git -C "$TMP/work" for-each-ref --format='%(refname:short)' \
    --sort=-committerdate 'refs/heads/release/promote-*' | head -1
}

setup_repo() {
  cd /
  rm -rf "$TMP/origin.git" "$TMP/work"
  git init -q --bare "$TMP/origin.git"
  git clone -q "$TMP/origin.git" "$TMP/work"
  cd "$TMP/work"
  git config user.email test@example.com
  git config user.name "Test"
  git config commit.gpgsign false
  # Windows checkouts default to autocrlf=true, which floods the run with
  # "LF will be replaced by CRLF" warnings and buries a real failure.
  git config core.autocrlf false
  mkdir -p src docs spike
  echo v1 > src/app.txt
  echo agent-instructions > CLAUDE.md
  echo design > docs/design.md
  echo poc > spike/poc.txt
  git add -A
  git commit -qm "chore: seed"
  git branch -M main
  git push -q -u origin main
  git switch -qc develop
  git push -q -u origin develop
}

run_promote() { # run_promote -> echoes exit code, never aborts the test run
  set +e
  ( cd "$TMP/work" && PROMOTE_PUSH=0 PROMOTE_OPEN_PR=0 bash "$PROMOTE" ) >/dev/null 2>&1
  local rc=$?
  set -e
  echo "$rc"
}

echo "test 1: excluded paths are stripped, source changes carry over"
setup_repo
echo v2 > src/app.txt
echo more > docs/new-design.md
git add -A
git commit -qm "feat: bump app to v2"
git push -q origin develop
rc=$(run_promote)
check "promote exits 0" "$rc"
BR="$(promoted_branch)"
if [ -n "$BR" ]; then check "a release/promote-* branch was created" 0; else check "a release/promote-* branch was created" 1; fi
tree_of "$BR"
assert_has    "src/app.txt"
assert_absent "CLAUDE.md"
assert_absent "docs/design.md"
assert_absent "docs/new-design.md"
assert_absent "spike/poc.txt"
if [ "$(git -C "$TMP/work" show "$BR:src/app.txt")" = "v2" ]; then
  check "src/app.txt carries develop's change" 0
else
  check "src/app.txt carries develop's change" 1
fi
if git -C "$TMP/work" rev-list "$BR" | grep -qx "$(git -C "$TMP/work" rev-parse develop)"; then
  check "develop's commits stay reachable (true merge, not a snapshot)" 0
else
  check "develop's commits stay reachable (true merge, not a snapshot)" 1
fi

echo "test 2: a modify/delete conflict on an excluded path resolves itself"
git -C "$TMP/work" switch -q main
git -C "$TMP/work" merge -q --no-ff -m "chore: land promote" "$BR"
git -C "$TMP/work" push -q origin main
git -C "$TMP/work" switch -q develop
echo edited-after-strip > "$TMP/work/docs/design.md"
echo v3 > "$TMP/work/src/app.txt"
git -C "$TMP/work" add -A
git -C "$TMP/work" commit -qm "docs: edit a design doc main has already deleted"
git -C "$TMP/work" push -q origin develop
rc=$(run_promote)
check "second promote exits 0 despite the modify/delete conflict" "$rc"
BR2="$(promoted_branch)"
tree_of "$BR2"
assert_absent "docs/design.md"
assert_has    "src/app.txt"

echo "test 3: a real conflict in shipping code stops the promote"
setup_repo
# setup_repo leaves us on develop; the main-side change has to land on main.
git switch -q main
echo main-side > src/app.txt
git add -A
git commit -qm "fix: main-side change"
git push -q origin main
git switch -q develop
echo develop-side > src/app.txt
git add -A
git commit -qm "fix: develop-side change"
git push -q origin develop
rc=$(run_promote)
if [ "$rc" -ne 0 ]; then check "promote exits non-zero on a real conflict" 0; else check "promote exits non-zero on a real conflict" 1; fi

echo
echo "passed: $pass   failed: $fail"
[ "$fail" -eq 0 ]
```

- [ ] **Step 3: Run the test to verify it fails**

```bash
bash scripts/test-promote.sh
```

Expected: FAIL. The script aborts because `scripts/promote-to-main.sh` does not exist yet —
`run_promote` returns non-zero, so `check "promote exits 0"` reports `FAIL`, and the final
`[ "$fail" -eq 0 ]` makes the script exit non-zero.

- [ ] **Step 4: Write the promote script**

Create `scripts/promote-to-main.sh`:

```bash
#!/usr/bin/env bash
# Promote develop to main.
#
# main is release-only: it must not carry the development record (CLAUDE.md,
# docs/, spike/). Promotion is a REAL git merge, so every Conventional Commit
# subject from develop stays reachable from main - that is what release-please
# walks to compute the version bump. A generated snapshot would break it.
#
# The excluded paths are purged UNCONDITIONALLY after the merge rather than
# deleted once, because as soon as main has deleted docs/ and develop later
# modifies it, every subsequent merge raises a modify/delete conflict. Purging
# makes the resolution deterministic: main's side (absent) always wins.
#
# The work happens in a throwaway git worktree rather than by switching the
# current checkout, because this script lives in scripts/ - switching the
# current worktree to a main-based branch could remove the file bash is still
# reading.
#
# Env overrides (used by scripts/test-promote.sh):
#   PROMOTE_REMOTE   remote name             (default: origin)
#   PROMOTE_PUSH     1 = push the branch     (default: 1)
#   PROMOTE_OPEN_PR  1 = open a PR via gh    (default: 1)
set -euo pipefail

EXCLUDED=(CLAUDE.md docs spike)
REMOTE="${PROMOTE_REMOTE:-origin}"
DO_PUSH="${PROMOTE_PUSH:-1}"
DO_PR="${PROMOTE_OPEN_PR:-1}"

ORIG_DIR="$PWD"
WT="$(mktemp -d)"
DELETE_BRANCH=0

# Two promotes inside the same second would otherwise collide on the branch name
# and `git worktree add -b` would fail. Unlikely by hand; certain in the test.
BRANCH="release/promote-$(date +%Y%m%d-%H%M%S)"
suffix=1
while git show-ref --quiet --verify "refs/heads/$BRANCH"; do
  BRANCH="release/promote-$(date +%Y%m%d-%H%M%S)-$suffix"
  suffix=$((suffix + 1))
done

# cd out of the worktree BEFORE removing it - `git worktree remove` fails while
# the shell's cwd is inside the directory it is deleting.
cleanup() {
  cd "$ORIG_DIR" 2>/dev/null || cd /
  git worktree remove --force "$WT" >/dev/null 2>&1 || true
  rm -rf "$WT" 2>/dev/null || true
  git worktree prune >/dev/null 2>&1 || true
  if [ "$DELETE_BRANCH" = "1" ]; then
    git branch -D "$BRANCH" >/dev/null 2>&1 || true
  fi
}
trap cleanup EXIT

git fetch --quiet "$REMOTE" main develop
git worktree add -q -b "$BRANCH" "$WT" "$REMOTE/main"
cd "$WT"

# A modify/delete conflict on an excluded path makes `git merge` exit non-zero.
# That is expected and is resolved by the purge below, so under `set -e` it must
# not abort the script. Genuine conflicts are caught by the --diff-filter=U
# check afterwards.
git merge --no-ff --no-commit "$REMOTE/develop" || true

for p in "${EXCLUDED[@]}"; do
  git rm -r -f -q --cached --ignore-unmatch -- "$p" >/dev/null 2>&1 || true
  rm -rf -- "$p"
done

if git diff --name-only --diff-filter=U | grep -q .; then
  echo "error: conflict outside the excluded set - resolve it on develop first:" >&2
  git diff --name-only --diff-filter=U >&2
  exit 1
fi

if ! git rev-parse -q --verify MERGE_HEAD >/dev/null && git diff --cached --quiet; then
  echo "nothing to promote - main already matches develop"
  DELETE_BRANCH=1   # the trap removes the worktree first, then the empty branch
  exit 0
fi

git commit -q -m "chore: promote develop to main"
echo "created $BRANCH"

if [ "$DO_PUSH" = "1" ]; then
  git push -q -u "$REMOTE" "$BRANCH"
  echo "pushed $BRANCH to $REMOTE"
fi

if [ "$DO_PR" = "1" ]; then
  gh pr create --base main --head "$BRANCH" \
    --title "chore: promote develop to main" \
    --body "Promotes \`develop\` to \`main\`.

Excluded from \`main\`: ${EXCLUDED[*]}

**Merge this with a merge commit - never squash.** Squashing collapses every
\`develop\` commit into one and destroys the Conventional Commit subjects
release-please needs to compute the version bump."
  echo "opened PR"
fi
```

- [ ] **Step 5: Run the test to verify it passes**

```bash
bash scripts/test-promote.sh
```

Expected output ends with:

```
passed: 13   failed: 0
```

and exit code 0. (This exact suite was run against these exact two scripts while the plan was
written: 13 passed, 0 failed, and the test-3 failure path was confirmed to abort with
`error: conflict outside the excluded set` naming `src/app.txt` — not to exit non-zero for some
unrelated reason.)

If any line reports `FAIL`, fix the script — do not adjust the test to match.

- [ ] **Step 6: Correct the script sketch in the design doc**

The design doc sketches the promote script using `git switch -c ... origin/main`. That is wrong —
it would switch the current checkout to a `main`-based tree, which on the first promote does not
contain `scripts/`, deleting the script bash is reading. Replace the fenced `bash` block in the
"Why the strip must be scripted, not one-time" section of
`docs/2026-08-03-branching-model-design.md` with a pointer to the real file, and add the
`scripts/` rationale.

Replace the whole fenced code block with:

```
See `scripts/promote-to-main.sh` for the implementation, and
`scripts/test-promote.sh` for the test that proves it. The essential moves are:
fetch, create a throwaway git worktree on a new `release/promote-*` branch based
on `origin/main`, `git merge --no-ff --no-commit origin/develop` tolerating a
non-zero exit, purge every excluded path from index and worktree, abort if any
path is *still* unmerged, then commit as `chore: promote develop to main`.

The work happens in a throwaway worktree rather than by switching the current
checkout, because the script lives in `scripts/` and switching the current
worktree to a `main`-based branch could delete the file bash is still reading.
For the same reason `scripts/` is deliberately **not** in the excluded set - it
is release tooling and belongs on `main`.
```

- [ ] **Step 7: Commit, marking both scripts executable**

`git update-index --chmod=+x` must run *after* `git add`, and is what makes the scripts
executable when CI checks them out on Linux — Windows does not track the on-disk permission bit.

```bash
git add .gitattributes scripts/promote-to-main.sh scripts/test-promote.sh docs/2026-08-03-branching-model-design.md
git update-index --chmod=+x scripts/promote-to-main.sh scripts/test-promote.sh
git commit -F - <<'EOF'
ci: add the develop-to-main promote script and its test

Promotion is a real git merge plus an unconditional purge of CLAUDE.md, docs/
and spike/. The purge has to be unconditional rather than a one-time deletion:
once main has deleted docs/ and develop later edits it, every subsequent merge
raises a modify/delete conflict, and resolving that by hand on every promote is
exactly the toil this is meant to remove.

test-promote.sh builds a throwaway repo with a local bare origin and proves the
three properties that matter - excluded paths stripped while source carries
over, the recurring modify/delete conflict self-resolving, and a real conflict
in shipping code stopping the promote rather than being swallowed by the purge.

*.sh is pinned to eol=lf because `* text=auto` would otherwise hand CRLF to
bash on ubuntu-latest.
EOF
```

---

### Task 3: CI — triggers, branch policy, and the merge-commit exemption

**Files:**
- Modify: `.github/workflows/ci.yml` — trigger lists (lines 9-14), `commit-format` job
  (lines 150-177), plus two new jobs.

**Interfaces:**
- Consumes: `scripts/test-promote.sh` from Task 2.
- Produces: a `branch-policy` job that rejects any PR into `main` not from
  `release/promote-*` or `release-please--branches--main`. Task 6's promote PR depends on being
  allow-listed by it.

- [ ] **Step 1: Widen the triggers**

In `.github/workflows/ci.yml`, replace:

```yaml
on:
  push:
    branches: [main, "feat/**", "fix/**"]
  pull_request:
    branches: [main]
  workflow_dispatch:
```

with:

```yaml
on:
  push:
    branches: [main, develop, "feat/**", "fix/**"]
  pull_request:
    branches: [main, develop]
  workflow_dispatch:
```

- [ ] **Step 2: Exempt GitHub's own PR-merge commits from `commit-format`**

This is required, not cosmetic. `commit-format` checks every commit in the PR's range. A promote
PR's range includes the merge commits GitHub created when feature PRs landed on `develop`, whose
subjects read `Merge pull request #12 from Ank-KhoaHo/feat/x` — which does not match the
Conventional Commits regex. Without this exemption the first promote PR in Task 6 fails CI.

The exemption is deliberately narrow: it matches only GitHub's own generated subject, so a
hand-made `Merge branch 'develop' into feat/x` still fails, preserving the guard `CLAUDE.md`
already documents.

In the `commit-format` job, replace the `while` loop body:

```bash
          while IFS= read -r line; do
            sha="${line%% *}"
            subject="${line#* }"
            if ! [[ "$subject" =~ $regex ]]; then
              echo "::error::Commit $sha does not match Conventional Commits format (type(scope)?: description): \"$subject\""
              bad=1
            fi
          done < <(git log --format='%H %s' "$base..$head")
```

with:

```bash
          while IFS= read -r line; do
            sha="${line%% *}"
            subject="${line#* }"
            # GitHub's own PR-merge commits are generated, never hand-written, and
            # a promote PR's range is full of them (every feature PR that landed on
            # develop). A hand-made "Merge branch 'develop' into feat/x" is NOT
            # exempt - that one still fails, which is the guard we want to keep.
            if [[ "$subject" =~ ^Merge\ pull\ request\ \#[0-9]+\ from\  ]]; then
              continue
            fi
            if ! [[ "$subject" =~ $regex ]]; then
              echo "::error::Commit $sha does not match Conventional Commits format (type(scope)?: description): \"$subject\""
              bad=1
            fi
          done < <(git log --format='%H %s' "$base..$head")
```

- [ ] **Step 3: Add the `branch-policy` job**

Append to `.github/workflows/ci.yml`:

```yaml
  # main is release-only. Without this the rule is convention that a single
  # mis-targeted PR quietly breaks. Gated on pull_request like commit-format, so
  # it never runs on a plain push or a workflow_dispatch.
  branch-policy:
    name: main is release-only
    runs-on: ubuntu-latest
    if: github.event_name == 'pull_request' && github.base_ref == 'main'
    steps:
      - name: Assert this PR may target main
        run: |
          head="${{ github.head_ref }}"
          case "$head" in
            release/promote-*|release-please--branches--main)
              echo "ok - '$head' may target main" ;;
            *)
              echo "::error::main is release-only - retarget this PR at 'develop'. Only a promote branch (scripts/promote-to-main.sh, which creates release/promote-*) and release-please's own Release PR may target main; '$head' may not."
              exit 1 ;;
          esac
```

- [ ] **Step 4: Add the `promote-script` job**

Append to `.github/workflows/ci.yml`:

```yaml
  # The promote script is the only thing standing between develop and a main that
  # still carries the development record. Re-prove its merge/purge/conflict logic
  # on every push, the same way the premise guards re-prove the dependency graph.
  promote-script:
    name: promote script logic
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4

      - name: Test the merge-and-strip logic
        run: bash scripts/test-promote.sh
```

- [ ] **Step 5: Verify the YAML is well-formed and the edits landed**

```bash
grep -c 'branches: \[main, develop, "feat/\*\*", "fix/\*\*"\]' .github/workflows/ci.yml
grep -c 'branches: \[main, develop\]'                          .github/workflows/ci.yml
grep -c '^  branch-policy:'                                    .github/workflows/ci.yml
grep -c '^  promote-script:'                                   .github/workflows/ci.yml
grep -c "PR-merge commits are generated"                       .github/workflows/ci.yml
```

Expected: `1` five times.

Then confirm the YAML still parses (GitHub will not tell you until you push):

```bash
python3 -c "import yaml,sys; yaml.safe_load(open('.github/workflows/ci.yml')); print('yaml ok')"
```

Expected: `yaml ok`. If `python3` has no `yaml` module, skip this — Step 7 catches a broken file
when the run either appears or does not.

- [ ] **Step 6: Commit and push to `develop` to see CI run**

```bash
git add .github/workflows/ci.yml
git commit -F - <<'EOF'
ci: run on develop, and keep main release-only

Widens the triggers to develop now that it is the trunk, and adds branch-policy,
which rejects any PR into main that is not a promote branch or release-please's
own Release PR - the rule is otherwise convention that one mis-targeted PR
quietly breaks.

commit-format now skips GitHub's own generated PR-merge subjects. A promote PR's
range includes every "Merge pull request #N from ..." commit created when feature
PRs landed on develop, and those do not match the Conventional Commits regex, so
without this the first promote PR fails. A hand-made "Merge branch ..." still
fails, which is the guard worth keeping.
EOF
git push origin develop
```

- [ ] **Step 7: Confirm CI actually ran on `develop`**

```bash
gh run list --branch develop --limit 3
```

Expected: a `CI` run for the commit just pushed. Wait for it and confirm the
`promote script logic` job passed:

```bash
gh run watch
```

Expected: all jobs green. `branch-policy` is *skipped* (this is a push, not a PR) — that is
correct, not a failure.

---

### Task 4: README — remove the dev-phase rows, add a contributor note

`README.md` ships inside both `.nupkg`s and must stay byte-identical on both branches, so it is
edited once here on `develop` and reaches `main` via the promote in Task 6.

**Files:**
- Modify: `README.md` — Layout block (lines 257-258), and a new note after the badges.

- [ ] **Step 1: Remove the two dev-phase rows from the Layout block**

Delete these two lines from the fenced Layout block:

```
spike/                                                  the original proof-of-concept, kept as reference
docs/                                                   the implementation plan this was built from
```

The block's last remaining line is the `docfx/` row. Leave every other row untouched.

- [ ] **Step 2: Add the contributor note**

Immediately after the line `📖 [API documentation](https://ank-khoaho.github.io/DocToolkit/)`,
insert a blank line and then:

```markdown
> **Contributing?** `main` is release-only. Development happens on **`develop`** — run
> `git switch develop` after cloning, and target pull requests at `develop`, not `main`.
```

- [ ] **Step 3: Verify the edits**

```bash
grep -nE '^(spike|docs)/' README.md; echo "exit=$?"
grep -c 'git switch develop' README.md
```

Expected: the first `grep` prints no matching lines and `exit=1`. The second prints `1`.

- [ ] **Step 4: Confirm the file still names everything that stays on main**

```bash
for p in src/DocToolkit/ tests/DocToolkit.Tests/ samples/ConsoleSample/ docfx/ Dockerfile.linux-test; do
  grep -q "$p" README.md && echo "ok   $p" || echo "MISS $p"
done
```

Expected: all five print `ok`.

- [ ] **Step 5: Commit**

```bash
git add README.md
git commit -F - <<'EOF'
docs: point contributors at develop and drop the dev-phase rows

README ships inside both nupkgs and has to stay byte-identical on develop and
main, or it conflicts on every promote. So the spike/ and docs/ rows go away
here rather than only on main - that information now lives in CLAUDE.md, which
is develop-only anyway.
EOF
git push origin develop
```

---

### Task 5: CLAUDE.md — document the two-branch model

`CLAUDE.md` is itself develop-only, so it is the right home for everything a contributor or agent
needs and a consumer does not.

**Files:**
- Modify: `CLAUDE.md` — new "Branches" section before "Releasing" (currently line 209), edits
  inside "Releasing", and the Layout block (lines 302-314).

- [ ] **Step 1: Insert the Branches section immediately before the `## Releasing` heading**

```markdown
## Branches

Two branches, two jobs.

- **`develop` is the trunk.** All work merges here. It carries everything, including
  `CLAUDE.md`, `docs/` and `spike/`.
- **`main` is release-only.** It carries the shipping library and nothing about the process that
  produced it: `CLAUDE.md`, `docs/` and `spike/` are stripped. It is what release-please watches
  and what `release.yml` tags, packs and publishes. It is also the GitHub default branch, so it
  is the tree a consumer arriving from nuget.org lands on — that is the entire point.

Feature branches (`feat/**`, `fix/**`) branch from `develop` and PR back into `develop`.
`ci.yml`'s `branch-policy` job rejects any PR into `main` that is not a `release/promote-*`
branch or release-please's own Release PR, so a mis-targeted PR fails rather than merging quietly.

**Promote with `scripts/promote-to-main.sh`.** It merges `develop` into a new
`release/promote-*` branch based on `main`, purges the excluded paths, and opens a PR. Promotion
is a *real* `git merge`, not a generated snapshot, so every Conventional Commit subject from
`develop` stays reachable from `main` — that is what keeps release-please working unchanged.
**Merge that PR with a merge commit, never a squash**; squashing collapses those subjects and
release-please would compute the wrong bump, or none.

`scripts/` is deliberately *not* excluded from `main`: it is release tooling, and stripping it
would let the promote script delete itself out from under the running bash process.

### Never merge `main` into `develop`

`main` carries *deletions* of `CLAUDE.md`, `docs/` and `spike/`. A `git merge main` on `develop`
would propagate them and wipe the development record. This is the single most dangerous operation
in this repo.

`CHANGELOG.md` and `.release-please-manifest.json` are **main-owned** — release-please writes
them there. Never edit them on `develop`: as long as `develop` leaves them alone, git resolves
them cleanly on every promote (main's side wins), and the moment `develop` edits either one,
every promote conflicts. Sync them back by content copy, which carries no deletions:

```bash
git switch develop
git checkout main -- CHANGELOG.md .release-please-manifest.json
git commit -m "chore: sync changelog and manifest from the last release"
```

`README.md` must stay byte-identical on both branches — it is packed into both `.nupkg`s and
asserted by CI, and divergence would conflict on every promote. Edit it on `develop` only.

There is no hotfix branch. An urgent fix goes `fix/*` → `develop` → promote; a second path into
`main` would add a way around CI without adding meaningful speed.
```

- [ ] **Step 2: Retire the hand-maintained `## Unreleased` guidance**

Inside the existing "Releasing" section, replace this sentence:

```
A manual `git tag v1.2.3 && git push origin v1.2.3` still works as a fallback — `release.yml` only
cares that a `v*` tag arrived, not how — but if you tag manually, **move the current `## Unreleased`
content in `CHANGELOG.md` under a new `## [X.Y.Z] - YYYY-MM-DD` heading yourself first**;
```

with:

```
A manual `git tag v1.2.3 && git push origin v1.2.3` still works as a fallback — `release.yml` only
cares that a `v*` tag arrived, not how — but tag `main`, never `develop`, and write the
`## [X.Y.Z] - YYYY-MM-DD` heading into `CHANGELOG.md` **on `main`** first (`CHANGELOG.md` is
main-owned; see "Branches");
```

- [ ] **Step 3: Update the Layout block**

Replace the fenced Layout block's `spike/` and `docs/` rows and add `scripts/`, so the block reads:

```
src/DocToolkit/                                         the library
tests/DocToolkit.Tests/                                 182 tests, including StreamOverloadTests, AirGapGuardTests, DependencyGuardTests
src/DocToolkit.Extensions.DependencyInjection/          DI extensions package (services.AddDocToolkit())
tests/DocToolkit.Extensions.DependencyInjection.Tests/  23 tests, including ServiceCollectionExtensionsTests
samples/ConsoleSample/                                  core package, all five capabilities
samples/MinimalApiSample/                               DI extensions package, one endpoint per interface
docfx/                                                  DocFX site source, published to GitHub Pages on release
scripts/                                                promote-to-main.sh and its test — on main too, see Branches
spike/                                                  original proof-of-concept, kept as reference — do not modify (develop only)
docs/                                                   design docs and implementation plans this was built from (develop only)
```

- [ ] **Step 4: Verify**

```bash
grep -c '^## Branches$' CLAUDE.md
grep -c 'Never merge `main` into `develop`' CLAUDE.md
grep -c 'scripts/promote-to-main.sh' CLAUDE.md
grep -c 'develop only' CLAUDE.md
```

Expected: `1`, `1`, at least `1`, and `2`.

- [ ] **Step 5: Commit**

```bash
git add CLAUDE.md
git commit -F - <<'EOF'
docs: describe the two-branch model and its one dangerous operation

develop is the trunk, main is release-only and carries no development record.
The section leads with what must never happen - merging main back into develop,
which would propagate main's deletions and wipe docs/, spike/ and this file -
because that is the failure this model makes possible and nothing else guards it.

Also retires the hand-maintained "## Unreleased" fallback: CHANGELOG.md is now
main-owned, so a manual tag writes its heading there, not on develop.
EOF
git push origin develop
```

---

### Task 6: First promote — produce main's stripped tree

This is the verification step for the whole design. It happens before any tag is pushed, so a
wrong exclusion costs a re-run rather than an irreversible nuget.org publish.

**Files:** none directly — this runs the tooling built in Tasks 2-5.

**Interfaces:**
- Consumes: `scripts/promote-to-main.sh` (Task 2), `branch-policy` (Task 3).

- [ ] **Step 1: Confirm develop is fully pushed**

```bash
git status --short
git log --oneline origin/develop..develop
```

Expected: both print nothing.

- [ ] **Step 2: Run the promote script**

```bash
bash scripts/promote-to-main.sh
```

Expected output: `created release/promote-<timestamp>`, then `pushed ...`, then `opened PR`
followed by the PR URL.

- [ ] **Step 3: Verify the promoted tree before merging**

```bash
BR=$(git for-each-ref --format='%(refname:short)' --sort=-committerdate 'refs/heads/release/promote-*' | head -1)
echo "branch: $BR"
echo "--- must be ABSENT ---"
git ls-tree -r --name-only "$BR" | grep -E '^(CLAUDE\.md|docs/|spike/)' || echo "(none - correct)"
echo "--- must be PRESENT ---"
for p in src/DocToolkit/DocToolkit.csproj tests/DocToolkit.Tests/DocToolkit.Tests.csproj \
         samples/ConsoleSample/ConsoleSample.csproj docfx/docfx.json CHANGELOG.md \
         .gitattributes Dockerfile.linux-test scripts/promote-to-main.sh \
         .github/workflows/release.yml; do
  if git ls-tree -r --name-only "$BR" | grep -qx "$p"; then echo "ok   $p"; else echo "MISS $p"; fi
done
```

Expected: the ABSENT section prints `(none - correct)`, and every PRESENT line prints `ok`.
If any prints `MISS`, stop — do not merge the PR. Fix the exclusion list in
`scripts/promote-to-main.sh` on `develop` and re-run from Step 2.

- [ ] **Step 4: Prove the stripped tree still builds and tests**

`.worktrees/` is already gitignored, so the checkout cannot be committed by accident. Use it
rather than `/tmp` — `dotnet` is a Windows binary and does not understand Git Bash's `/tmp/...`
paths.

```bash
git worktree add .worktrees/promote-check "$BR"
dotnet build .worktrees/promote-check/DocToolkit.sln -c Release -warnaserror
dotnet test  .worktrees/promote-check/DocToolkit.sln -c Release --no-build
```

Expected: build succeeds with 0 warnings; 410 test results pass (205 tests × 2 TFMs).
This is the direct check that stripping `spike/` did not break the solution.

Clean up:

```bash
git worktree remove .worktrees/promote-check
```

- [ ] **Step 5: Confirm CI is green on the PR**

```bash
gh pr checks "$BR" --watch
```

Expected: all jobs pass, including `main is release-only` (the promote branch is allow-listed),
`commit message format` (the merge-commit exemption from Task 3 is what makes this pass), and
`build docs site` (proving `docs/` was never a DocFX input).

- [ ] **Step 6: Merge the PR with a merge commit**

```bash
gh pr merge "$BR" --merge --delete-branch
```

`--merge` is mandatory. **Do not use `--squash` or `--rebase`** — either destroys the
Conventional Commit subjects release-please needs.

- [ ] **Step 7: Verify main's remote tree**

```bash
git fetch origin
git ls-tree -r --name-only origin/main | grep -E '^(CLAUDE\.md|docs/|spike/)' || echo "main is clean - correct"
git ls-tree -r --name-only origin/main | grep -cE '^(src|tests|samples)/'
```

Expected: `main is clean - correct`, and a non-zero count of source files.

- [ ] **Step 8: Check what release-please proposed**

```bash
gh pr list --base main --state open
```

Expected: a Release PR from `release-please--branches--main` now exists or has been updated.
**Read its diff before merging it** — per `CLAUDE.md`, a version bump with an empty-looking
changelog entry is a signal to hold off. This plan does not merge it; releasing stays a separate
human decision.

---

### Task 7: Repository settings — default branch and branch protection

**Files:** none — GitHub repository settings.

- [ ] **Step 1: Confirm `main` is the default branch**

```bash
gh repo view Ank-KhoaHo/DocToolkit --json defaultBranchRef -q .defaultBranchRef.name
```

Expected: `main`. If it prints anything else:

```bash
gh api -X PATCH repos/Ank-KhoaHo/DocToolkit -f default_branch=main
```

- [ ] **Step 2: Protect `main`**

```bash
gh api -X PUT repos/Ank-KhoaHo/DocToolkit/branches/main/protection \
  --input - <<'JSON'
{
  "required_status_checks": {
    "strict": true,
    "contexts": ["build & test (ubuntu-latest)", "no native binaries / no banned packages", "main is release-only"]
  },
  "enforce_admins": false,
  "required_pull_request_reviews": null,
  "restrictions": null,
  "allow_force_pushes": false,
  "allow_deletions": false
}
JSON
```

`enforce_admins` is `false` deliberately: you are the sole maintainer, and locking yourself out
of your own release branch during a botched promote would be worse than the protection is worth.

- [ ] **Step 3: Protect `develop`**

```bash
gh api -X PUT repos/Ank-KhoaHo/DocToolkit/branches/develop/protection \
  --input - <<'JSON'
{
  "required_status_checks": {
    "strict": false,
    "contexts": ["build & test (ubuntu-latest)", "promote script logic"]
  },
  "enforce_admins": false,
  "required_pull_request_reviews": null,
  "restrictions": null,
  "allow_force_pushes": false,
  "allow_deletions": false
}
JSON
```

- [ ] **Step 4: Verify both**

```bash
gh api repos/Ank-KhoaHo/DocToolkit/branches/main/protection    -q '.required_status_checks.contexts'
gh api repos/Ank-KhoaHo/DocToolkit/branches/develop/protection -q '.required_status_checks.contexts'
```

Expected: the two context lists above. If either call returns `404`, branch protection is not
available for this repository's plan — record that and move on; `branch-policy` still runs on
every PR into `main` regardless, so the core guarantee holds without it.

- [ ] **Step 5: No commit**

Repository settings are not tracked in git.

---

### Task 8: Retarget the in-flight `di-extensions-parity` work

There is an active worktree at `.worktrees/di-extensions-parity` on branch `di-extensions-parity`,
created before `develop` existed and therefore based on `main`. Its implementation plan
(`docs/2026-08-03-di-extensions-parity-implementation-plan.md`) assumes a single-branch repo.

**Files:**
- Modify: git refs only, plus a one-line note in
  `docs/2026-08-03-di-extensions-parity-implementation-plan.md`.

- [ ] **Step 1: See how far the branch has diverged**

```bash
git log --oneline main..di-extensions-parity
git -C .worktrees/di-extensions-parity status --short
```

Note the commit count. If the worktree has uncommitted changes, commit them there before
continuing.

- [ ] **Step 2: Rebase it onto `develop`**

```bash
git fetch origin
git -C .worktrees/di-extensions-parity rebase origin/develop
```

Expected: `Successfully rebased and updated refs/heads/di-extensions-parity`.

If it reports conflicts, resolve them in the worktree, `git add` the resolved files, and
`git rebase --continue`. The branch was based on a commit that is an ancestor of `develop`, so
conflicts are unlikely.

- [ ] **Step 3: Rename the branch to match the `feat/**` convention CI triggers on**

```bash
git -C .worktrees/di-extensions-parity branch -m feat/di-extensions-parity
```

`ci.yml` triggers pushes on `feat/**` and `fix/**`; the bare name `di-extensions-parity` matches
neither, so pushes to it would run no CI at all.

- [ ] **Step 4: Add the retargeting note to its plan**

Insert immediately after the `**Goal:**` line of
`docs/2026-08-03-di-extensions-parity-implementation-plan.md`:

```markdown
> **Branch note:** this plan predates the two-branch model
> (`docs/2026-08-03-branching-model-design.md`). Work on `feat/di-extensions-parity` and open the
> PR against **`develop`**, not `main` — `ci.yml`'s `branch-policy` job rejects PRs into `main`
> that are not promote or Release PRs.
```

- [ ] **Step 5: Commit the note on `develop`**

```bash
git switch develop
git add docs/2026-08-03-di-extensions-parity-implementation-plan.md
git commit -F - <<'EOF'
docs: note that the DI parity plan now targets develop

The plan was written against a single-branch repo. Its branch is rebased onto
develop and renamed feat/di-extensions-parity so ci.yml's feat/** trigger picks
it up - the bare name matched no trigger and would have run no CI at all.
EOF
git push origin develop
```

- [ ] **Step 6: Verify**

```bash
git branch --list 'feat/di-extensions-parity'
git log --oneline origin/develop..feat/di-extensions-parity | head
git worktree list
```

Expected: the branch exists under its new name, its commits sit on top of `develop`, and the
worktree still resolves.

---

## Verification of the whole plan

After Task 8, these must all hold:

```bash
# main carries no development record
git ls-tree -r --name-only origin/main | grep -E '^(CLAUDE\.md|docs/|spike/)' || echo "ok - main clean"

# main carries everything the release needs
git ls-tree -r --name-only origin/main | grep -cE '^(src|tests|samples|docfx|scripts|\.github)/'

# develop carries everything
git ls-tree -r --name-only origin/develop | grep -cE '^(CLAUDE\.md|docs/|spike/)'

# README is identical on both - if this prints anything, every future promote conflicts
git diff origin/main:README.md origin/develop:README.md && echo "ok - README identical"

# the promote logic still passes its own test
bash scripts/test-promote.sh
```

Expected: `ok - main clean`; a non-zero count; a non-zero count; `ok - README identical` with no
diff output; and `failed: 0`.
