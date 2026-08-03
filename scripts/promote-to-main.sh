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
