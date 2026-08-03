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

# The branch name comes from the promote script's own `created $BRANCH` line
# (captured by run_promote below), not from sorting refs by committer date.
# Committer-date has 1-second granularity and git's tiebreak on equal sort
# keys is refname ascending, so two promotes landing in the same second would
# make ref-sorting silently return the FIRST branch - re-testing test 1's tree
# instead of the second promote's.
branch_from_output() {
  sed -n 's/^created //p' "$TMP/promote-out.txt" | head -1
}

setup_repo() {
  cd /
  rm -rf "$TMP/origin.git" "$TMP/work"
  git init -q --bare "$TMP/origin.git"
  # git init + remote add instead of `git clone` of the still-empty bare repo -
  # clone prints "warning: You appear to have cloned an empty repository." even
  # with -q, which pollutes CI logs; init+remote add reaches the same end state
  # (a repo with an `origin` remote and the standard fetch refspec) silently.
  mkdir -p "$TMP/work"
  git init -q "$TMP/work"
  git -C "$TMP/work" remote add origin "$TMP/origin.git"
  cd "$TMP/work"
  git config user.email test@example.com
  git config user.name "Test"
  git config commit.gpgsign false
  # Windows checkouts default to autocrlf=true, which floods the run with
  # "LF will be replaced by CRLF" warnings and buries a real failure.
  git config core.autocrlf false
  mkdir -p src docs spike scripts
  echo v1 > src/app.txt
  echo agent-instructions > CLAUDE.md
  echo design > docs/design.md
  echo poc > spike/poc.txt
  echo keep > scripts/keep.sh
  git add -A
  git commit -qm "chore: seed"
  git branch -M main
  git push -q -u origin main
  git switch -qc develop
  git push -q -u origin develop
}

run_promote() { # run_promote -> echoes exit code, never aborts the test run
  # Combined stdout+stderr is captured to a file (overwritten each call) rather
  # than discarded, so callers can assert on the script's messages - both the
  # `created $BRANCH` line (see branch_from_output) and error text.
  set +e
  ( cd "$TMP/work" && PROMOTE_PUSH=0 PROMOTE_OPEN_PR=0 bash "$PROMOTE" ) >"$TMP/promote-out.txt" 2>&1
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
BR="$(branch_from_output)"
if [ -n "$BR" ]; then check "a release/promote-* branch was created" 0; else check "a release/promote-* branch was created" 1; fi
tree_of "$BR"
assert_has    "src/app.txt"
assert_has    "scripts/keep.sh"
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
BR2="$(branch_from_output)"
if [ -n "$BR2" ] && [ "$BR2" != "$BR" ]; then
  check "second promote created a branch distinct from test 1's" 0
else
  check "second promote created a branch distinct from test 1's" 1
fi
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
# Deleting the excluded-set conflict guard entirely would still leave `git
# commit` refusing on unmerged index entries under `set -e`, so a bare
# non-zero-exit check above would pass for the wrong reason. Pin the guard
# itself by asserting its actual message, not just any failure.
if grep -q "conflict outside the excluded set" "$TMP/promote-out.txt"; then
  check "error message names the excluded-set guard" 0
else
  check "error message names the excluded-set guard" 1
fi
if grep -q "src/app.txt" "$TMP/promote-out.txt"; then
  check "error message names the conflicting path" 0
else
  check "error message names the conflicting path" 1
fi

echo
echo "passed: $pass   failed: $fail"
[ "$fail" -eq 0 ]
