#!/usr/bin/env python3
"""Decide whether a Dependabot pull request may be auto-merged.

Auto-merging dependency updates rests on "CI is green, so the change is safe".
For this repository that is true of NuGet bumps and **not uniformly true of
GitHub Actions bumps**, for a reason `.github/dependabot.yml` already records:

    PR CI cannot catch a regression there: it only runs ci.yml, while
    release-please.yml, release.yml and docs.yml are exercised solely after a
    merge. A green check on such a PR means less than it looks like.

Measured 2026-08-08, six of the ten actions this repo uses are never run by
ci.yml - NuGet/login, actions/attest-build-provenance, actions/configure-pages,
actions/deploy-pages, actions/upload-pages-artifact and
googleapis/release-please-action. Between them they hold the publish path, the
provenance attestation and the docs deploy. A green check on a PR bumping one of
those is evidence about a workflow that never ran.

So the rule is not "actions are excluded". It is:

    an actions bump may auto-merge only if EVERY action it touches is one that
    ci.yml actually runs.

That set is DERIVED from the workflows on every invocation rather than written
down here, so it cannot go stale when someone adds an action to release.yml -
the same reason gen-third-party-notices.py reads the lockfile instead of keeping
a table. Adding a new action to ci.yml widens the rule automatically; adding one
to release.yml narrows it automatically.

Patch only, in every ecosystem. A minor bump is a behaviour change somebody
should look at, and the test-dependency group deliberately mixes minor with
patch, so grouped PRs containing a minor are held.

Usage:
    python scripts/automerge-eligible.py <update-type> <dependency-names-csv>

Exit 0 = eligible (and prints why), 1 = not eligible (and prints why).
"""

import glob
import os
import re
import sys

PATCH = "version-update:semver-patch"


def actions_exercised_by_ci(workflow_dir=".github/workflows"):
    """Every `uses:` in ci.yml — the actions a pull request actually runs."""
    ci = os.path.join(workflow_dir, "ci.yml")
    if not os.path.exists(ci):
        sys.exit(f"::error::{ci} not found; cannot tell which actions PR CI runs.")
    with open(ci, encoding="utf-8") as handle:
        return {m.group(1) for m in re.finditer(r"uses:\s*([A-Za-z0-9._\-]+/[A-Za-z0-9._\-]+)@", handle.read())}


def all_actions(workflow_dir=".github/workflows"):
    found = set()
    for path in glob.glob(os.path.join(workflow_dir, "*.yml")):
        with open(path, encoding="utf-8") as handle:
            found |= {m.group(1)
                      for m in re.finditer(r"uses:\s*([A-Za-z0-9._\-]+/[A-Za-z0-9._\-]+)@", handle.read())}
    return found


def main():
    if len(sys.argv) != 3:
        sys.exit(__doc__)

    update_type = sys.argv[1].strip()
    names = [n.strip() for n in sys.argv[2].split(",") if n.strip()]

    if update_type != PATCH:
        print(f"HOLD: update type is '{update_type}', not {PATCH}. "
              "Only patch updates auto-merge; anything else gets a human.")
        return 1

    if not names:
        print("HOLD: no updated dependencies reported. Refusing to auto-merge a PR "
              "whose contents could not be determined.")
        return 1

    exercised = actions_exercised_by_ci()
    known_actions = all_actions()

    # A dependency name is an action only if it appears in a workflow. NuGet package
    # ids never do, so this classifies without needing the ecosystem passed in.
    touched_actions = [n for n in names if n in known_actions]
    unexercised = sorted(n for n in touched_actions if n not in exercised)

    if unexercised:
        print("HOLD: this PR bumps action(s) that ci.yml never runs, so a green check "
              "is evidence about a workflow that did not execute:")
        for name in unexercised:
            print(f"  - {name}")
        print("Review by hand, then merge manually.")
        return 1

    print(f"ELIGIBLE: patch update to {', '.join(names)}.")
    if touched_actions:
        print("  Every action touched is one ci.yml runs: " + ", ".join(sorted(touched_actions)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
