#!/usr/bin/env python3
"""Fail the build when line or branch coverage regresses below an agreed floor.

Coverage was REPORTED and never ENFORCED: ci.yml uploaded to Codecov with
`fail_ci_if_error: false` and no codecov.yml target existed, so the number could
fall to anything without a single check going red. This is the gate.

Why a script here rather than the two obvious alternatives:

  - Codecov statuses would put a third-party service in the merge path. This
    repository pins its SDK, pins every action by SHA and restores with
    --locked-mode; letting an external service decide whether a PR can merge is
    out of character, and Codecov is currently only half-configured here (it
    posts no status checks at all, only a "please install the app" comment).
    codecov.yml still exists alongside this, marked informational, so the patch-
    coverage signal keeps arriving without ever blocking.
  - coverlet's own /p:Threshold needs coverlet.msbuild. The test projects use
    coverlet.collector, so that would mean a new package for a check that is
    thirty lines of XML reading.

MERGING THE REPORTS IS THE POINT, and is what neither alternative does properly.
`dotnet test` on this solution emits FOUR cobertura files - two test projects x
two target frameworks - and net8.0 and net10.0 execute the same tests over the
same source. Treating those as separate samples understates coverage, so a line
hit by ANY report counts as covered here.

Usage:
    python scripts/check-coverage.py <results-dir>
"""

import collections
import glob
import os
import sys
import xml.etree.ElementTree as ET

# Floors, per assembly: (line %, branch %).
#
# Measured 2026-08-08 on d543ef6: DocToolkit 95.44 line / 88.34 branch,
# DocToolkit.Extensions.DependencyInjection 100 / 100.
#
# DocToolkit sits a little under the measurement on purpose. A floor set exactly
# at the current number fails on any single uncovered line, including ones that
# are correct to leave uncovered - a guard clause for a state that cannot be
# constructed, say - which trains people to lower the floor rather than write a
# test. The slack here is roughly 8 lines and 16 branches: far too small to hide
# an untested feature, far too large to trip on noise.
#
# The DI package is held at 100 because it is PURE DELEGATION - every member is
# one line calling a static method, so any member that is not covered is simply
# a member nobody tested, and the fix is always a two-line test. See CLAUDE.md
# on how often that mirror has gone stale.
FLOORS = {
    "DocToolkit": (95.0, 87.5),
    "DocToolkit.Extensions.DependencyInjection": (100.0, 100.0),
}

# Report when an assembly has climbed this far above its floor, so the gate does
# not quietly stop meaning anything. Advisory - it never fails the build.
RATCHET_SLACK = 2.0


def load(results_dir):
    """Merge every cobertura report under results_dir into per-file hit maps."""
    paths = sorted(glob.glob(
        os.path.join(results_dir, "**", "coverage.cobertura.xml"), recursive=True))
    if not paths:
        sys.exit(f"::error::No coverage.cobertura.xml under {results_dir}. "
                 "Coverage was not collected, so this gate proves nothing - "
                 "that is a failure, not a pass.")

    lines = collections.defaultdict(dict)     # (asm, file) -> {line: hits}
    branches = collections.defaultdict(dict)  # (asm, file) -> {line: (covered, total)}

    for path in paths:
        for package in ET.parse(path).getroot().iter("package"):
            asm = package.get("name")
            for cls in package.iter("class"):
                key = (asm, cls.get("filename"))
                for line in cls.iter("line"):
                    number, hits = int(line.get("number")), int(line.get("hits"))
                    lines[key][number] = max(lines[key].get(number, 0), hits)

                    if line.get("branch") != "True":
                        continue
                    condition = line.get("condition-coverage", "")
                    if "(" not in condition:
                        continue
                    covered, total = condition.split("(")[1].rstrip(")").split("/")
                    previous = branches[key].get(number, (0, 0))
                    branches[key][number] = (max(previous[0], int(covered)), int(total))

    return paths, lines, branches


def percent(part, whole):
    """Coverage of nothing is complete coverage - an assembly with no branches
    must not fail a branch floor it cannot possibly meet."""
    return 100.0 * part / whole if whole else 100.0


def main():
    if len(sys.argv) != 2:
        sys.exit(__doc__)

    paths, lines, branches = load(sys.argv[1])
    print(f"Merged {len(paths)} cobertura report(s) from {sys.argv[1]}")

    totals = collections.defaultdict(lambda: [0, 0, 0, 0])
    per_file = collections.defaultdict(list)

    for (asm, filename), hits in lines.items():
        covered = sum(1 for h in hits.values() if h > 0)
        valid = len(hits)
        br = branches[(asm, filename)].values()
        br_covered, br_valid = sum(c for c, _ in br), sum(t for _, t in br)

        bucket = totals[asm]
        for index, value in enumerate((covered, valid, br_covered, br_valid)):
            bucket[index] += value
        if valid - covered:
            per_file[asm].append((valid - covered, os.path.basename(filename.replace("\\", "/"))))

    failures = []
    print()
    print(f"{'assembly':<46}{'line':>18}{'branch':>18}")
    for asm, (covered, valid, br_covered, br_valid) in sorted(totals.items()):
        line_pct, branch_pct = percent(covered, valid), percent(br_covered, br_valid)
        line_floor, branch_floor = FLOORS.get(asm, (0.0, 0.0))

        def mark(actual, floor):
            return f"{actual:6.2f}% (>={floor:.1f})" + ("  " if actual >= floor else " X")

        print(f"{asm:<46}{mark(line_pct, line_floor):>18}{mark(branch_pct, branch_floor):>18}")

        if asm not in FLOORS:
            # Fails rather than warns, deliberately. A new assembly with no floor is
            # precisely the silently-unprotected case this gate exists to prevent, and a
            # warning in a green build is a warning nobody reads. Adding the entry is one
            # line; being asked to choose a number is the point.
            failures.append(f"{asm} has no floor in FLOORS. Add one - measured "
                            f"{line_pct:.2f}% line / {branch_pct:.2f}% branch.")
            continue

        if line_pct < line_floor:
            failures.append(f"{asm}: line coverage {line_pct:.2f}% is below the "
                            f"{line_floor:.1f}% floor")
        if branch_pct < branch_floor:
            failures.append(f"{asm}: branch coverage {branch_pct:.2f}% is below the "
                            f"{branch_floor:.1f}% floor")

        if line_pct >= line_floor + RATCHET_SLACK:
            print(f"::notice::{asm} line coverage is {line_pct:.2f}%, "
                  f"{line_pct - line_floor:.2f} points above its floor. Consider raising it - "
                  "a floor far below reality stops catching regressions.")

    if failures:
        print()
        for asm, missed in sorted(per_file.items()):
            worst = sorted(missed, reverse=True)[:5]
            if worst:
                print(f"  most uncovered lines in {asm}: "
                      + ", ".join(f"{name} ({count})" for count, name in worst))

    print()
    if not failures:
        print("Coverage floors held.")
        return 0

    for failure in failures:
        print(f"::error::{failure}")
    return 1


if __name__ == "__main__":
    sys.exit(main())
