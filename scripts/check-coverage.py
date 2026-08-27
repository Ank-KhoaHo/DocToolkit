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
`dotnet test` on this solution emits four distinct cobertura files - two test
projects x two target frameworks - and net8.0 and net10.0 execute the same tests
over the same source. Treating those as separate samples understates coverage,
so a line hit by ANY report counts as covered here.

MAX-PER-LINE, not a sum, and that is load-bearing rather than a stylistic
choice. Measured on the runner: CI yields EIGHT files, not four, because vstest
stages a second identical copy of each report under
`_<machine>_<timestamp>/In/<machine>/`. Verified identical by checksum. Taking
the max per line makes duplicate reports a no-op; summing valid-line counts
would have been skewed by staging directories nobody controls, and the local
run - which produces four - would have disagreed with CI for a reason that has
nothing to do with coverage.

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
# Re-measured 2026-08-15: DocToolkit 96.30 line (2627/2728) / 90.64 branch (920/1015),
# DocToolkit.Extensions.DependencyInjection 100 / 100. Confirmed IDENTICAL on the Linux
# CI runner, which is the only leg that runs this gate - so these are not a
# Windows-only artefact.
#
# DocToolkit sits a little under the measurement on purpose. A floor set exactly
# at the current number fails on any single uncovered line, including ones that
# are correct to leave uncovered - a guard clause for a state that cannot be
# constructed, say - which trains people to lower the floor rather than write a
# test. The slack is roughly 8 lines and 16 branches: far too small to hide
# an untested feature, far too large to trip on noise.
#
# THE PREVIOUS FLOORS (95.0 / 87.5) HAD DRIFTED, and the drift is the reason to state
# the slack in LINES rather than points. They were set 2026-08-08 against 95.44 / 88.34
# with exactly that ~8-line, ~16-branch intent. Coverage then climbed to 96.30 / 90.64
# and nobody re-tightened, so by 2026-08-15 the slack was 35 lines and 31 branches -
# comfortably enough to absorb a whole untested public method, which is precisely what
# the gate exists to refuse. A percentage looks stable while the absolute number it
# implies quadruples.
#
# The DI package is held at 100 because it is PURE DELEGATION - every member is
# one line calling a static method, so any member that is not covered is simply
# a member nobody tested, and the fix is always a two-line test. See CLAUDE.md
# on how often that mirror has gone stale.
# ONE FLOOR PER SHIPPED ASSEMBLY since the per-format project split. The package still ships as
# one nupkg; it is now built from seven projects, and this gate FAILS on an assembly with no floor
# rather than warning - which is how the split was noticed here rather than sliding through.
#
# THE SLACK HAD TO BE RE-DERIVED, and that is the part worth reading. The 8-line / 16-branch
# figure above was calibrated against ONE assembly of ~2,500 lines and ~1,300 branches. Applied
# unchanged to seven smaller ones it becomes absurd: DocToolkit.Pptx has 98 branches in total, so a
# 16-branch slack is a 16-POINT drop, and its floor would sit at 66.3% against 82.65% measured - a
# gate that could not fail. The same absolute number means something entirely different once the
# thing it is measured against is a seventh of the size.
#
# So the slack is 3 lines and 6 branches PER ASSEMBLY, CAPPED AT 2 POINTS - the same threshold
# RATCHET_SLACK below uses to call a floor too loose. Without the cap the two disagreed: a 6-branch
# allowance on a 98-branch assembly is 6 points, so the gate passed while the ratchet immediately
# advised tightening it, on every run. A guard that nags on every green build is one people learn
# to scroll past.
#
# TOTAL COVERAGE DID NOT DROP, and that was measured before any floor was written: 96.45% line and
# 90.71% branch across all seven, against 96.30 / 90.64 for the single assembly beforehand.
# DocToolkit's own number falls to 92.94% only because the well-covered editors and primitives
# moved out from under it, leaving the converters. That is arithmetic, not a regression - and
# checking it is the difference between re-deriving a floor and lowering one.
FLOORS = {
    "DocToolkit": (92.5, 85.0),
    "DocToolkit.Primitives": (97.9, 91.6),
    "DocToolkit.Docx": (96.8, 82.6),
    "DocToolkit.Html": (93.5, 94.7),
    "DocToolkit.Pdf": (98.4, 95.6),
    "DocToolkit.Pptx": (96.3, 80.7),
    # Ratcheted 2026-08-26 with A68 (conditional formats, validations, widths, freeze,
    # autofilter), from 95.5 / 92.4 - the branch floor was 3.81 points below reality,
    # which is a floor that has stopped catching anything.
    #
    # It was first set at 95.9 against a measured 96.21, and that FAILED CI two commits
    # later at 95.85 - the closed-vocabulary guards added `_ => throw` arms that are
    # unreachable by construction and correct to leave uncovered. Which is precisely the
    # failure this file's own header warns about: a floor set at the measurement fails on
    # the first line it is right to not cover, and trains people to lower it. Measured
    # 97.04 / 95.85 now, with the floor a genuine step below both.
    "DocToolkit.Xlsx": (96.3, 95.0),
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

        # BOTH metrics, and the second one is why this row exists. Until 2026-08-15 this
        # checked line_pct only, so the branch floor drifted to 3.14 points of slack - four
        # times its design - while the advisory built to catch exactly that stayed silent,
        # because branch was never compared to anything. A ratchet with a blind spot on one
        # of the two numbers it guards is worse than none: the green run reads as proof.
        for label, pct, floor in (("line", line_pct, line_floor),
                                  ("branch", branch_pct, branch_floor)):
            if pct >= floor + RATCHET_SLACK:
                print(f"::notice::{asm} {label} coverage is {pct:.2f}%, "
                      f"{pct - floor:.2f} points above its floor. Consider raising it to "
                      f"{pct - 0.3:.1f} - a floor far below reality stops catching regressions.")

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
