#!/usr/bin/env python3
"""Generate CONTRIBUTING.md's guard inventory from the workflows themselves.

WHY THIS IS GENERATED. CONTRIBUTING.md told contributors that the `formatting`
check runs "`check-readme-coverage.py` ... and the third-party notices check".
Measured 2026-08-15 against the workflows: that job runs SIX guard scripts, and
the notices check is not one of them - it runs in `no native binaries / no banned
packages`, which the SAME TABLE says two rows earlier. So a contributor whose
`formatting` check went red on check-core-sharing.py or gen-capability-matrix.py
was sent to look at a README and a notices file, neither of which had anything to
do with it.

That is the cost this file exists to remove. A red check is exactly the moment a
contributor is least able to absorb a wrong answer, and the list they are reading
was hand-maintained.

The same drift had already happened twice in the maintainer-side checklist
(SOURCE-QUALITY.md Part 1): three CI-wired guards missing on the day it was
written, and a closing paragraph describing a check as unshipped that had shipped
the same day.

DERIVED, NOT LISTED. The mapping comes from .github/workflows/*.yml, so adding a
guard to a job automatically brings it into the table, and moving one between
jobs moves its row. Same principle as gen-third-party-notices.py reading the
lockfile, automerge-eligible.py reading the workflows, check-readme-coverage.py
reading the approved API, gen-capability-matrix.py reading the approved API, and
StreamOverloadTests reflecting over the assembly. Derive, do not remember.

WHY THE TARGET IS CONTRIBUTING.md AND NOT SOURCE-QUALITY.md. SOURCE-QUALITY.md is
gitignored - it is maintainer guidance and never reaches a runner. A `--check`
pointed at it would find no file in CI and pass, every time, which is the
vacuously-green failure this repository has already had to fix twice in its
probe-based suites. CONTRIBUTING.md is tracked, so the check actually discriminates,
and the contributor-facing copy is the one that was wrong anyway. Part 1 of
SOURCE-QUALITY.md now points here instead of keeping a second copy.

WHAT IT DOES NOT ATTEMPT. It reports which guard runs in which check. It does not
describe what a guard means or what to do when it fails - that is the prose table
above it in CONTRIBUTING.md, which is worth keeping and cannot be derived.

USAGE

    python scripts/gen-guard-inventory.py            # write
    python scripts/gen-guard-inventory.py --check    # verify, change nothing

`--check` is what CI runs. It exits non-zero when CONTRIBUTING.md no longer
matches the workflows, and prints a diff.
"""
from __future__ import annotations

import difflib
import pathlib
import re
import sys

REPO = pathlib.Path(__file__).resolve().parent.parent
WORKFLOWS = REPO / ".github" / "workflows"
SCRIPTS = REPO / "scripts"
PAGE = REPO / "CONTRIBUTING.md"

BEGIN = "<!-- BEGIN GENERATED (scripts/gen-guard-inventory.py) - do not edit by hand -->"
END = "<!-- END GENERATED -->"

# A job key sits at two spaces; a job's display name at four, with no leading dash.
# The dash matters: `    - name:` is a STEP name and there are dozens per file.
JOB_KEY = re.compile(r"^  ([a-z][a-z0-9_-]*):\s*$")
JOB_NAME = re.compile(r"^    name:\s*(.+?)\s*$")

# Matrix expressions read as noise in a table. `build & test (${{ matrix.label }})`
# becomes `build & test (…)`, which is how CONTRIBUTING already refers to that family.
MATRIX = re.compile(r"\$\{\{[^}]*\}\}")

# Not every script under scripts/ is a guard. nuget-stats.py collects download
# figures on a schedule; it decides nothing and gates nothing. Excluded by name
# rather than by heuristic, because a heuristic that silently dropped a real guard
# would be the same class of bug this script exists to catch.
# Scripts that are not guards, each for a stated reason. An entry here is a hole in the
# inventory, so argue for one rather than adding it to make this script quiet.
#
#   nuget-stats.py        reports download figures; asserts nothing.
#   gen-backlog-index.py  regenerates BACKLOG.html from the ticket specs. It CANNOT be wired into
#                         a workflow: both its input (docs/superpowers/specs/) and its output
#                         (BACKLOG.html) are gitignored, so CI never sees either. What guards its
#                         output is check-docs-layout.py's LOCAL mode, which compares the index
#                         against the specs and names this script in the failure message.
NOT_A_GUARD = {"nuget-stats.py", "gen-backlog-index.py"}


def guards() -> list[str]:
    return sorted(p.name for p in SCRIPTS.glob("*.py") if p.name not in NOT_A_GUARD)


def scan() -> tuple[list[tuple[str, str, list[str]]], dict[str, list[str]]]:
    """Return (rows, script -> where it runs).

    A row is (workflow file, check display name, scripts invoked).
    """
    names = guards()
    rows: list[tuple[str, str, list[str]]] = []
    seen: dict[str, list[str]] = {n: [] for n in names}

    for wf in sorted(WORKFLOWS.glob("*.yml")):
        text = wf.read_text(encoding="utf-8")
        job_key: str | None = None
        display: dict[str, str] = {}
        found: dict[str, list[str]] = {}

        for line in text.split("\n"):
            m = JOB_KEY.match(line)
            if m:
                job_key = m.group(1)
                found.setdefault(job_key, [])
                continue

            if job_key is None:
                continue

            n = JOB_NAME.match(line)
            if n and job_key not in display:
                display[job_key] = MATRIX.sub("…", n.group(1)).strip("\"'")
                continue

            # A comment mentioning a script is not an invocation. Both YAML and the
            # shell inside `run: |` use `#`, so one test covers both.
            if line.lstrip().startswith("#"):
                continue

            for g in names:
                if g in line and g not in found[job_key]:
                    found[job_key].append(g)

        for key, invoked in found.items():
            if not invoked:
                continue
            label = display.get(key, key)
            rows.append((wf.name, label, sorted(invoked)))
            for g in invoked:
                seen[g].append(label)

    rows.sort(key=lambda r: (r[0], r[1]))
    return rows, seen


def render(rows: list[tuple[str, str, list[str]]]) -> str:
    out = ["| Check | Workflow | Guards it runs |", "|---|---|---|"]
    for wf, label, invoked in rows:
        listed = "<br>".join(f"`{g}`" for g in invoked)
        out.append(f"| `{label}` | `{wf}` | {listed} |")
    out += [
        "",
        "Generated from `.github/workflows/*.yml`. A guard added to a job appears here "
        "automatically; one moved between jobs moves with it. What each guard means, and "
        "what to do when it fails, is the table above — that part is written by hand "
        "because it cannot be derived.",
        "",
    ]
    return "\n".join(out)


def main() -> int:
    # The rendered block and its diff carry `…` and `—`, and a Windows console defaults
    # to cp1252 where encoding either one raises. gen-capability-matrix.py hit exactly
    # this: --check correctly exited 1, but by crashing with a UnicodeEncodeError
    # instead of printing the diff. CI is UTF-8 and would never have shown it.
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    check = "--check" in sys.argv[1:]
    rows, seen = scan()

    # Refuse to emit an empty table. If the workflows' shape changes so this script
    # stops matching, an empty block looks exactly like "no guards run", which is a
    # far worse claim than a crash. Same guard as gen-capability-matrix.py's.
    if not rows:
        sys.exit("error: found no guard invocations in any workflow - the workflows' shape "
                 "changed and this script is now blind. Fix the parser rather than letting it "
                 "report an empty table.")

    orphans = sorted(g for g, where in seen.items() if not where)
    if orphans:
        sys.exit("error: these guards are in scripts/ but no workflow runs them: "
                 + ", ".join(orphans)
                 + "\n       Either wire the guard into a workflow, delete it, or add it to "
                   "NOT_A_GUARD in this script with a reason. A guard nothing runs is a guard "
                   "that proves nothing.")

    body = render(rows)

    if not PAGE.exists():
        sys.exit(f"error: {PAGE.name} does not exist.")

    text = PAGE.read_text(encoding="utf-8")
    if BEGIN not in text or END not in text:
        sys.exit(f"error: {PAGE.name} is missing the BEGIN/END GENERATED markers. Add them "
                 f"where the inventory should appear.")

    head, rest = text.split(BEGIN, 1)
    _, tail = rest.split(END, 1)
    updated = f"{head}{BEGIN}\n\n{body}\n{END}{tail}"

    print(f"{len(rows)} checks run {len(seen)} guards")

    if updated == text:
        print("guard inventory is up to date")
        return 0

    if check:
        diff = difflib.unified_diff(
            text.splitlines(keepends=True), updated.splitlines(keepends=True),
            fromfile="committed", tofile="derived from the workflows")
        sys.stdout.writelines(diff)
        print("\n::error::CONTRIBUTING.md's guard inventory no longer matches the workflows. "
              "Run `python scripts/gen-guard-inventory.py` and commit the result.")
        return 1

    PAGE.write_text(updated, encoding="utf-8", newline="\n")
    print(f"rewrote {PAGE.name}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
