#!/usr/bin/env python3
"""Assert every file `stryker-config.json` says to mutate was actually mutated.

WHY THIS EXISTS. On 2026-08-24 the mutation suite was found covering **3 of its 15 intended
files**, and it had been passing weekly in that state.

`stryker-config.json` set `project: DocToolkit.csproj`, so Stryker mutated that one project. The
per-concern split had moved 11 of the scoped files into sibling projects, where Stryker could not
see them - `GuardedResourceLoader` to DocToolkit.Html, `PageSetup` and `StreamPipeline` to
DocToolkit.Primitives, and so on. A 16th pattern, `**/OfflineResourceLoader.cs`, had never matched
anything: that class is a `private sealed class` inside `HtmlToDocxConverter.cs`, so it has no file
of its own, and **Stryker silently ignores a mutate pattern that matches nothing.**

    intended   575 mutants tested, 96.55%
    actual      74 mutants tested, 98.65%     <- and the gate is break: 95

THE SCORE WENT UP. That is the whole hazard: a mutation score means nothing without its
denominator, and nothing reported the denominator. Three weekly runs passed while 87% of the
intended surface went unmutated - including `GuardedResourceLoader`, the only network path in the
library, which `CLAUDE.md` says this suite exists to protect.

WHAT THIS CHECKS. Every pattern in the config's `mutate` list matched at least one file that
Stryker actually **tested** - not merely saw. A file Stryker analysed but filtered out has mutants
recorded as `Ignored`, so presence in the report is not enough; the check requires a real verdict.

That catches both ways the scope silently shrank:

  * a pattern matching no file at all           (OfflineResourceLoader.cs)
  * a file present on disk but outside the run  (the 11 the split moved)

    python scripts/check-mutation-scope.py                       # newest StrykerOutput run
    python scripts/check-mutation-scope.py <report.json>
    python scripts/check-mutation-scope.py --self-test
"""

import glob
import io
import json
import os
import pathlib
import sys

REPO = pathlib.Path(__file__).resolve().parent.parent
CONFIG = REPO / "stryker-config.json"

# A mutant Stryker never ran tells us nothing about coverage. `Ignored` is the filter's own
# verdict, and `CompileError` means the mutant could not be built - neither proves the file was
# in scope. Anything else means Stryker ran a test against it.
TESTED = {"Killed", "Survived", "Timeout", "NoCoverage"}


def patterns(config_text):
    """The mutate patterns, exclusions dropped - a `!` entry removes rather than requires."""
    doc = json.loads(config_text)
    section = doc.get("stryker-config", doc)
    return [p for p in section.get("mutate", []) if not p.startswith("!")]


def newest_report():
    hits = sorted(glob.glob(str(REPO / "StrykerOutput" / "*" / "reports" / "mutation-report.json")))
    if not hits:
        sys.exit("::error::No StrykerOutput report found. Run `dotnet stryker` first, or pass a "
                 "report path. This check cannot confirm the scope without one, and must not "
                 "pass while unable to.")
    return max(hits, key=os.path.getmtime)


def tested_files(report_text):
    """{basename: number of mutants Stryker actually ran} for the run."""
    doc = json.loads(report_text)
    out = {}
    for path, entry in (doc.get("files") or {}).items():
        ran = sum(1 for m in (entry.get("mutants") or []) if m.get("status") in TESTED)
        if ran:
            out[os.path.basename(path.replace("\\", "/"))] = ran
    return out


def check(config_text, report_text):
    """Returns (failures, matched) - failures is empty when every pattern was mutated."""
    ran = tested_files(report_text)
    failures, matched = [], {}

    for pattern in patterns(config_text):
        name = os.path.basename(pattern.replace("\\", "/"))
        if name in ran:
            matched[name] = ran[name]
            continue

        on_disk = glob.glob(str(REPO / "src" / "**" / name), recursive=True)
        if on_disk:
            where = os.path.dirname(on_disk[0]).replace("\\", "/").split("/")[-1]
            failures.append(
                f"{pattern} matched {on_disk[0]} but Stryker tested no mutant in it. That file is "
                f"in project {where}; if the run is scoped to one project it cannot see it.")
        else:
            failures.append(
                f"{pattern} matches no file under src/ at all. Stryker ignores an unmatched "
                "mutate pattern silently, so this entry has been doing nothing. If the code moved "
                "into another file, name that file; if it is a nested class, name its container.")

    return failures, matched


def self_test():
    """Controls, including the two real failure modes and the vacuity case."""
    cfg = json.dumps({"stryker-config": {"mutate": ["**/A.cs", "**/B.cs"]}})

    def report(entries):
        return json.dumps({"files": {p: {"mutants": [{"status": s} for s in st]}
                                     for p, st in entries.items()}})

    cases = [
        ("both files tested passes",
         report({"src/X/A.cs": ["Killed"], "src/Y/B.cs": ["Survived"]}), 0),
        # THE SPLIT CASE: analysed but every mutant filtered out.
        ("a file present but all-Ignored FAILS",
         report({"src/X/A.cs": ["Killed"], "src/Y/B.cs": ["Ignored", "Ignored"]}), 1),
        # THE DEAD-PATTERN CASE.
        ("a file absent from the report FAILS",
         report({"src/X/A.cs": ["Killed"]}), 1),
        ("CompileError alone does not count as tested",
         report({"src/X/A.cs": ["Killed"], "src/Y/B.cs": ["CompileError"]}), 1),
        ("an empty report FAILS both patterns",
         report({}), 2),
    ]

    bad = 0
    for name, rep, expected in cases:
        failures, _ = check(cfg, rep)
        ok = len(failures) == expected
        print(f"  {'ok  ' if ok else 'FAIL'} {name} -> {len(failures)} failure(s), expected {expected}")
        bad += 0 if ok else 1

    # And the real config, so a shape change is caught rather than assumed.
    real = patterns(io.open(CONFIG, encoding="utf-8").read())
    if len(real) < 2:
        print(f"  FAIL only {len(real)} mutate pattern(s) parsed from the real config; a "
              "derivation that finds nothing would pass every run")
        bad += 1
    else:
        print(f"  ok   {len(real)} mutate pattern(s) parsed from the real config")

    print()
    if bad:
        print(f"::error::self-test failed {bad} case(s)")
        return 1
    print("self-test passed, both real failure modes covered")
    return 0


def main(argv):
    if "--self-test" in argv:
        return self_test()

    paths = [a for a in argv if not a.startswith("-")]
    report_path = paths[0] if paths else newest_report()

    config_text = io.open(CONFIG, encoding="utf-8").read()
    report_text = io.open(report_path, encoding="utf-8").read()

    failures, matched = check(config_text, report_text)

    print(f"report: {report_path}")
    print(f"{len(patterns(config_text))} mutate pattern(s); {len(matched)} mutated, "
          f"{sum(matched.values())} mutant(s) tested across them")
    for name in sorted(matched):
        print(f"  OK  {name:<32} {matched[name]:>4} tested")

    if not failures:
        print()
        print("Every file the config says to mutate was actually mutated.")
        return 0

    print()
    for f in failures:
        print(f"::error::{f}")
    print()
    print("A mutation score means nothing without its denominator. The suite covered 3 of 15 files "
          "for weeks while the score ROSE, because nothing checked what was being mutated.")
    return 1


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
