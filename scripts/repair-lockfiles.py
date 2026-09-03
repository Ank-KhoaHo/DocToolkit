#!/usr/bin/env python3
"""Regenerate every lockfile under src/ and THIRD-PARTY-NOTICES.txt.

WHY THIS EXISTS. Dependabot's NuGet updater rewrites `packages.lock.json` with
`net8.0` ONLY, dropping `net10.0`. Measured 2026-08-13 across all three open
bumps at once - and PR #218 truncated BOTH packable projects even though it only
bumped TEST dependencies, so the blast radius is wider than a PR's title
suggests:

    main            ['net10.0', 'net8.0']
    dependabot PR   ['net8.0']

`premise-guard` then fails its `--locked-mode` restore with NU1004 and the PR
sits. That is the guard behaving correctly: a lockfile that no longer describes
the package this repo ships is exactly what locked mode exists to refuse.
Automerge was armed on all three and could not merge any of them.

The repair is always the same four commands, which is what this script is. It is
not Dependabot-specific - any time a lockfile needs rebuilding from the csproj,
this is the sequence, and doing it by hand is how the notices step gets forgotten
(a CI failure this repo has hit twice).

TWO TRAPS, both paid for once already.

1. NU1004's message has its labels THE WRONG WAY ROUND. It reports

       Lock file target frameworks: net8.0,net10.0. Project target frameworks net8.0.

   when the truth is the reverse - the PROJECT multi-targets and the LOCKFILE is
   the truncated one. Read literally it sends you hunting for a change to
   src/Directory.Build.props that does not exist.

2. A lockfile conflict is resolved by REGENERATING, never by merging hunks. Once
   one such PR merges the others conflict on packages.lock.json and
   THIRD-PARTY-NOTICES.txt; merge main in, then run this. Both are generated
   artefacts and hand-merging them produces a file that matches neither side.

Pushing to a Dependabot branch also stops Dependabot managing it, so it will not
rebase itself afterwards. That is the cost of the repair rather than a fault.

USAGE

    python scripts/repair-lockfiles.py            # repair
    python scripts/repair-lockfiles.py --check    # report, change nothing

`--check` is what CI runs: it exits non-zero if ANY lockfile under src/ is missing a
target framework the project builds, WITHOUT rewriting anything. Which projects those
are is read from disk - see projects() for why that is not a list.
"""
from __future__ import annotations

import json
import pathlib
import re
import subprocess
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent


def projects() -> list[pathlib.Path]:
    """Every project under src/, read from disk rather than listed here.

    IT WAS A TWO-ENTRY LIST, and it went stale the moment #353 split the library into per-concern
    projects. src/Directory.Build.props sets <RestorePackagesWithLockFile>, so all EIGHT projects
    under src/ commit a lockfile - and `--check` was reading two of them while printing
    "Every lockfile covers every target framework" for the other six.

    Measured 2026-09-03 (C41), not inferred: src/DocToolkit.Docx/packages.lock.json was truncated
    to net8.0 - the exact Dependabot failure mode this docstring opens with - and this script
    exited 0. The `--locked-mode` step two lines later in ci.yml then failed instead, with NU1004
    and its labels the wrong way round (Trap 1). So the truncation could never reach a release;
    what was lost was the diagnostic that names the project, the framework and the repair.

    Derived for exactly the reason target_frameworks() below is derived, and the sentence there
    turned out to describe the list twenty lines above it: a list kept in a script goes stale
    silently, and staleness in THIS direction means the check passes while a framework is missing.

    Globs the CSPROJ rather than the lockfile on purpose - a project whose packages.lock.json is
    missing entirely is a failure this script has always reported, and globbing lockfiles would
    make it invisible.
    """
    found = sorted((ROOT / "src").glob("*/*.csproj"))
    if not found:
        sys.exit("error: no project under src/ - a run over nothing would report success")
    return found


def target_frameworks() -> list[str]:
    """The TFMs the packable projects build, read from src/Directory.Build.props.

    Derived rather than hard-coded here for the same reason
    gen-third-party-notices.py reads the lockfile and automerge-eligible.py reads
    the workflows: a list kept in a script goes stale silently, and staleness in
    THIS direction means the check passes while a framework is missing.
    """
    props = (ROOT / "src" / "Directory.Build.props").read_text(encoding="utf-8")
    match = re.search(r"<TargetFrameworks>([^<]+)</TargetFrameworks>", props)
    if not match:
        sys.exit("error: no <TargetFrameworks> in src/Directory.Build.props")
    return [tfm.strip() for tfm in match.group(1).split(";") if tfm.strip()]


def lockfile_frameworks(project: pathlib.Path) -> list[str] | None:
    lock = project.parent / "packages.lock.json"
    if not lock.exists():
        return None
    return list(json.loads(lock.read_text(encoding="utf-8"))["dependencies"].keys())


def run(*args: str) -> None:
    result = subprocess.run(args, cwd=ROOT)
    if result.returncode != 0:
        sys.exit(f"error: {' '.join(args)} failed with {result.returncode}")


def main() -> int:
    check_only = "--check" in sys.argv[1:]
    expected = set(target_frameworks())
    all_projects = projects()
    print(f"projects build: {', '.join(sorted(expected))}")
    print(f"checking {len(all_projects)} projects under src/\n")

    broken = []
    for project in all_projects:
        found = lockfile_frameworks(project)
        name = project.parent.name
        if found is None:
            print(f"  {name}: no packages.lock.json")
            broken.append(name)
            continue

        missing = expected - set(found)
        status = "OK" if not missing else f"MISSING {', '.join(sorted(missing))}"
        print(f"  {name}: [{', '.join(sorted(found))}] {status}")
        if missing:
            broken.append(name)

    if check_only:
        if broken:
            print(
                "\n::error::A lockfile is missing a target framework the project builds. "
                "Run `python scripts/repair-lockfiles.py` and commit the result. "
                "Dependabot's NuGet updater does this; see the script's docstring."
            )
            return 1
        print("\nEvery lockfile covers every target framework.")
        return 0

    if not broken:
        print("\nNothing to repair - regenerating anyway, since you asked.")

    # RESTORING EACH ONE IS REDUNDANT AND DELIBERATE. Restoring a project restores its whole
    # project graph, so `dotnet restore src/DocToolkit/DocToolkit.csproj` alone already rewrites
    # seven of the eight lockfiles - measured, and it is why ci.yml's --locked-mode step catches
    # a truncation `--check` used to miss. Naming the two roots instead would be faster and would
    # reintroduce exactly the assumption that went stale here: that the graph reaches everything
    # committed. It does today. The check above is what tells you when it stops.
    print()
    for project in all_projects:
        print(f"restoring {project.parent.name} ...")
        run("dotnet", "restore", str(project), "--force-evaluate")

    # The resolved graph moved, so the notices must follow. Forgetting this is a
    # CI failure this repository has hit twice.
    print("\nregenerating third-party notices ...")
    run(sys.executable, str(ROOT / "scripts" / "gen-third-party-notices.py"))

    print()
    for project in all_projects:
        found = lockfile_frameworks(project) or []
        print(f"  {project.parent.name}: [{', '.join(sorted(found))}]")

    still_broken = [
        p.parent.name for p in all_projects
        if expected - set(lockfile_frameworks(p) or [])
    ]
    if still_broken:
        print(f"\n::error::still missing a framework after repair: {', '.join(still_broken)}")
        return 1

    print("\nRepaired. Commit packages.lock.json and THIRD-PARTY-NOTICES.txt together.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
