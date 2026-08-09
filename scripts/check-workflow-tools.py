#!/usr/bin/env python3
"""Fail when a workflow invokes a local dotnet tool without restoring the manifest first.

A tool in .config/dotnet-tools.json does not exist on a runner until
`dotnet tool restore` has run. Miss that step and the job fails with:

    Run "dotnet tool restore" to make the "dotnet-CycloneDX" command available.

Which is exactly what happened on 2026-08-09. release.yml gained an SBOM step
using CycloneDX and did not gain a restore, because until then it used no local
tool at all - ci.yml, docs.yml and mutation.yml each already had one, so the
pattern existed and simply was not followed. BOTH v0.17.0 release runs failed and
that version never reached nuget.org.

It survived review because the command was dry-run locally, on a machine where
the tool was already installed. That is the failure this check exists for: the
difference between "works here" and "works on a runner" is invisible to the
person who has already installed it.

DERIVED, NOT LISTED. The tool commands come from .config/dotnet-tools.json, so
adding a tool to the manifest automatically brings it under this check. A list
kept here would go stale the same way the missing step did.

Note `dotnet stryker` invokes the tool whose command is `dotnet-stryker`: the
dotnet driver accepts both spellings, so both are treated as an invocation.
"""

import glob
import io
import json
import re
import sys

MANIFEST = ".config/dotnet-tools.json"
WORKFLOWS = ".github/workflows/*.yml"
RESTORE = "dotnet tool restore"


def commands():
    """Every command name the local manifest provides, plus its `dotnet x` spelling."""
    manifest = json.load(io.open(MANIFEST, encoding="utf-8"))
    names = set()
    for tool in manifest.get("tools", {}).values():
        for command in tool.get("commands", []):
            names.add(command)
            if command.startswith("dotnet-"):
                names.add(command[len("dotnet-"):])
    return sorted(names)


def main():
    try:
        available = commands()
    except FileNotFoundError:
        sys.exit(f"::error::{MANIFEST} not found; this check cannot run.")

    if not available:
        sys.exit(f"::error::No tool commands parsed from {MANIFEST}. The check would pass "
                 "vacuously, which is worse than failing.")

    print(f"{len(available)} local tool command(s): {', '.join(available)}")
    print()

    failures = []
    checked = 0

    for workflow in sorted(glob.glob(WORKFLOWS)):
        raw = io.open(workflow, encoding="utf-8").read()
        name = workflow.replace("\\", "/").split("/")[-1]

        # Comment lines are stripped before matching, and that is not tidiness. The first
        # version of this check searched the whole file, and the comment explaining the
        # v0.17.0 failure QUOTES the missing command - so deleting the real step still left
        # the string present and the check went green on a file it should have rejected.
        # Caught by running the red case; it would otherwise have shipped as decoration.
        text = "\n".join(line for line in raw.splitlines()
                         if not line.lstrip().startswith("#"))

        used = [c for c in available
                if re.search(r"dotnet (?:tool run )?" + re.escape(c) + r"\b", text)]
        if not used:
            continue

        checked += 1
        restores = RESTORE in text
        print(f"  {name:24} uses {', '.join(used):22} restore={'yes' if restores else 'NO'}")

        if not restores:
            failures.append(
                f"{name} invokes the local tool(s) {', '.join(used)} but never runs "
                f"`{RESTORE}`. The command does not exist on a runner until it does.")

    if checked == 0:
        sys.exit("::error::No workflow appears to invoke a local tool. Either they stopped, or "
                 "this check stopped matching them; both are worth a red build.")

    print()
    if not failures:
        print("Every workflow using a local tool restores the manifest first.")
        return 0

    for failure in failures:
        print(f"::error::{failure}")
    print()
    print("Add a `dotnet tool restore` step before the one that uses the tool. This is not "
          "reproducible locally on a machine where the tool is already installed - which is how "
          "it reached main the first time.")
    return 1


if __name__ == "__main__":
    sys.exit(main())
