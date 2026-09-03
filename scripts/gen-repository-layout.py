#!/usr/bin/env python3
"""Generate CONTRIBUTING.md's src/ inventory from the project files themselves.

WHY THIS IS GENERATED. `CONTRIBUTING.md`'s *Repository layout* named **two** projects under `src/`
while eight were committed. The other six arrived with #353 on 2026-08-23, which split the library
into per-concern projects, and the hand-written block did not follow - so for the whole of that
period the only public map of this repository pointed a contributor at `src/DocToolkit/`, which
holds 18 converter files and no `DocxEditor.cs`. Filed as D41.

Same principle as `gen-guard-inventory.py` reading the workflows, `gen-third-party-notices.py`
reading the lockfile and `automerge-eligible.py` reading the workflows: a list kept by hand goes
stale silently, and staleness here is invisible to every check this repository has, because a
prose paragraph is not wrong in any way a build can see.

WHAT IS DERIVED AND WHAT IS NOT, which is the line worth holding.

Everything inside the markers is read off disk: which projects exist, which are packable and under
what package id, which are packed into another package, and which project is the shared floor the
others reference. Nothing here is a description, because a description cannot be derived and a
generator that carried one would just be a hand-maintained list with extra steps.

The PROSE ABOVE AND BELOW THE BLOCK is hand-written and stays that way. It says why the split
exists and what `tests/`, `samples/` and `docfx/` are for - none of which is on disk in any form a
script can read. The machine says what is there; the person says why.

USAGE

    python scripts/gen-repository-layout.py            # rewrite the block
    python scripts/gen-repository-layout.py --check    # report, change nothing

`--check` is what CI runs, in the `formatting` job beside `gen-guard-inventory.py --check`. It
exits non-zero when the block no longer matches `src/`, and the fix is in the message.
"""
from __future__ import annotations

import pathlib
import re
import sys

REPO = pathlib.Path(__file__).resolve().parent.parent
SRC = REPO / "src"
PAGE = REPO / "CONTRIBUTING.md"

BEGIN = "<!-- BEGIN GENERATED (scripts/gen-repository-layout.py) - do not edit by hand -->"
END = "<!-- END GENERATED (scripts/gen-repository-layout.py) -->"

PACKAGE_ID = re.compile(r"<PackageId>\s*([^<\s]+)\s*</PackageId>")
IS_PACKABLE = re.compile(r"<IsPackable>\s*(\w+)\s*</IsPackable>", re.I)
PROJECT_REF = re.compile(r'<ProjectReference\s+Include="([^"]+)"')

# The path column, wide enough for the longest name plus a space. Recomputed per run rather than
# fixed: DocToolkit.Extensions.DependencyInjection is 46 characters today and a longer one would
# silently produce a ragged block against a hard-coded width.
GAP = 2


class Project:
    def __init__(self, csproj: pathlib.Path) -> None:
        text = csproj.read_text(encoding="utf-8")
        self.name = csproj.parent.name
        self.slug = f"src/{self.name}/"

        pid = PACKAGE_ID.search(text)
        packable = IS_PACKABLE.search(text)

        # A project is packable unless it says otherwise - that is the SDK's own default, so
        # reading the absence of <IsPackable> as "packable" matches what dotnet pack does rather
        # than guessing. A packable project without a <PackageId> would pack under its assembly
        # name; refuse rather than print a name nothing declares.
        self.packable = not (packable and packable.group(1).lower() == "false")
        self.package_id = pid.group(1) if pid else None

        self.references = sorted(
            pathlib.PurePath(m.replace("\\", "/")).stem
            for m in PROJECT_REF.findall(text)
        )


def load() -> list[Project]:
    found = sorted(SRC.glob("*/*.csproj"))
    if not found:
        sys.exit("error: no project under src/ - a block generated from nothing would look "
                 "exactly like a repository with no source in it")
    return [Project(p) for p in found]


def render(projects: list[Project]) -> str:
    packable = [p for p in projects if p.packable]
    packed_in = [p for p in projects if not p.packable]

    for p in packable:
        if not p.package_id:
            sys.exit(f"error: {p.slug} is packable but declares no <PackageId>. Add one, or set "
                     "<IsPackable>false</IsPackable> if it is packed into another package.")

    # Which package swallows each non-packable project. Derived from the reference graph rather
    # than assumed to be the core package: a second packable project could pack its own.
    packs: dict[str, list[str]] = {p.name: [] for p in packed_in}
    for host in packable:
        for ref in host.references:
            if ref in packs:
                packs[ref].append(host.package_id or host.name)

    orphans = [name for name, hosts in packs.items() if not hosts]
    if orphans:
        sys.exit(f"error: {', '.join(sorted(orphans))} is not packable and no packable project "
                 "references it, so it ships nowhere. Either reference it or delete it.")

    # Who leans on whom, AMONG THE PACKED-IN PROJECTS ONLY. This is the line a contributor
    # actually needs: it says which project is the shared floor without anybody writing that down.
    #
    # The packing host is deliberately excluded. Every packed-in project is referenced by the
    # project that packs it - that is WHY it is packed - so listing it restates the column to its
    # left on every row, and the one row that means something (Primitives, which the other five
    # also use) stops standing out. Noise that appears on every row hides the signal on one.
    depended_on: dict[str, list[str]] = {p.name: [] for p in packed_in}
    for p in packed_in:
        for ref in p.references:
            if ref in depended_on and ref != p.name:
                depended_on[ref].append(p.name.replace("DocToolkit.", ""))

    width = max(len(p.slug) for p in projects) + GAP
    lines = ["```"]

    lines.append("PUBLISHED to nuget.org")
    for p in packable:
        lines.append(f"  {p.slug.ljust(width)}{p.package_id}")

    lines.append("")
    lines.append("PACKED IN - built into a published package, never published alone")
    for p in packed_in:
        hosts = ", ".join(sorted(set(packs[p.name])))
        note = f"into {hosts}"
        users = sorted(set(depended_on[p.name]))
        if users:
            note += f"; referenced by {', '.join(users)}"
        lines.append(f"  {p.slug.ljust(width)}{note}")

    lines.append("```")
    return "\n".join(lines)


def main() -> int:
    check_only = "--check" in sys.argv[1:]
    body = render(load())

    if not PAGE.exists():
        sys.exit(f"error: {PAGE.name} does not exist.")

    text = PAGE.read_text(encoding="utf-8")
    if BEGIN not in text or END not in text:
        sys.exit(f"error: {PAGE.name} is missing the BEGIN/END GENERATED markers for this "
                 f"script. Add them around the src/ block in Repository layout:\n"
                 f"  {BEGIN}\n  {END}")

    head, rest = text.split(BEGIN, 1)
    _, tail = rest.split(END, 1)
    updated = f"{head}{BEGIN}\n\n{body}\n\n{END}{tail}"

    if updated == text:
        print(f"{PAGE.name}'s src/ inventory is up to date")
        return 0

    if check_only:
        print(body)
        print(f"\n::error::{PAGE.name}'s Repository layout no longer matches src/. "
              "Run `python scripts/gen-repository-layout.py` and commit the result.")
        return 1

    PAGE.write_text(updated, encoding="utf-8")
    print(f"{PAGE.name}'s src/ inventory rewritten from {len(load())} projects")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
