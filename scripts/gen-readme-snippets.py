#!/usr/bin/env python3
"""Inject compiled README snippets from ReadmeExamples.cs into the three READMEs.

WHY. The DocFX guides pull every sample from a compiled test, because "a snippet retyped
into a comment is a claim nothing compiles" - and that mechanism caught two private
constructors on its first run. The READMEs were exempt: 28 code blocks, none verified, and
src/DocToolkit/README.md is the one nuget.org renders, so its snippets are the first code a
consumer copies.

check-readme-coverage.py already guards these files, but it asserts only that every shipped
TYPE is named. Coverage, not correctness. It cannot tell a working example from a wrong one.

Five blocks appeared in more than one README, kept in sync by hand - the same hand-sync that
failed six times in one day. A shared region removes that.

TWO SOURCES, not one. The extensions package's own snippets (services.AddDocToolkit(),
IHtmlToPdfConverter and friends) cannot live in tests/DocToolkit.Tests: that project holds a
ProjectReference to src/DocToolkit, while tests/DocToolkit.Extensions.DependencyInjection.Tests
references Ank.DocToolkit as a published PackageReference - deliberately, so a DI-flavoured
snippet compiling there proves it against what a consumer's restore actually gets, not against
whatever is on main. A region name must be unique ACROSS both sources: the same name defined in
each would leave it ambiguous which body a README marker means, which is worse than the region
being missing outright.

USAGE
    python scripts/gen-readme-snippets.py            # write
    python scripts/gen-readme-snippets.py --check    # verify, change nothing
"""
from __future__ import annotations

import difflib
import pathlib
import re
import sys

REPO = pathlib.Path(__file__).resolve().parent.parent
SOURCES = [
    REPO / "tests" / "DocToolkit.Tests" / "ReadmeExamples.cs",
    REPO / "tests" / "DocToolkit.Extensions.DependencyInjection.Tests" / "ReadmeExamples.cs",
]
READMES = [
    REPO / "README.md",
    REPO / "src" / "DocToolkit" / "README.md",
    REPO / "src" / "DocToolkit.Extensions.DependencyInjection" / "README.md",
]

BEGIN = "<!-- BEGIN SNIPPET: {} -->"
END = "<!-- END SNIPPET -->"

# Blocks that CANNOT be compiled, each with the reason. An unexplained exclusion list is how
# a scope decision quietly becomes an oversight.
EXCLUDED = {
    "before/after migration blocks": "they show two VERSIONS of the output, only one of which exists",
    "OpenTelemetry wiring": "OpenTelemetry is deliberately not a dependency; compiling it would mean adding one to prove a doc",
    "dotnet add package": "not C#",
    "readme-di-stream (Stream overload in a minimal-API handler)": "it is an ASP.NET Core "
        "minimal-API handler, and the shared framework is deliberately not referenced by a "
        "test project - compiling it would mean taking a dependency to prove a doc",
}


def regions() -> dict[str, str]:
    found: dict[str, str] = {}
    origin: dict[str, pathlib.Path] = {}
    for source in SOURCES:
        text = source.read_text(encoding="utf-8")
        for m in re.finditer(r"^[ \t]*#region (readme-[\w-]+)\r?\n(.*?)^[ \t]*#endregion",
                             text, re.S | re.M):
            name, body = m.group(1), m.group(2)
            if name in found:
                sys.exit(f"error: region {name} is defined in both "
                         f"{origin[name].relative_to(REPO).as_posix()} and "
                         f"{source.relative_to(REPO).as_posix()} - an ambiguous source is worse "
                         f"than a missing one.")
            lines = [ln for ln in body.split("\n") if ln.strip()]
            if not lines:
                sys.exit(f"error: region {name} is empty. A blank snippet renders as a short "
                         f"example rather than a broken one.")
            indent = min(len(ln) - len(ln.lstrip()) for ln in lines)
            found[name] = "\n".join(ln[indent:].rstrip() for ln in body.split("\n")).strip()
            origin[name] = source
    return found


def main() -> int:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    check = "--check" in sys.argv[1:]

    found = regions()
    if not found:
        sys.exit("error: no readme-* regions found in either ReadmeExamples.cs - the region "
                 "format changed and this script is now blind. Fix the parser rather than "
                 "letting it silently inject nothing.")

    failures, used = [], set()
    for path in READMES:
        text = original = path.read_text(encoding="utf-8")
        for name, body in found.items():
            begin = BEGIN.format(name)
            if begin not in text:
                continue
            used.add(name)
            pattern = re.compile(
                re.escape(begin) + r".*?" + re.escape(END), re.S)
            replacement = f"{begin}\n\n```csharp\n{body}\n```\n\n{END}"
            text = pattern.sub(lambda _: replacement, text)

        if text == original:
            continue
        if check:
            failures.append((path, original, text))
        else:
            path.write_text(text, encoding="utf-8", newline="\n")
            print(f"rewrote {path.relative_to(REPO).as_posix()}")

    orphans = sorted(set(found) - used)
    if orphans:
        sys.exit("error: these regions exist in ReadmeExamples.cs but no README references "
                 "them: " + ", ".join(orphans) + "\n       Either add the marker or delete "
                 "the region - a snippet nothing renders is a test pretending to be a doc.")

    print(f"{len(found)} regions, {len(used)} referenced")

    if failures:
        for path, before, after in failures:
            sys.stdout.writelines(difflib.unified_diff(
                before.splitlines(keepends=True), after.splitlines(keepends=True),
                fromfile=f"committed {path.name}", tofile="derived from ReadmeExamples.cs"))
        print("\n::error::A README snippet no longer matches ReadmeExamples.cs. "
              "Run `python scripts/gen-readme-snippets.py` and commit the result.")
        return 1

    if check:
        print("README snippets are up to date")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
