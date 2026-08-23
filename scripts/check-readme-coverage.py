#!/usr/bin/env python3
"""Fail when a shipped public type is not named in the README its package publishes.

There are three READMEs - the repository root, and one per package - and they were
kept in sync by hand. On 2026-08-08 that failed six times in a single day: Known
limitations, Telemetry, page setup, macOS/arm64 support, XLSX/PPTX->PDF and
DOCX->HTML/Markdown were all added to the ROOT readme only, while
`src/DocToolkit/README.md` - the one nuget.org actually renders, via
PackageReadmeFile - went untouched. The package page advertised a smaller,
Linux-only library than the one that shipped. Nothing detected it; it surfaced by
accident, while diffing the two files for an unrelated change.

Then, one day later and with that fresh in mind, the same thing happened again to
the extensions README: four new interfaces shipped and it still said "six".

So the check exists because hand-sync has a measured failure rate, not because it
might theoretically drift.

WHAT IT CHECKS, and why this shape. The public types come from the APPROVED API
files, which are already this repository's reviewed source of truth for what
ships - so there is no second list to maintain and nothing to forget to update.
If a type is in the approved surface, a consumer can use it, and the page they
land on should at least name it.

It deliberately does NOT require the three files to match each other. They
legitimately differ: the root carries badges and repository layout, the packages
carry install instructions for themselves. Requiring equality would produce
false positives forever and be switched off within a month.
"""

import io
import pathlib
import re
import sys

REPO = pathlib.Path(__file__).resolve().parent.parent

# EVERY approved file the core package ships, not just the core assembly's.
#
# Ank.DocToolkit packs SIX assemblies into one nupkg since the per-concern project split, each
# with its own approved file. This read `DocToolkit.approved.txt` alone until 2026-08-23, so
# DocxEditor, PdfEditor, WorkbookEditor, PresentationEditor and every DTO beside them had stopped
# being checked - a guard that exists BECAUSE this README went stale six times, quietly narrowed
# to a sixth of the surface it is supposed to watch, while still reporting green.
#
# Derived by glob rather than listed, for the same reason gen-capability-matrix.py globs: a
# hand-maintained list is exactly what went stale, and a seventh project would rejoin the check
# by existing rather than by somebody remembering.
CORE_APPROVED = sorted(
    p for p in (REPO / "tests" / "DocToolkit.Tests" / "PublicApi").glob("DocToolkit*.approved.txt")
)

EXTENSIONS_APPROVED = [
    REPO / "tests" / "DocToolkit.Extensions.DependencyInjection.Tests" / "PublicApi"
    / "DocToolkit.Extensions.DependencyInjection.approved.txt"
]

# (approved API files, the README that package publishes, label)
TARGETS = [
    (CORE_APPROVED, "src/DocToolkit/README.md", "Ank.DocToolkit"),
    (EXTENSIONS_APPROVED,
     "src/DocToolkit.Extensions.DependencyInjection/README.md",
     "Ank.DocToolkit.Extensions.DependencyInjection"),
]

# Types a consumer never writes by name. Keep this short and argue for additions -
# every entry is a hole in the check.
EXEMPT = {
    # Reached as services.AddDocToolkit(); nobody types the class name.
    "ServiceCollectionExtensions",
}

TYPE = re.compile(r"public (?:static |sealed |abstract |readonly )*(?:class|interface|enum|struct) (\w+)")


def main():
    failures = []

    for approved, readme, package in TARGETS:
        if not approved:
            sys.exit(f"::error::No approved API files found for {package}; this check cannot run. "
                     "A glob that matches nothing would pass vacuously.")
        try:
            surface = chr(10).join(io.open(a, encoding="utf-8").read() for a in approved)
            doc = io.open(readme, encoding="utf-8").read()
        except FileNotFoundError as missing:
            sys.exit(f"::error::{missing.filename} not found; this check cannot run.")

        types = sorted(set(TYPE.findall(surface)) - EXEMPT)
        if not types:
            names = ", ".join(str(a) for a in approved)
            sys.exit(f"::error::No public types parsed from {names}. The check would "
                     "pass vacuously, which is worse than failing.")

        undocumented = [t for t in types if t not in doc]
        print(f"{package}: {len(types)} public types across {len(approved)} assembl"
              f"{'y' if len(approved) == 1 else 'ies'}, "
              f"{len(types) - len(undocumented)} named in {readme}")

        for name in undocumented:
            failures.append(f"{package}: `{name}` ships but is not named in {readme}")

    print()
    if not failures:
        print("Every shipped type is named in the README its package publishes.")
        return 0

    for failure in failures:
        print(f"::error::{failure}")
    print()
    print("A type in the approved API is a type a consumer can use. The page they land on "
          "should name it. Add it to that README - and remember the ROOT readme is NOT what "
          "nuget.org renders.")
    return 1


if __name__ == "__main__":
    sys.exit(main())
