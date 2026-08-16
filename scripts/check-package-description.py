#!/usr/bin/env python3
"""Fail when a package `<Description>` makes a claim the approved API contradicts.

WHY THIS EXISTS. The release-curation checklist asks a human "does the nuget.org
Description still describe what ships?", and says outright that CI proves the field is
non-empty while *nothing can prove it is true*. That is correct in general and it left
a real hole, which 0.28.0 walked into twice at once:

  - the CORE description omitted legacy .doc conversion and password protection, both
    shipped in that release;
  - the EXTENSIONS description said "six injectable interfaces" and listed those six by
    name, while FIFTEEN shipped.

The second one is the same drift `check-readme-coverage.py` was written for - its own
docstring records the extensions README saying "six" after four more interfaces shipped -
one field further out, where no check was looking.

WHAT THIS CAN AND CANNOT PROVE. It cannot prove a description is COMPLETE or well
written; no check can. It proves the description contains no FALSIFIABLE claim that is
already false, which is a strictly weaker property and happens to be the one that broke:

  1. every format the approved API ships is named in the core description;
  2. no description names an `I…` interface that does not exist;
  3. no description states an interface count that disagrees with the approved API.

So the checklist item stays human - the wording, the ordering, whether it reads well -
and the part a machine can settle is settled here instead of being re-read every release.
"""

import io
import re
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

CORE_CSPROJ = "src/DocToolkit/DocToolkit.csproj"
EXT_CSPROJ = ("src/DocToolkit.Extensions.DependencyInjection/"
              "DocToolkit.Extensions.DependencyInjection.csproj")
CORE_API = "tests/DocToolkit.Tests/PublicApi/DocToolkit.approved.txt"
EXT_API = ("tests/DocToolkit.Extensions.DependencyInjection.Tests/PublicApi/"
           "DocToolkit.Extensions.DependencyInjection.approved.txt")

# Same map as gen-capability-matrix.py: the fragment in a type name is not what a reader
# calls the format. Kept in step with that file by hand is exactly what this repo avoids,
# so anything missing here is reported rather than silently skipped - see UNKNOWN below.
DISPLAY = {
    "Docx": "DOCX", "Pdf": "PDF", "Html": "HTML", "Xlsx": "XLSX",
    "Pptx": "PPTX", "Csv": "CSV", "Markdown": "Markdown", "Doc": "DOC",
}

WORDS = {"one": 1, "two": 2, "three": 3, "four": 4, "five": 5, "six": 6, "seven": 7,
         "eight": 8, "nine": 9, "ten": 10, "eleven": 11, "twelve": 12, "thirteen": 13,
         "fourteen": 14, "fifteen": 15, "sixteen": 16, "seventeen": 17, "eighteen": 18,
         "nineteen": 19, "twenty": 20}


def read(path):
    return io.open(path, encoding="utf-8").read()


def description(csproj):
    m = re.search(r"<Description>(.*?)</Description>", read(csproj), re.S)
    if not m:
        sys.exit(f"::error::{csproj} has no <Description>.")
    text = " ".join(m.group(1).split())
    if not text:
        sys.exit(f"::error::{csproj} has an empty <Description>.")
    return text


def formats(api_text):
    """Every format the approved API converts between, plus the editors' own."""
    found, unknown = set(), set()
    for a, b in re.findall(r"class (\w+?)To(\w+?)Converter", api_text):
        for part in (a, b):
            (found if part in DISPLAY else unknown).add(DISPLAY.get(part, part))
    for editor, fmt in (("DocxEditor", "DOCX"), ("WorkbookEditor", "XLSX"),
                        ("PresentationEditor", "PPTX"), ("PdfEditor", "PDF")):
        if f"class {editor}" in api_text:
            found.add(fmt)
    return found, unknown


def names(text):
    return set(re.findall(r"\bI[A-Z][A-Za-z]+\b", text))


def main():
    errors = []

    core, ext = description(CORE_CSPROJ), description(EXT_CSPROJ)
    core_api, ext_api = read(CORE_API), read(EXT_API)

    shipped, unknown = formats(core_api)
    if not shipped:
        # A parser that matches nothing would pass every check below, which is the
        # vacuously-green failure this repo has already had to fix twice.
        sys.exit(f"::error::No formats parsed from {CORE_API}; the checks below would be vacuous.")
    if unknown:
        errors.append(f"{CORE_CSPROJ}: DISPLAY in this script has no entry for "
                      f"{', '.join(sorted(unknown))} — add it, or the format is unchecked.")

    # 1. Every shipped format named in the core description. Word boundaries matter:
    #    "DOC" is a substring of "DOCX", so a naive search passes on DOCX alone.
    for fmt in sorted(shipped):
        if not re.search(rf"(?<![A-Za-z]){re.escape(fmt)}(?![A-Za-z])", core, re.I):
            errors.append(f"{CORE_CSPROJ}: ships {fmt} but the <Description> never names it. "
                          f"That is the page a consumer reads first.")

    # 2 and 3, for both packages.
    for label, text, api in ((CORE_CSPROJ, core, core_api), (EXT_CSPROJ, ext, ext_api)):
        real = set(re.findall(r"public interface (I\w+)", api))
        for claimed in sorted(names(text) - real):
            errors.append(f"{label}: <Description> names `{claimed}`, which is not in the "
                          f"approved API for this package.")

        for m in re.finditer(r"\b([a-z]+|\d+)\s+(?:injectable\s+)?interfaces\b", text, re.I):
            token = m.group(1).lower()
            claimed = int(token) if token.isdigit() else WORDS.get(token)
            if claimed is not None and claimed != len(real):
                errors.append(f"{label}: <Description> claims {claimed} interfaces; the approved "
                              f"API has {len(real)}. Prefer describing the shape to counting.")

    if errors:
        for e in errors:
            print(f"::error::{e}")
        print("\nA description is the first thing a consumer reads and cannot be edited once the "
              "version is published. Fix the description, not this check.")
        return 1

    print(f"core <Description> names all {len(shipped)} shipped formats "
          f"({', '.join(sorted(shipped))}); no stale interface names or counts in either package.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
