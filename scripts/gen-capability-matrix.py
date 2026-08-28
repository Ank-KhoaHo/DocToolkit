#!/usr/bin/env python3
"""Generate the docs site's capability matrix from the approved public API.

WHY THIS IS GENERATED. The same table has been written by hand three times in
this repository - in README.md, in the enhancement backlog, and in the docs
guides - and every copy went stale. The backlog's was wrong in eight cells and
both of its conclusions were false; the docs landing page described a library
without Markdown in either direction, without PDF text extraction and without
XLSX export, five days after all three shipped. A capability table is exactly
the kind of claim nobody re-derives while adding a feature.

So it is derived, from the file this repository already treats as the reviewed
source of truth for what ships: tests/DocToolkit.Tests/PublicApi/*.approved.txt.
Same principle as gen-third-party-notices.py reading the lockfile,
automerge-eligible.py reading the workflows, check-readme-coverage.py reading
the approved API, and StreamOverloadTests reflecting over the assembly. Derive,
do not remember.

WHAT IT CAN HONESTLY DERIVE, and what it deliberately does not attempt.

The converter naming convention is strict and enforced by review: every
converter is `<From>To<To>Converter`. That is a fact about the name, not a guess
about behaviour, so the conversion grid is read straight off the class names.
All eleven converters follow it today.

The editing table is the same idea one level down: the four `*Editor` classes'
public method NAMES. It reports what exists, not what it does - `ReadTable` is
listed, its 0-based index is not. Prose belongs in the guides; this table's job
is to be complete and current, which prose has repeatedly failed to be.

It does NOT try to describe fidelity, options, or caveats. A generated table
that editorialised would be a worse version of the guides rather than an index
into them.

USAGE

    python scripts/gen-capability-matrix.py            # write
    python scripts/gen-capability-matrix.py --check    # verify, change nothing

`--check` is what CI runs. It exits non-zero when the committed page no longer
matches the approved API, and prints a diff.
"""
from __future__ import annotations

import difflib
import pathlib
import re
import sys

REPO = pathlib.Path(__file__).resolve().parent.parent
# EVERY approved file, not just the core one. The library packs several assemblies into one
# package, so the shipped surface is spread across DocToolkit.approved.txt,
# DocToolkit.Primitives.approved.txt, DocToolkit.Docx.approved.txt and their siblings. Reading one
# would silently describe a smaller library than ships - which is the exact drift this generated
# page exists to prevent, arriving through a different door.
APPROVED = sorted((REPO / "tests" / "DocToolkit.Tests" / "PublicApi").glob("DocToolkit*.approved.txt"))
PAGE = REPO / "docfx" / "guides" / "capabilities.md"

BEGIN = "<!-- BEGIN GENERATED (scripts/gen-capability-matrix.py) - do not edit by hand -->"
END = "<!-- END GENERATED -->"

# How a class-name fragment is spelled for a reader. Anything absent falls back to
# the fragment itself, so a new format shows up looking slightly wrong rather than
# silently vanishing - a visible gap beats a missing row.
DISPLAY = {
    "Docx": "DOCX", "Pdf": "PDF", "Html": "HTML", "Xlsx": "XLSX",
    "Pptx": "PPTX", "Csv": "CSV", "Markdown": "Markdown",
    # The legacy binary format. It must sort and read as DOC, not "Doc" - and the comment above
    # about a new format "showing up looking slightly wrong" is exactly what happened when
    # DocToDocxConverter landed, which is the design working as intended.
    "Doc": "DOC",
}

# The editors do not carry their format in the name the way converters do.
EDITOR_FORMAT = {
    "DocxEditor": "DOCX",
    "WorkbookEditor": "XLSX",
    "PresentationEditor": "PPTX",
    "PdfEditor": "PDF",
    "MarkdownEditor": "Markdown",
}

# Async and file-path forms are the same capability wearing a different signature;
# listing them would treble the table without adding a fact. `ToFileAsync` must come
# first: Python's alternation is leftmost-FIRST, not leftmost-longest, so the other
# order strips only `Async` and leaves `CreateToFile` sitting beside `Create`.
NOISE = re.compile(r"(ToFileAsync|Async)$")

# Any four-space `public` line opens a new type. Matching only `public static class`
# is not enough, and the difference is not theoretical: the approved file is
# alphabetical, so `DocxHeader` follows `DocxEditor` and `PdfMetadata` follows
# `PdfEditor`. Without a boundary that closes the previous type, those neighbours'
# static factories were read as editor operations - the first run of this script
# credited `DocxEditor` with `Of` and `Text`, `PdfEditor` with `Titled`, and
# `WorkbookEditor` with `Named` and `From`. All five belong to the class below.
TYPE_LINE = re.compile(r"^    public ")


def parse() -> tuple[list[tuple[str, str]], dict[str, list[str]]]:
    """Read the approved surface into (conversions, editor -> method names)."""
    text = chr(10).join(p.read_text(encoding="utf-8") for p in APPROVED)

    conversions: list[tuple[str, str]] = []
    editors: dict[str, list[str]] = {}
    current: str | None = None

    for line in text.split("\n"):
        if TYPE_LINE.match(line):
            current = None
            cls = re.match(r"^    public static class (\w+)\s*$", line)
            if not cls:
                continue
            name = cls.group(1)
            pair = re.fullmatch(r"(\w+?)To(\w+?)Converter", name)
            if pair:
                conversions.append((
                    DISPLAY.get(pair.group(1), pair.group(1)),
                    DISPLAY.get(pair.group(2), pair.group(2)),
                ))
            elif name in EDITOR_FORMAT:
                current = name
                editors[name] = []
            continue

        if current is not None:
            m = re.match(r"^        public static [\w<>\[\],\.\s\?]+?\s(\w+)\(", line)
            if m:
                name = NOISE.sub("", m.group(1))
                if name and name not in editors[current]:
                    editors[current].append(name)

    return conversions, editors


def render(conversions: list[tuple[str, str]], editors: dict[str, list[str]]) -> str:
    formats = sorted({f for pair in conversions for f in pair})
    edges = set(conversions)

    out = ["| From ↓ / To → | " + " | ".join(formats) + " |",
           "|---|" + "---|" * len(formats)]
    for src in formats:
        cells = []
        for dst in formats:
            cells.append("—" if src == dst else ("**✅**" if (src, dst) in edges else "·"))
        out.append(f"| **{src}** | " + " | ".join(cells) + " |")

    out += ["", "A **✅** is a converter that ships; **·** is a pair with no converter, not a "
                "promise about one. Read a row as \"from this format, into these\".", ""]

    out += ["## Editing an existing document", "",
            "| Format | Operations |", "|---|---|"]
    for editor, methods in sorted(editors.items(), key=lambda kv: EDITOR_FORMAT[kv[0]]):
        listed = ", ".join(f"`{m}`" for m in methods)
        out.append(f"| **{EDITOR_FORMAT[editor]}** (`{editor}`) | {listed} |")

    out += ["", "Method names only. What each one does, and the traps in it, are in the guides — "
                "this table exists to be complete and current, which prose has repeatedly failed "
                "to be.", ""]
    return "\n".join(out)


def main() -> int:
    # The diff this prints contains the table's own ✅ and ·, and a Windows console
    # defaults to cp1252, where encoding either one raises. Found by running --check
    # against a deliberately stale page: the gate correctly exited 1, but did it by
    # crashing with a UnicodeEncodeError traceback instead of printing the diff and
    # the one-line instruction. CI is UTF-8 and would never have shown it.
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    check = "--check" in sys.argv[1:]
    conversions, editors = parse()

    if not conversions or not editors:
        sys.exit("error: parsed no conversions or no editors from the approved API - "
                 "the approved file's shape changed and this script is now blind. Fix the parser "
                 "rather than letting it report an empty table.")

    body = render(conversions, editors)

    if not PAGE.exists():
        sys.exit(f"error: {PAGE.relative_to(REPO)} does not exist; create it with the "
                 f"BEGIN/END markers first.")

    text = PAGE.read_text(encoding="utf-8")
    if BEGIN not in text or END not in text:
        sys.exit(f"error: {PAGE.relative_to(REPO)} is missing the BEGIN/END GENERATED markers.")

    head, rest = text.split(BEGIN, 1)
    _, tail = rest.split(END, 1)
    updated = f"{head}{BEGIN}\n\n{body}\n{END}{tail}"

    print(f"{len(conversions)} conversions, {len(editors)} editors, "
          f"{sum(len(v) for v in editors.values())} operations")

    if updated == text:
        print("capability matrix is up to date")
        return 0

    if check:
        diff = difflib.unified_diff(
            text.splitlines(keepends=True), updated.splitlines(keepends=True),
            fromfile="committed", tofile="derived from the approved API")
        sys.stdout.writelines(diff)
        print("\n::error::The capability matrix no longer matches the approved public API. "
              "Run `python scripts/gen-capability-matrix.py` and commit the result.")
        return 1

    PAGE.write_text(updated, encoding="utf-8", newline="\n")
    print(f"rewrote {PAGE.relative_to(REPO)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
