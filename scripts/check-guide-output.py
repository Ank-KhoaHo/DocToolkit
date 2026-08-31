#!/usr/bin/env python3
"""Asserts every hand-copied ```text block in docfx/guides/*.md matches what its sample
actually prints today.

B29: the guides pull CODE from compiled samples so a snippet cannot drift from the API -
check-doc-snippets.py exists for that - but their OUTPUT was nine hand-copied blocks across
five guides with nothing re-running the sample and comparing. It went stale within one commit:
a sample edit changed a printed byte count from 7,370 to 7,907 and the guide kept saying 7,576
until this was written.

Each block is a CURATED EXCERPT of a sample's full stdout, not a verbatim capture - some lines
the sample prints (a byte count judged too volatile to show, mostly) are simply left out of the
guide. So this does not diff the whole run; it looks up each expected line by a unique LABEL
prefix (the text before the aligned ":") and reconstructs what the block should read from the
live run, in the order BLOCKS below lists them.

Several labels carry a byte count from a compressed artefact (a .docx/.pdf/.xlsx ZIP) rather
than plain text or a simple length, and are handled differently depending on what was actually
measured, not assumed:

  - PDF byte counts (PdfUtilities' "Merged"/"Page 2 alone") are masked, never compared exactly.
    CLAUDE.md already establishes why: PDF size varies up to ~100x with which fonts are
    installed on the machine that rendered it, which is a property of the runner image, not of
    this code. The same rule applies here that already applies to this repo's test suite.
  - DOCX byte counts (DocxImages' two lines) are ALSO masked, found the hard way while writing
    this: two otherwise-identical DOCX outputs from the same sample, in the same process, varied
    by a byte between consecutive calls. DocumentFormat.OpenXml/OfficeIMO assigns each document a
    fresh random relationship ID, and that shifts a compressed ZIP's size independent of content
    - so the two documents cannot be asserted equal OR asserted at an exact value, on any
    platform. The guide's own prose was rewritten alongside this script once that was found; it
    no longer claims "byte-for-byte", which measured to be false.
  - The XLSX byte count (Spreadsheets' "Formatted" line) IS compared exactly. Measured
    separately: five consecutive identical `WorkbookEditor.Format` calls produced the same byte
    count all five times, on one platform - ClosedXML does not carry the same random-ID
    behaviour. This check also always runs on one pinned CI leg (ubuntu-24.04, the same runner
    the `formatting` job already uses for every other generated-content guard), so run-to-run
    drift on THAT leg is not a risk this check has to hedge against.

    **The claim that follows from that ("no cross-platform DEFLATE question to hedge against
    even if there were") was wrong, and cost this PR a red CI run to find out.** The number this
    file first shipped with (7,907 bytes) was captured on Windows, not on ubuntu-24.04, and a
    live run on the pinned CI leg produced 7,880 - measured directly, same .NET SDK version on
    both machines (10.0.302), so the 27-byte difference is `System.IO.Compression`'s DEFLATE
    output differing by OS, not by SDK patch or by anything in this code. The check itself was
    never wrong: it is comparing against exactly the platform it always runs on, which is the
    whole reason exact comparison is safe here. What was wrong is generating or hand-verifying
    the recorded guide text anywhere other than that same platform - a local Windows (or macOS)
    run is not a substitute for what `--check` will actually compare against on the next push.

Run with --check to fail on drift (CI); without it, rewrites every affected guide block in place
so `git diff` shows exactly what changed (the maintainer's own workflow for a real drift, one
level up from "go and hand-edit the number").
"""
import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

MASK = object()  # sentinel: this label's value is never compared, only that the label exists


def line(label, value=None):
    """One expected guide line. value=None reads the live run; value=MASK never compares."""
    return (label, value)


# Each block: which guide, which fenced ```text occurrence (1-based, in file order), which
# sample supplies its lines, and the ordered (label, override) pairs the block is built from.
# A None override means "read this label's line verbatim from the live run"; MASK means "the
# label must appear, but its value is never compared" (see PDF note above).
BLOCKS = [
    {
        "guide": "docfx/guides/getting-started.md",
        "occurrence": 1,
        "sample": "HtmlConversion",
        "lines": [line("Rejected"), line("Inner cause")],
    },
    {
        "guide": "docfx/guides/conversions/html-to-word-and-pdf.md",
        "occurrence": 1,
        "sample": "HtmlConversion",
        # All three masked: each is a freshly-written DOCX or PDF, and OfficeIMO/OpenXml assigns
        # a fresh random relationship id per document - the same non-determinism DocxImages'
        # pair below is masked for, just three ways instead of two.
        "lines": [
            line("HTML -> DOCX", MASK),
            line("HTML -> PDF", MASK),
            line("DOCX -> PDF", MASK),
        ],
    },
    {
        "guide": "docfx/guides/conversions/html-to-word-and-pdf.md",
        "occurrence": 2,
        "sample": "HtmlConversion",
        # "Landscape :" also appears between these two in the sample's real output - the guide
        # curates it out, so it is absent from this list on purpose, not a missed line.
        "lines": [line("Default"), line("This one"), line("Letter intact")],
    },
    {
        "guide": "docfx/guides/conversions/html-to-word-and-pdf.md",
        "occurrence": 3,
        "sample": "DocxImages",
        # Masked, not compared or asserted equal: measured 2026-08-28, two otherwise-identical
        # DOCX outputs from THIS SAME sample varied by a byte between consecutive runs in the
        # same process - the writer assigns a fresh random relationship ID per document, which
        # shifts a compressed ZIP's size independent of content. See the guide prose for what
        # this pair actually demonstrates instead (no network I/O, either path).
        "lines": [line("Remote off", MASK), line("Not on list", MASK)],
    },
    {
        "guide": "docfx/guides/conversions/html-to-word-and-pdf.md",
        "occurrence": 4,
        "sample": "PdfUtilities",
        "lines": [
            line("Three documents"),
            line("Merged", MASK),
            line("Page 2 alone", MASK),
        ],
    },
    {
        "guide": "docfx/guides/conversions/html-to-word-and-pdf.md",
        "occurrence": 5,
        "sample": "PdfUtilities",
        "lines": [line("After retitling")],
    },
    {
        "guide": "docfx/guides/conversions/markdown.md",
        "occurrence": 1,
        "sample": "MarkdownConversion",
        # A blank line separates the warning table from the summary in the guide - reproduced
        # literally below rather than modelled as a third "label", since it carries none.
        "lines": [line("  Approximation"), line(""), line("HasLoss"), line("What it says")],
    },
    {
        "guide": "docfx/guides/editing/spreadsheets-and-presentations.md",
        "occurrence": 1,
        "sample": "Spreadsheets",
        "lines": [line("Formatted")],
    },
    {
        "guide": "docfx/guides/editing/spreadsheets-and-presentations.md",
        "occurrence": 2,
        "sample": "Spreadsheets",
        "lines": [line("As CSV"), line("As HTML")],
    },
    {
        "guide": "docfx/guides/editing/spreadsheets-and-presentations.md",
        "occurrence": 3,
        "sample": "Spreadsheets",
        "lines": [line("Pivot cell D1 right after creation")],
    },
    {
        "guide": "docfx/guides/editing/spreadsheets-and-presentations.md",
        "occurrence": 4,
        "sample": "Presentations",
        # DocumentFormat.OpenXml assigns each saved part a fresh random relationship id, which
        # shifts a compressed ZIP's size independent of content - the same non-determinism
        # DocxImages' pair is masked for above, measured directly here too (three consecutive
        # runs: 12,034 / 12,041 / 12,045 bytes).
        "lines": [line("With chart", MASK)],
    },
    {
        "guide": "docfx/guides/editing/spreadsheets-and-presentations.md",
        "occurrence": 5,
        "sample": "Presentations",
        "lines": [line("SmartArt"), line("Diagram text"), line("In ExtractText too")],
    },
    {
        "guide": "docfx/guides/editing/word-documents.md",
        "occurrence": 1,
        "sample": "DocxTemplating",
        "lines": [line("As HTML"), line("As Markdown")],
    },
]


def run_sample(name):
    """Runs samples/<name> in Release and returns its stdout, split into lines."""
    proc = subprocess.run(
        ["dotnet", "run", "--project", str(ROOT / "samples" / name), "-c", "Release",
         "--no-launch-profile"],
        cwd=ROOT, capture_output=True, text=True, timeout=180,
    )
    if proc.returncode != 0:
        print(f"samples/{name} exited {proc.returncode}", file=sys.stderr)
        print(proc.stdout, file=sys.stderr)
        print(proc.stderr, file=sys.stderr)
        sys.exit(1)
    return proc.stdout.splitlines()


def find_by_label(sample_lines, label):
    """The one line in a sample's real output that STARTS WITH label. Most guide lines are
    "Label   : value", but MarkdownConversion's warning table row has no colon at all - so this
    matches by prefix rather than by parsing a colon, which handles both shapes the same way.
    Each label here is chosen to be unique across its sample's own output."""
    if label == "":
        return ""
    matches = [candidate for candidate in sample_lines if candidate.startswith(label)]
    if len(matches) == 1:
        return matches[0]
    if not matches:
        print(f"  no line starting with {label!r} in this sample's output", file=sys.stderr)
    else:
        print(f"  {len(matches)} lines start with {label!r}, need exactly one: {matches}",
              file=sys.stderr)
    return None


def extract_block(md_path, occurrence):
    """Returns (start_line_index, end_line_index, current_lines) for the Nth ```text block,
    both indices into the file's list of lines, end EXCLUSIVE of the closing fence."""
    text = md_path.read_text(encoding="utf-8")
    lines = text.split("\n")
    seen = 0
    i = 0
    while i < len(lines):
        if lines[i].strip() == "```text":
            seen += 1
            if seen == occurrence:
                start = i + 1
                end = start
                while lines[end].strip() != "```":
                    end += 1
                return start, end, lines[start:end]
            i += 1
        i += 1
    raise ValueError(f"{md_path}: fewer than {occurrence} ```text blocks")


def build_expected(sample_lines, block):
    expected = []
    for label, override in block["lines"]:
        if override is MASK:
            actual = find_by_label(sample_lines, label)
            if actual is None:
                sys.exit(1)
            # Keep everything up to and including the number's own unit word, replace the
            # digits only - so a genuine wording change ("bytes" -> "KB") still fails loudly.
            masked = re.sub(r"[\d,]+(\s*bytes)", "<varies>\\1", actual, count=1)
            expected.append(masked)
        elif override is None:
            actual = find_by_label(sample_lines, label)
            if actual is None:
                sys.exit(1)
            expected.append(actual)
        else:
            expected.append(override)
    return expected


def main():
    check_only = "--check" in sys.argv
    cache = {}
    failures = []

    for block in BLOCKS:
        sample = block["sample"]
        if sample not in cache:
            print(f"Running samples/{sample} ...")
            cache[sample] = run_sample(sample)
        sample_lines = cache[sample]

        expected = build_expected(sample_lines, block)

        md_path = ROOT / block["guide"]
        start, end, current = extract_block(md_path, block["occurrence"])

        if current != expected:
            failures.append(
                f"{block['guide']} occurrence {block['occurrence']} (samples/{sample}):\n"
                f"  guide says : {current}\n"
                f"  live run   : {expected}")
            if not check_only:
                lines = md_path.read_text(encoding="utf-8").split("\n")
                lines[start:end] = expected
                md_path.write_text("\n".join(lines), encoding="utf-8")
                print(f"  rewrote {block['guide']} occurrence {block['occurrence']}")

    if failures:
        print(f"\n{len(failures)} guide output block(s) do not match a live sample run:\n")
        for f in failures:
            print(f)
            print()
        if check_only:
            sys.exit(1)
        else:
            print("Rewritten in place - review the diff before committing.")
            sys.exit(0)

    print(f"All {len(BLOCKS)} guide output blocks match their samples.")


if __name__ == "__main__":
    main()
