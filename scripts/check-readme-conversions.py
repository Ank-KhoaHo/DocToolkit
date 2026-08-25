#!/usr/bin/env python3
"""Fail when a README enumerates a format's conversion targets and leaves one out.

WHY THIS EXISTS. Issue #381: the README headline read

    plus Markdown -> DOCX/PDF, DOCX -> HTML/Markdown, XLSX -> CSV/HTML ...

while `DocxToPdfConverter` and `XlsxToPdfConverter` both ship. The headline advertised a
narrower package than the one released, on the surface a search result shows and a
first-time reader reads. "DOCX to PDF" is one of the highest-volume searches in this
space and the phrase was absent from it.

The reporter nearly skipped the package over it, having found the capability by reading
to line 225.

WHY THE EXISTING CHECKS MISSED IT, which is the interesting part. Three checks already
guard README and description accuracy and none of them could see this:

  * `gen-capability-matrix.py` derives the grid from the approved API and fails on drift -
    so the GRID stayed correct while the prose beside it went stale.
  * `check-readme-coverage.py` asserts every shipped public TYPE is named. `DocxToPdfConverter`
    is named further down, so it passed.
  * `check-package-description.py` asserts every shipped FORMAT appears in the csproj
    description. "PDF" appears there many times over, so it passed.

Every one of them was satisfied by a document that told a reader the wrong thing. What
none of them checked was a claim about which conversions EXIST.

WHAT THIS CHECKS, and what it deliberately does not. It finds every enumeration of the
form `SOURCE -> A/B/C` and asserts the set {A, B, C} is exactly the shipped targets for
that source, derived from the `<From>To<To>Converter` names in the approved API - the
same source `gen-capability-matrix.py` reads.

It does NOT try to prove a README is complete, or well written, or that it mentions every
conversion somewhere. No check can, and a check that demanded a headline list all twelve
conversions would produce an unreadable sentence and be switched off within a month -
which is what `check-readme-coverage.py` says about requiring the three READMEs to match.

It proves one falsifiable thing: **where a README enumerates a source's targets, the
enumeration is complete.** That is the property that broke.

    python scripts/check-readme-conversions.py
    python scripts/check-readme-conversions.py --self-test
"""

import glob
import io
import pathlib
import re
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = pathlib.Path(__file__).resolve().parent.parent

READMES = [
    "README.md",
    "src/DocToolkit/README.md",
    "src/DocToolkit.Extensions.DependencyInjection/README.md",
]

# Same map as gen-capability-matrix.py and check-package-description.py: what a reader calls
# the format, not the fragment in a type name.
DISPLAY = {
    "Docx": "DOCX", "Pdf": "PDF", "Html": "HTML", "Xlsx": "XLSX",
    "Pptx": "PPTX", "Csv": "CSV", "Markdown": "Markdown", "Doc": "DOC",
}

CONVERTER = re.compile(r"public static class (\w+)To(\w+)Converter")

# `DOCX → HTML/Markdown`, `XLSX -> CSV/HTML/PDF`. Both arrow forms, since the READMEs use
# the unicode one and a contributor may type the ASCII one.
ENUMERATION = re.compile(
    r"\*{0,2}(" + "|".join(sorted(DISPLAY.values(), key=len, reverse=True)) + r")\*{0,2}"
    r"\s*(?:→|->)\s*"
    r"\*{0,2}([A-Za-z]+(?:/[A-Za-z]+)+)\*{0,2}")


def shipped_targets(api_text=None):
    """{source: {targets}} derived from the approved API's converter names."""
    if api_text is None:
        files = sorted(glob.glob(str(REPO / "tests" / "DocToolkit.Tests" / "PublicApi"
                                     / "DocToolkit*.approved.txt")))
        if not files:
            sys.exit("::error::No approved API files found; this check cannot derive what ships "
                     "and must not pass while unable to.")
        api_text = "\n".join(io.open(f, encoding="utf-8").read() for f in files)

    targets = {}
    unknown = set()
    for source, target in CONVERTER.findall(api_text):
        for fragment in (source, target):
            if fragment not in DISPLAY:
                unknown.add(fragment)
        if source in DISPLAY and target in DISPLAY:
            targets.setdefault(DISPLAY[source], set()).add(DISPLAY[target])

    if unknown:
        sys.exit(f"::error::converter name fragment(s) {', '.join(sorted(unknown))} are not in "
                 "DISPLAY. A new format must be added there rather than silently skipped, or this "
                 "check would stop covering it.")
    return targets


def check(readme_text, targets, label):
    """Failures for one README. Empty means every enumeration in it is complete."""
    failures = []
    seen = 0

    for source, listed in ENUMERATION.findall(readme_text):
        if source not in targets:
            continue                      # not a source this library converts from
        seen += 1
        claimed = {t for t in listed.split("/")}
        # Only flag OMISSIONS. A README naming something not in the API is check-readme-
        # coverage.py's business, and flagging it here would double-report.
        missing = sorted(targets[source] - claimed)
        if missing:
            failures.append(
                f"{label}: enumerates `{source} → {listed}` but also ships "
                f"{source} → {', '.join(missing)}. A reader takes an enumeration as complete.")

    return failures, seen


def self_test():
    api = """
        public static class DocxToPdfConverter
        public static class DocxToHtmlConverter
        public static class DocxToMarkdownConverter
        public static class HtmlToPdfConverter
    """
    targets = shipped_targets(api)
    assert targets["DOCX"] == {"PDF", "HTML", "Markdown"}, targets

    cases = [
        ("complete enumeration passes", "see DOCX → HTML/Markdown/PDF for details", 0),
        ("THE #381 CASE: an omission fails", "plus **DOCX → HTML/Markdown**, and more", 1),
        ("ascii arrow is caught too", "DOCX -> HTML/Markdown", 1),
        ("bold markers around the parts", "**DOCX** → **HTML/Markdown**", 1),
        ("prose that enumerates nothing is ignored", "turn DOCX into HTML or Markdown", 0),
        ("a single target is not an enumeration", "DOCX → PDF", 0),
        ("an unrelated source is ignored", "CSV → JSON/YAML", 0),
    ]

    bad = 0
    for name, text, expected in cases:
        failures, _ = check(text, targets, "t")
        ok = len(failures) == expected
        print(f"  {'ok  ' if ok else 'FAIL'} {name} -> {len(failures)}, expected {expected}")
        bad += 0 if ok else 1

    real = shipped_targets()
    if len(real) < 2:
        print(f"  FAIL only {len(real)} source(s) derived from the real approved API; a "
              "derivation finding nothing would pass every README")
        bad += 1
    else:
        print(f"  ok   {len(real)} source(s) derived from the real approved API: "
              + ", ".join(f"{s} → {'/'.join(sorted(t))}" for s, t in sorted(real.items())))

    print()
    if bad:
        print(f"::error::self-test failed {bad} case(s)")
        return 1
    print("self-test passed, including the case from issue #381")
    return 0


def main(argv):
    if "--self-test" in argv:
        return self_test()

    targets = shipped_targets()
    failures, total = [], 0

    for readme in READMES:
        path = REPO / readme
        if not path.exists():
            sys.exit(f"::error::{readme} not found; this check cannot run.")
        found, seen = check(io.open(path, encoding="utf-8").read(), targets, readme)
        failures += found
        total += seen
        print(f"  {readme}: {seen} conversion enumeration(s), {len(found)} incomplete")

    if total == 0:
        sys.exit("::error::No conversion enumerations found in any README. Either they were all "
                 "rephrased or this check stopped matching them; both are worth a red build, "
                 "because a check that silently examines nothing is worse than no check.")

    print()
    if not failures:
        print(f"Every conversion enumeration is complete ({total} checked).")
        return 0

    for f in failures:
        print(f"::error::{f}")
    print()
    print("Issue #381: the headline advertised a narrower package than the one that shipped, on "
          "the surface a search result shows. The generated capability grid stayed correct the "
          "whole time - prose beside a derived table is not covered by it.")
    return 1


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
