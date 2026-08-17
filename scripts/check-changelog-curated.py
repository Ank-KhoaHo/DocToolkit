#!/usr/bin/env python3
"""Report whether the pending release's changelog entry has been curated, or is still
release-please's generated text.

WHY THIS EXISTS. Curating a Release PR entry and then merging anything else to main
loses the curation, silently. Every push to main makes release-please rewrite the
entry from commit titles, so a curation survives only if nothing merges afterwards -
and there is no signal when it does not. The Release PR simply has different content
than it had an hour ago, and it looks entirely fine, because generated text always
looks fine.

Measured three times: twice on 0.28.0, where PDF password protection shipped with no
changelog entry AT ALL, and again on 0.30.0 on 2026-08-17 - that third time with the
warning already written in CLAUDE.md and a curation checklist already posted to the
pull request. Being told to curate is evidently not the same as being told the
curation is currently absent.

The cost is permanent. A published changelog entry can be unlisted but never edited.

HOW IT DECIDES, and why it is not simply "does an item look generated". release-please
ends every item it writes with a link to the commit:

    * **core:** claim legacy .ppt on the PDF path ([#290](...)) ([21cb905](...))

The first version of this check tested only that, and got 0.29.0 wrong: its single
item is that exact generated line with a full stop appended, so a per-item test called
the whole entry uncurated - when in fact it carries a hand-written paragraph above the
list saying the release is extensions-only. Curation is not confined to rewriting
items; it adds prose, Migrating notes and section headings.

So the question asked here is broader and matches what actually happened: IS THERE ANY
CONTENT IN THIS ENTRY THAT RELEASE-PLEASE DID NOT WRITE? Every non-blank line is
classified as either a generated list item, a section heading, or human content. An
entry with no human content anywhere has not been touched by anybody.

WHAT IT DELIBERATELY DOES NOT CLAIM. It detects authorship, not quality. An entry
curated badly passes. It answers "has a human touched this?" and nothing more - which
is exactly the question that was answered wrong three times, and a check that tried to
judge prose would be switched off within a month.

REPORTS, DOES NOT BLOCK, and that is deliberate. release-please only proposes a
release for a version-bumping commit, so an entry is rarely worth nothing - but
"rarely" is not "never", and a blocking check would be disabled the first time
somebody legitimately wanted to ship generated text. release.yml's existing
non-empty-entry guard already covers the unrecoverable case at publish time; this one
only has to make the state visible while it can still be changed.

USAGE

    python scripts/check-changelog-curated.py                  # the newest entry
    python scripts/check-changelog-curated.py 0.31.0           # a named one (leading v ok)
    python scripts/check-changelog-curated.py --changelog F    # read F instead of CHANGELOG.md
    python scripts/check-changelog-curated.py --self-test      # positive and negative controls

Exit 0 when the entry carries human content, 1 when it does not, 2 on a usage or parse
error. The caller decides what a 1 means - see .github/workflows/release-please.yml,
which turns it into a comment rather than a failure.
"""
from __future__ import annotations

import pathlib
import re
import sys

REPO = pathlib.Path(__file__).resolve().parent.parent
CHANGELOG = REPO / "CHANGELOG.md"
TESTDATA = REPO / "scripts" / "testdata"

# Same heading shape extract-changelog-section.py matches, and for the same reason: the
# link and the date are release-please's to change, the bracketed version is not.
HEADING = re.compile(r"^## \[(?P<version>[^\]]+)\]", re.M)

# release-please's signature: a ONE-LINE item ending in a link whose text is a commit
# sha, optionally followed by punctuation somebody added without otherwise touching it.
GENERATED_ITEM = re.compile(r"^\* .*\(\[[0-9a-f]{7,40}\]\([^)]*\)\)[.\s]*$")

SECTION_HEADING = re.compile(r"^#{3,} ")


def entry(text: str, wanted: str | None) -> tuple[str, str]:
    headings = list(HEADING.finditer(text))
    if not headings:
        sys.exit("error: no '## [version]' headings found - the changelog's shape changed and "
                 "this script is now blind. Fix the parser rather than letting it report "
                 "'curated' about nothing.")

    if wanted is None:
        chosen = 0
    else:
        matches = [i for i, h in enumerate(headings) if h.group("version") == wanted]
        if not matches:
            available = ", ".join(h.group("version") for h in headings[:5])
            sys.exit(f"error: version {wanted} not found. Most recent: {available}")
        chosen = matches[0]

    start = headings[chosen].start()
    end = headings[chosen + 1].start() if chosen + 1 < len(headings) else len(text)
    return headings[chosen].group("version"), text[start:end]


def classify(section: str) -> tuple[list[str], list[str]]:
    """Split the entry's content lines into (generated items, human content)."""
    generated: list[str] = []
    human: list[str] = []

    for line in section.splitlines()[1:]:          # [1:] drops the '## [version]' heading
        stripped = line.strip()
        if not stripped:
            continue
        if SECTION_HEADING.match(stripped):        # '### Added' is release-please's too
            continue
        if GENERATED_ITEM.match(stripped):
            generated.append(stripped)
        else:
            human.append(stripped)

    return generated, human


def report(version: str, section: str) -> int:
    generated, human = classify(section)

    if not generated and not human:
        # Distinct from "not curated": there is nothing here at all, which release.yml's
        # own guard refuses to publish. Say which problem it is.
        print(f"{version}: the entry has NO content at all.")
        print("  release.yml will refuse to publish this. Let more accumulate, or write an entry.")
        return 1

    print(f"{version}: {len(generated)} generated item(s), {len(human)} line(s) of human content.")

    if not human:
        print()
        print("  NOTHING HAS BEEN CURATED. Every line is still release-please's text, derived")
        print("  from commit titles - so a behaviour change committed as `feat:` is sitting")
        print("  under Added described as a new capability, and anything that arrived inside a")
        print("  squashed stacked PR may have no entry at all.")
        print()
        print("  If you curated this already, a later merge to main has REGENERATED it.")
        print("  Curate immediately before merging, with nothing else in flight.")
        return 1

    if generated:
        print(f"  {len(generated)} item(s) are still verbatim generated text:")
        for item in generated:
            print(f"    - {item[:100]}")
    return 0


# ---- controls ---------------------------------------------------------------------------
#
# A check that reported "curated" because its parser matched nothing would be worse than
# no check, and would look identical from the outside. So it is run against a known
# generated entry and a known curated one, and BOTH verdicts are asserted - along with the
# parser having actually found something in each, which is what makes the pass non-vacuous.

def self_test() -> int:
    cases = [
        ("changelog-generated.md", 1, "an untouched release-please entry"),
        ("changelog-curated.md", 0, "the same entry after curation"),
    ]

    failures = 0
    for name, expected, description in cases:
        path = TESTDATA / name
        if not path.exists():
            print(f"FAIL  {name} is missing - the controls cannot run")
            failures += 1
            continue

        version, section = entry(path.read_text(encoding="utf-8"), None)
        generated, human = classify(section)
        actual = report(version, section)

        ok = actual == expected
        # Non-vacuity: the fixture must actually have parsed into something. A parser that
        # matched nothing would return 1 for the generated case and look correct.
        parsed = bool(generated or human)
        if not parsed:
            print(f"FAIL  {name} parsed to NOTHING - the verdict is meaningless")
            ok = False

        print(f"{'PASS' if ok else 'FAIL'}  {description}: expected exit {expected}, got {actual}")
        print()
        failures += 0 if ok else 1

    print("self-test:", "all controls pass" if not failures else f"{failures} FAILED")
    return 0 if not failures else 1


def main() -> int:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    args = sys.argv[1:]
    if "--self-test" in args:
        return self_test()

    source = CHANGELOG
    if "--changelog" in args:
        i = args.index("--changelog")
        if i + 1 >= len(args):
            sys.exit("usage: --changelog <path>")
        source = pathlib.Path(args[i + 1])
        del args[i:i + 2]

    if len(args) > 1:
        sys.exit("usage: check-changelog-curated.py [version] [--changelog PATH] [--self-test]")
    wanted = args[0].lstrip("vV").strip() if args else None

    if not source.exists():
        sys.exit(f"error: {source} not found - is the repository checked out?")

    version, section = entry(source.read_text(encoding="utf-8"), wanted)
    return report(version, section)


if __name__ == "__main__":
    raise SystemExit(main())
