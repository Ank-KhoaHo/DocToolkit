#!/usr/bin/env python3
"""Print one version's section from CHANGELOG.md, for use as GitHub Release notes.

WHY THIS EXISTS. Curating a Release PR fixes CHANGELOG.md and NOT the published
GitHub Release notes. Measured 2026-08-15 on v0.27.2: the curated entry landed in
the repository, while the Release page still showed release-please's generated
text - including a duplicate line the curation had removed. release-please builds
the Release body from the COMMITS at release time, independently of the changelog
file it also writes, so the two surfaces diverge the moment anybody edits one.

The repo's whole curation ritual assumed one surface. SOURCE-QUALITY.md Part 3 and
the enhancement backlog both say "read the changelog entry" before merging a
Release PR; neither mentions the Release body, and nothing re-synced them. The
v0.27.2 divergence was fixed by hand afterwards, which works only because somebody
happened to look.

So the Release body is now DERIVED from the changelog rather than remembered
alongside it - the same move as gen-third-party-notices.py reading the lockfile
and gen-capability-matrix.py reading the approved API. One curated source, both
surfaces.

WHAT IT DELIBERATELY DOES NOT DO. It does not rewrite, reformat or summarise. The
section is emitted verbatim, because the whole point is that a human curated it.

USAGE

    python scripts/extract-changelog-section.py 0.27.2
    python scripts/extract-changelog-section.py v0.27.2   # leading v is tolerated

Prints the section to stdout. Exits non-zero, printing nothing, if the version is
absent or its section is empty - a Release body must never be silently blanked by
a parser that stopped matching.
"""
from __future__ import annotations

import pathlib
import re
import sys

REPO = pathlib.Path(__file__).resolve().parent.parent
CHANGELOG = REPO / "CHANGELOG.md"

# release-please writes headings as `## [0.27.2](compare-link) (date)`. Matching the
# bracketed version rather than the whole line keeps this working if the link or the
# date format changes, both of which are release-please's to decide, not ours.
HEADING = re.compile(r"^## \[(?P<version>[^\]]+)\]", re.M)


def main() -> int:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    if len(sys.argv) != 2:
        sys.exit("usage: extract-changelog-section.py <version>")

    wanted = sys.argv[1].lstrip("vV").strip()
    if not wanted:
        sys.exit("error: empty version argument")

    if not CHANGELOG.exists():
        sys.exit(f"error: {CHANGELOG.name} not found - is the repository checked out?")

    text = CHANGELOG.read_text(encoding="utf-8")

    headings = list(HEADING.finditer(text))
    if not headings:
        sys.exit("error: no '## [version]' headings found in CHANGELOG.md - the changelog's "
                 "shape changed and this script is now blind. Fix the parser rather than "
                 "letting it emit nothing.")

    for i, m in enumerate(headings):
        if m.group("version") != wanted:
            continue
        end = headings[i + 1].start() if i + 1 < len(headings) else len(text)
        section = text[m.start():end].rstrip() + "\n"

        # A heading with nothing under it is worse than a failure: it would replace
        # real generated notes with a bare title. Require actual content.
        body = section.split("\n", 1)[1].strip() if "\n" in section else ""
        if not body:
            sys.exit(f"error: the section for {wanted} is empty. Refusing to emit a Release "
                     f"body that is only a heading.")

        sys.stdout.write(section)
        return 0

    available = ", ".join(h.group("version") for h in headings[:5])
    sys.exit(f"error: version {wanted} not found in CHANGELOG.md. Most recent: {available}")


if __name__ == "__main__":
    raise SystemExit(main())
