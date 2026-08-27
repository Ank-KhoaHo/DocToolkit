#!/usr/bin/env python3
"""Fail when a tracked file names one of this machine's PRIVATE sibling repositories.

WHY THIS EXISTS. DocToolkit is the only public repository of five that sit side by side in one
folder. Anything committed here is world-readable, permanently, including in history.

On 2026-08-27 a doc comment on a PUBLIC class read:

    Pivoting through DOCX keeps the whole chain pure managed.
    See learning-docs/dotnet-doc-libs/report.html.

`learning-docs/` is in AutoLnD, a private repository. Because the sentence sat in a <summary>, it
was compiled into DocToolkit.xml, PACKED INSIDE THE NUGET PACKAGE, and published: verified by
downloading Ank.DocToolkit 0.41.0 from nuget.org and finding the string in lib/net10.0/DocToolkit.xml.
It also renders on the public API site, telling consumers to consult a path they can never reach.

No secret was disclosed. What leaked is the existence and layout of private work, and a broken
reference shipped to every consumer - and a published package version can be unlisted but never
edited, so it is permanent in every version that carried it.

**This had happened before.** CHANGELOG.md still carries "tighten the root README and drop the
AutoLnD reference", so the class of mistake was found and cleaned once by hand, and this instance
survived that cleanup. A hand-cleaned leak with nothing watching for the next one is the failure
mode this repository records everywhere else; hence a check rather than a third careful read.

WHAT IT SCANS. Tracked files only, via `git ls-files`, because the point is what a clone gets.
Untracked material - CLAUDE.md, docs/, BACKLOG.html, .superpowers/ - is deliberately ignored: it is
gitignored precisely so it can hold this vocabulary.

    python scripts/check-no-private-references.py
    python scripts/check-no-private-references.py --self-test
"""

import re
import subprocess
import sys

# The four private repositories beside this one, plus the container path they share. Names rather
# than paths, because a name is what leaks in prose.
PRIVATE = [
    "AutoLnD",
    "Prj-Indie-Alpha",
    "manga-translation-tool",
    "Transaltion-Game-AliceSoft",     # the typo is in the real repository name
    "learning-docs",                  # AutoLnD's research tree, the one that actually leaked
    "LnDPrj",                         # the container folder
]

# Two exemptions, each with its reason. Every entry is a hole in the check, so argue for one
# rather than adding it to make a build green.
#
# CHANGELOG.md is written by release-please and its entries are published; one historical line
# names a private repo as the SUBJECT of a removal. Editing it would not unpublish the release
# notes or the commit, and release-please would overwrite the edit anyway.
EXEMPT = {
    "CHANGELOG.md",
    # THIS FILE, which must contain every name in order to search for them.
    #
    # Worth knowing how the exemption was found rather than designed: the first local run passed
    # cleanly, because `git ls-files` lists TRACKED files and the script was not committed yet. It
    # started failing the moment it was - on itself, eleven times. A guard whose first green run
    # predates its own tracking has not been tested against the tree it guards.
    "scripts/check-no-private-references.py",
}

PATTERN = re.compile("|".join(re.escape(p) for p in PRIVATE), re.IGNORECASE)


def tracked_files():
    out = subprocess.run(["git", "ls-files"], capture_output=True, text=True, check=True).stdout
    return [line for line in out.splitlines() if line]


def scan(paths):
    hits = []
    for path in paths:
        if path in EXEMPT:
            continue
        try:
            with open(path, encoding="utf-8", errors="ignore") as handle:
                for number, line in enumerate(handle, 1):
                    match = PATTERN.search(line)
                    if match:
                        hits.append((path, number, match.group(0), line.strip()[:90]))
        except OSError:
            continue
    return hits


def self_test():
    """Controls: the pattern must fire on the real leak and not on ordinary prose."""
    must_match = [
        "/// DOCX keeps the whole chain pure managed. See learning-docs/dotnet-doc-libs/report.html.",
        "* tighten the root README and drop the AutoLnD reference",
        "cross-referenced against the Prj-Indie-Alpha backlog",
        "E:\\PJ\\LnDPrj\\DocToolkit",
    ]
    must_not = [
        "/// Converts HTML to PDF by pivoting through DOCX.",
        "the docs/ folder is gitignored",
        "learning to use this library takes a minute",   # 'learning' alone must not fire
        "see the specs folder",
    ]

    bad = [s for s in must_match if not PATTERN.search(s)]
    bad += [s for s in must_not if PATTERN.search(s)]
    if bad:
        print("::error::self-test FAILED on:")
        for s in bad:
            print(f"  {s}")
        return 1

    print(f"self-test passed: {len(must_match)} caught, {len(must_not)} correctly ignored")
    return 0


def main(argv):
    if "--self-test" in argv:
        return self_test()

    files = tracked_files()
    if not files:
        print("::error::git ls-files returned nothing - refusing to report a clean scan of no files.")
        return 1

    hits = scan(files)
    if not hits:
        print(f"{len(files)} tracked file(s) scanned; none names a private sibling repository.")
        return 0

    print()
    for path, number, term, line in hits:
        print(f"::error file={path},line={number}::names the private repository '{term}': {line}")
    print()
    print("DocToolkit is PUBLIC and the four repositories beside it are not. A name in a doc "
          "comment is worse than one in a script: it compiles into DocToolkit.xml, ships inside "
          "the .nupkg, and renders on the API site - which is exactly how one reached nuget.org "
          "0.41.0. Move the reference into the gitignored material, or say the same thing without "
          "naming private work.")
    return 1


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
