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

TWO MORE SURFACES (D48).

- NAMES THIS FILE CANNOT HOLD. The list below was written when four private repositories sat
  beside this one; there are more now, and writing their names here would publish them, which
  is the leak itself. So more names come from `.leak-patterns.local`: gitignored, one regular
  expression per line, case-sensitive (start a line with (?i) to change that), `#` for a
  comment. A clone and a CI runner have no such file and scan with the list below only; the
  run says which.
- COMMIT MESSAGES. A subject is as public as a file, permanently, and a subject that sums up
  what was just kept out is the same leak somewhere nobody reviews. Locally, every commit in
  `origin/main..HEAD` is scanned - the commits a push would publish. Where `origin/main` is
  unknown (a shallow CI checkout), the run says the messages were not scanned.

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

LOCAL_PATTERNS = ".leak-patterns.local"
BASE = "origin/main"


def load_local_patterns(path):
    """The compiled patterns of a gitignored patterns file; [] when there is none.

    A line that is not a valid regular expression stops the run: skipped, it would quietly look
    for nothing."""
    try:
        with open(path, encoding="utf-8") as handle:
            lines = handle.read().splitlines()
    except FileNotFoundError:
        return []
    patterns = []
    for number, line in enumerate(lines, 1):
        text = line.strip()
        if not text or text.startswith("#"):
            continue
        try:
            patterns.append(re.compile(text))
        except re.error as exc:
            raise SystemExit(f"::error file={path},line={number}::not a regular expression ({exc}): {text}")
    return patterns


def first_match(line, extra):
    """The private name a line holds, or None: the built-in list first, then the local file's."""
    for pattern in (PATTERN, *extra):
        match = pattern.search(line)
        if match:
            return match.group(0)
    return None


def tracked_files():
    out = subprocess.run(["git", "ls-files"], capture_output=True, text=True, check=True).stdout
    return [line for line in out.splitlines() if line]


def scan(paths, extra=()):
    hits = []
    for path in paths:
        if path in EXEMPT:
            continue
        try:
            with open(path, encoding="utf-8", errors="ignore") as handle:
                for number, line in enumerate(handle, 1):
                    term = first_match(line, extra)
                    if term:
                        hits.append((path, number, term, line.strip()[:90]))
        except OSError:
            continue
    return hits


def commit_messages(base=BASE):
    """[(short hash, full message)] for base..HEAD, or None when `base` is not known here."""
    known = subprocess.run(["git", "rev-parse", "--verify", "--quiet", base], capture_output=True, text=True)
    if known.returncode != 0:
        return None
    out = subprocess.run(["git", "log", "--format=%h%x00%B%x01", f"{base}..HEAD"],
                         capture_output=True, text=True, encoding="utf-8", errors="replace", check=True).stdout
    found = []
    for record in out.split("\x01"):
        if "\x00" in record:
            short, message = record.split("\x00", 1)
            found.append((short.strip(), message))
    return found


def scan_messages(messages, extra=()):
    hits = []
    for short, message in messages:
        for line in message.splitlines():
            term = first_match(line, extra)
            if term:
                hits.append((short, term, line.strip()[:90]))
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

    # The local patterns file: comments and blanks skipped, case-sensitive by default, a bad
    # line refused, and no file at all meaning no extra patterns.
    import os
    import tempfile
    with tempfile.TemporaryDirectory() as tmp:
        path = os.path.join(tmp, "patterns")
        with open(path, "w", encoding="utf-8") as handle:
            handle.write("# a comment\n\nSecret-Sibling\n(?i)codename-x\n")
        extra = load_local_patterns(path)
        if len(extra) != 2:
            bad.append(f"the patterns file gave {len(extra)} patterns, not 2")
        if first_match("we ported it from Secret-Sibling last week", extra) != "Secret-Sibling":
            bad.append("a local pattern did not match its name")
        if first_match("a secret-sibling, in lower case", extra) is not None:
            bad.append("a local pattern ignored case without (?i)")
        if first_match("CODENAME-X shipped", extra) is None:
            bad.append("a (?i) local pattern did not ignore case")
        if load_local_patterns(os.path.join(tmp, "absent")) != []:
            bad.append("a missing patterns file did not mean no extra patterns")
        with open(path, "w", encoding="utf-8") as handle:
            handle.write("broken[\n")
        try:
            load_local_patterns(path)
            refused = False
        except SystemExit:
            refused = True
        if not refused:
            bad.append("a line that is not a regular expression was accepted")

    messages = [("abc1234", "docs: tidy the README\n\nNothing private here.\n"),
                ("def5678", "fix: port the parser\n\nMoved over from Secret-Sibling.\n")]
    hits = scan_messages(messages, [re.compile("Secret-Sibling")])
    if [(short, term) for short, term, _ in hits] != [("def5678", "Secret-Sibling")]:
        bad.append(f"the commit-message scan found {hits}, not the one name in def5678's body")
    if scan_messages([("aaa1111", "chore: drop the AutoLnD reference\n")]) == []:
        bad.append("the commit-message scan missed a built-in name in a subject")
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

    extra = load_local_patterns(LOCAL_PATTERNS)
    print(f"{len(PRIVATE)} built-in names, and {len(extra)} pattern(s) from {LOCAL_PATTERNS}"
          + ("" if extra else " (none on this machine: only the built-in names are looked for)"))

    hits = scan(files, extra)
    messages = commit_messages()
    message_hits = scan_messages(messages, extra) if messages is not None else []
    if messages is None:
        print(f"commit messages NOT scanned: {BASE} is not known in this checkout.")
    else:
        print(f"{len(messages)} commit message(s) in {BASE}..HEAD scanned.")

    if not hits and not message_hits:
        print(f"{len(files)} tracked file(s) scanned; none names a private sibling repository.")
        return 0

    print()
    for path, number, term, line in hits:
        print(f"::error file={path},line={number}::names the private repository '{term}': {line}")
    for short, term, line in message_hits:
        print(f"::error::commit {short}'s message names the private repository '{term}': {line}")
    print()
    print("DocToolkit is PUBLIC and the four repositories beside it are not. A name in a doc "
          "comment is worse than one in a script: it compiles into DocToolkit.xml, ships inside "
          "the .nupkg, and renders on the API site - which is exactly how one reached nuget.org "
          "0.41.0. Move the reference into the gitignored material, or say the same thing without "
          "naming private work.")
    return 1


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
