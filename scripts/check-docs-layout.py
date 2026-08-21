#!/usr/bin/env python3
"""Guard the documentation layout: BACKLOG.html, its specs, and every referenced path.

WHY THIS RUNS IN TWO MODES, which is the unusual part.

The backlog and its specs live under paths this repository GITIGNORES - `BACKLOG.html` and
`docs/` - because DocToolkit is public and a backlog is not a consumer's business. PUBLIC.md
carries that reasoning. So CI can never see them, and a guard that quietly passed because the
files it guards are absent would be the exact silent-success shape this repo has corrected
twice this month.

So it runs either way and says which:

  LOCAL - the full check. The backlog is on disk, so its rows, its spec links, its markdown
          links and every path reference are resolved.
  CI    - the tracked half only: no COMMITTED file may reference a path that does not exist.

It refuses to pass vacuously in both modes - a walk that scanned nothing is a failure - and an
unreadable file is a failure rather than a skip.

Usage:
    python scripts/check-docs-layout.py
"""
import io
import os
import re
import subprocess
import sys
from html.parser import HTMLParser

sys.stdout.reconfigure(encoding="utf-8")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BACKLOG = os.path.join(ROOT, "BACKLOG.html")
SPECS = os.path.join(ROOT, "docs", "superpowers", "specs")

SKIP_DIRS = {".git", "node_modules", "bin", "obj", "_site", "artifacts", "TestResults",
             "StrykerOutput", "BenchmarkDotNet.Artifacts", ".worktrees", "packages"}

LINK = re.compile(r"\[[^\]]*\]\(([^)\s]+?)(?:\s+\"[^\"]*\")?\)")
ROW = re.compile(r"<tr data-id=\"([^\"]+)\"")
ROWLINK = re.compile(r"<tr data-id=\"[^\"]+\"[^>]*>\s*<td class=\"id\"><a href=\"([^\"]+)\"")
SCRIPT = re.compile(r"<script\b.*?</script>", re.S | re.I)
HEADING = re.compile(r"#\s+([A-Z]+[0-9][A-Za-z0-9-]*)\s*$")

# A path-shaped token in a comment or in prose.
#
# SCOPED TO docs/ DELIBERATELY. This guard owns the documentation layout; source paths are
# already held by the compiler and the other derived checks. A wider net was tried and produced
# only false positives - both bugs below were found by running it, not by reading it.
#
# (?<![A-Za-z0-9-]) because "learning-docs/x.html" CONTAINS "docs/x.html", and a bare word
# boundary matches happily after the hyphen. That is the same prefix collision that makes a
# naive A1 match inside A18.
#
# (?![A-Za-z0-9]) on the extension because otherwise a dotted DIRECTORY name reads as a file:
# "src/DocToolkit.Extensions.DependencyInjection" was reported as "...Extensions.Depend". Every
# candidate there is followed by another letter, so a single character class settles it.
#
# IT SAID (?![A-Za-z0-9.]) FOR ONE REVISION, AND THAT WAS WORSE THAN THE BUG IT FIXED. Excluding
# a following period also excludes a path that ENDS A SENTENCE - "...-design.md." - and two real
# references went silently unchecked. A guard that misses is worse than a guard that shouts,
# because nothing tells you. Found by testing the guard against known references rather than by
# reading it.
# The trailing pair reads as: the extension must not continue as more letters, AND must not be
# followed by another dotted segment. That second half is what tells "docs/x.md." (a path ending
# a sentence - allowed) from "docs/Foo.Bar.Baz/x" (a dotted DIRECTORY - not a file at all).
PATHREF = re.compile(
    r"(?<![A-Za-z0-9-])(docs/[A-Za-z0-9._\-/]+\.[A-Za-z0-9]{1,6})(?![A-Za-z0-9])(?!\.[A-Za-z0-9])")

problems = []
counts = {"md_files": 0, "links": 0, "rows": 0, "src_files": 0, "path_refs": 0, "specs": 0,
          "unpublished_refs": 0, "rendered_rows": 0}


def fail(where, what):
    problems.append((where, what))


def is_ignored(relpath):
    """Is this path one the repository deliberately does not publish?

    Asked of git rather than of a hand-written list, so it cannot drift from .gitignore - the
    same principle as reading the lockfile instead of keeping a table of packages.

    check-ignore exits 0 for ignored, 1 for not ignored, and 128 when it cannot answer. Only a
    clean 0 counts: an error must not be read as permission to skip.
    """
    r = subprocess.run(["git", "check-ignore", "-q", "--no-index", relpath],
                       cwd=ROOT, capture_output=True, text=True)
    return r.returncode == 0


def read(path):
    """An unreadable file is a FAILURE, not a skip - that is the whole point of the guard."""
    try:
        return io.open(path, encoding="utf-8", errors="strict").read()
    except Exception as exc:
        fail(os.path.relpath(path, ROOT), "cannot be read as UTF-8: " + str(exc))
        return None


def walk(exts, tracked_only):
    if tracked_only:
        out = subprocess.run(["git", "ls-files"], cwd=ROOT, capture_output=True, text=True)
        for rel in out.stdout.splitlines():
            if rel.endswith(exts):
                yield os.path.join(ROOT, rel)
        return
    for dirpath, dirnames, filenames in os.walk(ROOT):
        dirnames[:] = [d for d in dirnames if d not in SKIP_DIRS]
        for fn in filenames:
            if fn.endswith(exts):
                yield os.path.join(dirpath, fn)


def check_backlog():
    """Rows resolve, and no ticket data hides inside a script block."""
    html = read(BACKLOG)
    if html is None:
        return

    ids = ROW.findall(html)
    counts["rows"] = len(ids)
    if not ids:
        fail("BACKLOG.html", "contains no rows - either it is empty, or the rows stopped being "
                             "in the markup, and an empty backlog page looks exactly like a "
                             "backlog with nothing in it")
        return

    stripped = SCRIPT.sub("", html)
    if len(ROW.findall(stripped)) != len(ids):
        fail("BACKLOG.html", "rows disappear when the script is removed - ticket data must live "
                             "in the markup so a broken script cannot render the page empty")

    script = "".join(SCRIPT.findall(html))
    leaked = [i for i in ids
              if re.search(r"(?<![A-Za-z0-9])" + re.escape(i) + r"(?![A-Za-z0-9])", script)]
    if leaked:
        fail("BACKLOG.html", "ticket ids appear inside the script block: " + str(leaked[:5]))

    hrefs = ROWLINK.findall(html)
    if len(hrefs) != len(ids):
        fail("BACKLOG.html", "%d rows but %d spec links - some row carries none"
             % (len(ids), len(hrefs)))
    for href in hrefs:
        if not os.path.exists(os.path.normpath(os.path.join(ROOT, href))):
            fail("BACKLOG.html", "row link does not resolve: " + href)


class MarkupIndex(HTMLParser):
    """Collects the structure the page's correctness depends on.

    CLASS TOKENS, NOT SELECTOR TEXT, and that is the whole reason a parser is used here. A grep
    written while scoping this reported td.era, td.id and td.sec as undefined because it only
    matched selectors beginning with a dot. To a token collector, "td.id" and ".id" are the same
    class - which is the question actually being asked.
    """

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.classes = set()
        self.rows = []
        self.groups = []
        self._group = None
        self._row = None
        self._cell = None
        self._in_summary = False
        self._in_count = False
        self._summary_text = []
        self._count_text = []

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        for token in (a.get("class") or "").split():
            self.classes.add(token)

        if tag == "details":
            self._group = {"id": a.get("id", ""), "status": "", "claimed": -1, "rows": 0}
            self.groups.append(self._group)
            self._summary_text = []
        elif tag == "summary":
            self._in_summary = True
        elif tag == "span" and "count" in (a.get("class") or "").split():
            self._in_count = True
            self._count_text = []
        elif tag == "tr" and "data-id" in a:
            self._row = {"id": a["data-id"], "href": None, "cells": [],
                         "group": self._group["id"] if self._group else None}
            self.rows.append(self._row)
            if self._group is not None:
                self._group["rows"] += 1
        elif tag == "td":
            self._cell = []
        elif tag == "a" and self._row is not None and self._row["href"] is None:
            self._row["href"] = a.get("href")

    def handle_endtag(self, tag):
        if tag == "summary":
            self._in_summary = False
            if self._group is not None:
                self._group["status"] = " ".join("".join(self._summary_text).split())
        elif tag == "span" and self._in_count:
            self._in_count = False
            digits = "".join(ch for ch in "".join(self._count_text) if ch.isdigit())
            if self._group is not None and digits:
                self._group["claimed"] = int(digits)
        elif tag == "td" and self._row is not None:
            self._row["cells"].append("".join(self._cell or []).strip())
            self._cell = None
        elif tag == "tr":
            self._row = None
        elif tag == "details":
            self._group = None

    def handle_data(self, data):
        if self._in_count:
            self._count_text.append(data)
        elif self._in_summary:
            self._summary_text.append(data)
        if self._cell is not None:
            self._cell.append(data)


# Classes applied for scripting or semantics rather than styling. Every entry is a hole in the
# check - argue for additions rather than adding one to make a run green.
UNSTYLED_BY_DESIGN = set()

# The five statuses, from CLAUDE.md's table. A sixth means a typo reached the generator.
KNOWN_STATUSES = {"OPEN", "ACCEPTED", "DONE", "SUSPENDED", "DROPPED"}

STYLE = re.compile(r"<style\b[^>]*>(.*?)</style>", re.S | re.I)
SELECTOR_CLASS = re.compile(r"\.([A-Za-z][A-Za-z0-9_-]*)")


def inspect_markup(html, where):
    """Is the page self-consistent? Distinct from check_backlog, which asks if links RESOLVE.

    Takes the markup as a string rather than a path so the self-test can exercise it without
    touching the filesystem - the same seam that lets every assertion be observed failing.

    IT CANNOT SEE APPEARANCE. There is no browser here, deliberately, so this checks structure:
    a class nothing styles, a count that disagrees with its rows, a row that would render blank.
    A colour that vanishes in dark mode is out of reach and is not claimed.
    """
    index = MarkupIndex()
    index.feed(html)
    counts["rendered_rows"] = len(index.rows)

    defined = set(SELECTOR_CLASS.findall("".join(STYLE.findall(html))))
    orphans = sorted(index.classes - defined - UNSTYLED_BY_DESIGN)
    if orphans:
        fail(where, "class(es) used in the markup but defined in no stylesheet rule: "
             + ", ".join(orphans))

    if not index.groups:
        fail(where, "no <details> groups - the page has lost its structure")

    for group in index.groups:
        # A count is a claim the page makes about itself, and the cheapest kind to get wrong:
        # regenerate from a filtered list and the summary keeps the old number.
        if group["claimed"] != group["rows"]:
            fail(where, "group %s claims %d row(s) but contains %d"
                 % (group["id"] or "?", group["claimed"], group["rows"]))

        status = group["status"].split()[0] if group["status"] else ""
        if status and status not in KNOWN_STATUSES:
            fail(where, "group %s is headed %r, which is not a known status - a typo reaching "
                 "the generator would silently create a sixth group" % (group["id"], status))
        # The status colours are keyed on the id, so a mismatch paints the wrong group and
        # nothing else goes wrong - the page looks fine and means something else.
        elif status and group["id"] != "g-" + status.lower():
            fail(where, "group id %s does not match its status %s - the status colours are keyed "
                 "on the id, so the wrong group would be coloured" % (group["id"], status))

    for row in index.rows:
        if not row["href"]:
            fail(where, "row %s has no spec link - it would render as plain text" % row["id"])
        if len([c for c in row["cells"] if c]) < 2:
            fail(where, "row %s has fewer than two non-empty cells - it would render as a "
                 "near-blank line" % row["id"])

    # THE MARKUP MUST CARRY ROWS AT ALL - and this is deliberately not phrased as "compare the
    # row count with and without the script", which is what it said first.
    #
    # That comparison cannot work here and the self-test proved it. html.parser treats <script>
    # as CDATA, so a <tr> hidden in a JS string is invisible to BOTH parses and the counts match.
    # Worse, a page whose rows were BUILT by script would parse to zero rows either way and sail
    # through - the very failure the comparison was meant to catch.
    #
    # So the DOM-level question is "does the markup itself contain rows", and the leak question -
    # do ticket ids appear inside the script - stays in check_backlog, where a regex CAN see into
    # the script block. Two checks, each asking what its own tool can actually answer.
    if index.groups and not index.rows:
        fail(where, "the page has groups but no rows in the markup - if the rows are built by "
             "script, a reader with JavaScript off sees an empty backlog, which looks exactly "
             "like a backlog with nothing in it")

    return index


def check_backlog_renders():
    """Read the page and inspect it. Split from inspect_markup so the self-test can drive the
    assertions against a sample string without needing a file on disk."""
    html = read(BACKLOG)
    if html is not None:
        inspect_markup(html, "BACKLOG.html")


def check_specs():
    """Each spec declares its own id in its first heading - never parse it from the filename."""
    if not os.path.isdir(SPECS):
        fail("docs/superpowers/specs", "does not exist")
        return
    names = [f for f in sorted(os.listdir(SPECS)) if f.endswith(".md")]
    if not names:
        fail("docs/superpowers/specs", "holds no specs - nothing was checked")
        return
    seen = {}
    for name in names:
        text = read(os.path.join(SPECS, name))
        if text is None:
            continue
        counts["specs"] += 1
        # Only a TICKET spec is named "<ID>-slug.md". The dated design and decision documents
        # that predate this layout keep their own prose titles - the playbook's D4 says link
        # such documents rather than convert them, and demanding an id heading of them would be
        # inventing a rule to satisfy a checker.
        if not re.match(r"^[A-Z]+[0-9][A-Za-z0-9-]*-", name):
            continue
        first = text.split("\n", 1)[0].strip()
        m = HEADING.match(first)
        if not m:
            fail("docs/superpowers/specs/" + name, "first line is not a '# <ID>' heading")
            continue
        tid = m.group(1)
        if tid in seen:
            fail("docs/superpowers/specs/" + name, "id " + tid + " already declared by " + seen[tid])
        seen[tid] = name


def strip_code(text):
    """Remove fenced blocks and inline code spans.

    A link inside code is NOT a link, and treating one as a link is how a checker becomes a
    liar. Every false positive in the first run of this guard was of that kind: `[LIBRARY](LINK)`
    documenting an awesome-list entry format, `([#N](...))` documenting what release-please
    emits, and `![](../../etc/passwd)` naming a path this library deliberately refuses. Each is
    prose ABOUT a link, and none of them should resolve to anything.
    """
    text = re.sub(r"```.*?```", "", text, flags=re.S)
    text = re.sub(r"~~~.*?~~~", "", text, flags=re.S)
    return re.sub(r"`[^`\n]*`", "", text)


def check_markdown_links(tracked_only):
    """Resolve every relative link against its CONTAINING file - spelling checks miss moves."""
    for path in walk((".md",), tracked_only):
        text = read(path)
        if text is None:
            continue
        counts["md_files"] += 1
        base = os.path.dirname(path)
        for target in LINK.findall(strip_code(text)):
            if target.startswith(("http://", "https://", "mailto:", "#", "@", "xref:", "data:")):
                continue
            head = target.split("#")[0]
            if not head:
                continue
            counts["links"] += 1
            if not os.path.exists(os.path.normpath(os.path.join(base, head))):
                fail(os.path.relpath(path, ROOT), "link does not resolve: " + target)


def check_path_references(tracked_only):
    """Scan SOURCE too. A stale docs/ path in a code comment is a dangling reference."""
    for path in walk((".cs", ".py", ".yml", ".yaml", ".mjs", ".csproj", ".props"), tracked_only):
        if os.path.abspath(path) == os.path.abspath(__file__):
            continue
        text = read(path)
        if text is None:
            continue
        counts["src_files"] += 1
        # NOT joined across lines, and that is a reversal worth recording. Joining continuation
        # lines is the textbook fix for a soft-wrapped path - but here it FABRICATED references
        # that appear nowhere in the file, inventing one out of two adjacent YAML keys. Inventing
        # a reference is a worse failure than missing one, so this matches what is written.
        for ref in sorted(set(PATHREF.findall(text))):
            counts["path_refs"] += 1
            # A path in a csproj or a comment is as often relative to ITS OWN directory as to
            # the repository root. Only a reference that resolves from neither is a candidate.
            here = os.path.dirname(path)
            if (os.path.exists(os.path.join(ROOT, ref))
                    or os.path.exists(os.path.join(here, ref))):
                continue

            # THE ESCAPE HATCH BELONGS TO CI ONLY, and getting that wrong made this check
            # vacuous in BOTH modes for one revision.
            #
            # On a runner, docs/ is simply not there - it is gitignored because this repo is
            # public (PUBLIC.md). A comment citing docs/superpowers/specs/x.md points at a real
            # file a maintainer has and a consumer does not, so failing would make CI red
            # forever, and a permanently red guard gets deleted rather than fixed.
            #
            # But LOCALLY the file is right there, so absence means dangling and nothing else.
            # The first version skipped ignored paths in both modes - which meant a citation of
            # a spec that never existed passed everywhere. Two sabotages walked straight through
            # it. The distinction is not "is it ignored" but "can this run see it at all".
            if tracked_only and is_ignored(ref):
                counts["unpublished_refs"] += 1
                continue

            fail(os.path.relpath(path, ROOT), "references a path that does not exist: " + ref)


def main():
    local = os.path.exists(BACKLOG) and os.path.isdir(SPECS)
    print("mode: " + ("LOCAL - backlog present, checking everything"
                      if local else
                      "CI - backlog is gitignored and absent, checking tracked files only"))

    if local:
        check_backlog()
        check_backlog_renders()
        check_specs()
    check_markdown_links(tracked_only=not local)
    check_path_references(tracked_only=not local)

    if counts["md_files"] == 0 or counts["src_files"] == 0:
        print("FAIL: scanned %d markdown and %d source files - a walk that finds nothing "
              "proves nothing" % (counts["md_files"], counts["src_files"]))
        return 2
    if local and counts["rows"] == 0:
        print("FAIL: the backlog was present but yielded no rows")
        return 2
    if local and counts["rendered_rows"] == 0:
        print("FAIL: the backlog parsed to zero rows - the render check inspected nothing")
        return 2

    print("  %d rows, %d specs, %d markdown files, %d links, %d source files, %d path refs "
          "(%d of them deliberately unpublished)"
          % (counts["rows"], counts["specs"], counts["md_files"], counts["links"],
             counts["src_files"], counts["path_refs"], counts["unpublished_refs"]))

    # Say what was NOT checked. A guard that reports only its successes reads as broader
    # coverage than it has, which is how a check becomes a claim nobody verified.
    if counts["unpublished_refs"]:
        print("  NOT VERIFIED here: %d reference(s) to gitignored paths - this run cannot see "
              "them.\n  Run locally, where those files exist, to check them."
              % counts["unpublished_refs"])

    if problems:
        print("\n%d problem(s):" % len(problems))
        for where, what in problems:
            print("  " + where + "\n     " + what)
        return 1

    print("documentation layout is consistent")
    return 0


# ---------------------------------------------------------------------------------------------
# Self-test. Same convention as check-changelog-curated.py and check-render-policy.py: the
# controls ship with the guard and run wherever it runs, rather than in a test framework this
# repository does not install.
#
# A POSITIVE CONTROL IS REQUIRED, not optional. Without one, a checker that rejected everything
# would pass every negative case and look thorough. That is the failure mode this whole file
# exists to avoid, so it would be a poor thing to build into its own tests.
# ---------------------------------------------------------------------------------------------

SAMPLE_PAGE = """<!doctype html>
<style>
details.group{}
summary{}
.count{}
td.id{}
td.sum{}
.wrap{}
</style>
<div class="wrap">
<details class="group" id="g-open"><summary>OPEN <span class="count">(2)</span></summary>
<table><tbody>
<tr data-id="A1" data-status="OPEN"><td class="id"><a href="docs/a.md">A1</a></td><td class="sum">first headline</td></tr>
<tr data-id="A2" data-status="OPEN"><td class="id"><a href="docs/b.md">A2</a></td><td class="sum">second headline</td></tr>
</tbody></table></details>
</div>
<script>try{ /* additive only */ }catch(e){}</script>
"""


def self_test():
    """Every assertion is observed failing before it is trusted."""
    cases = [
        ("POSITIVE CONTROL - a well-formed page is accepted",
         lambda t: t, None),
        ("a class used but styled by nothing",
         lambda t: t.replace('class="sum"', 'class="orphan"', 1),
         "defined in no stylesheet rule"),
        ("a group count that disagrees with its rows",
         lambda t: t.replace('(2)', '(9)', 1),
         "claims 9"),
        ("a row that lost its link",
         lambda t: t.replace('<a href="docs/a.md">A1</a>', 'A1', 1),
         "no spec link"),
        ("a group id that contradicts its status",
         lambda t: t.replace('id="g-open"', 'id="g-done"', 1),
         "does not match its status"),
        ("a status outside the known vocabulary",
         lambda t: t.replace('<summary>OPEN ', '<summary>BANANA ', 1)
                    .replace('id="g-open"', 'id="g-banana"', 1),
         "not a known status"),
        ("rows built by script instead of markup",
         lambda t: SCRIPT.sub("", t.replace('<tr data-id="A1"', '<!--x', 1)
                                    .replace('<tr data-id="A2"', '<!--y', 1)),
         "no rows in the markup"),
    ]

    failures = 0
    for name, mutate, expect in cases:
        del problems[:]
        inspect_markup(mutate(SAMPLE_PAGE), "SAMPLE")
        reported = " | ".join(what for _, what in problems)
        if expect is None:
            ok = not problems
            detail = "" if ok else "unexpectedly reported: " + reported
        else:
            ok = expect.lower() in reported.lower()
            detail = "" if ok else "expected %r, got: %s" % (expect, reported or "nothing")
        failures += 0 if ok else 1
        print(("  ok    " if ok else "  FAIL  ") + name)
        if detail:
            print("          " + detail)

    del problems[:]
    print()
    if failures:
        print("%d self-test case(s) failed - do not trust this guard" % failures)
        return 1
    print("all %d self-test cases behave as expected" % len(cases))
    return 0


if "--self-test" in sys.argv:
    sys.exit(self_test())

sys.exit(main())
