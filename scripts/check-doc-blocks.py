#!/usr/bin/env python3
"""Fail when one XML doc comment carries two <summary> or two <remarks> blocks.

WHY THIS IS NOT CAUGHT ALREADY, which is the whole reason it exists.

The C# compiler does not diagnose a duplicated documentation tag. A member can carry two
<summary> blocks and the build stays green under -warnaserror, while DocFX renders whichever one
it picks - so the API site describes the wrong thing and nothing anywhere says so.

Both instances found on 2026-08-20 were the same accident: a doc block copied from a neighbouring
overload and left stacked above the real one. In HtmlToPdfConverter the stray block described the
overload that takes no fonts; in DocxPdfFailureDiagnosis it belonged to a method further down,
which was left with no documentation at all. Neither is visible in a diff unless you already know
to look, because both files legitimately contain many <summary> tags.

Same standard as check-doc-snippets.py, which made an empty <code source> region a CI failure
rather than a habit: if it is cheap to detect, it should not be somebody's job to remember.

Usage:
    python scripts/check-doc-blocks.py
    python scripts/check-doc-blocks.py --self-test
"""
import io
import os
import re
import sys

sys.stdout.reconfigure(encoding="utf-8")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SOURCES = [os.path.join(ROOT, "src"), os.path.join(ROOT, "tests"), os.path.join(ROOT, "samples")]
SKIP_DIRS = {"bin", "obj", "TestResults", "StrykerOutput", "BenchmarkDotNet.Artifacts"}

# Only the tags that must appear at most once per member. <param>, <exception> and <typeparam>
# are legitimately repeated, and <example> may be - none of them belong here.
SINGLETON = ("summary", "remarks", "returns", "value")

OPEN = re.compile(r"^\s*///\s*<(" + "|".join(SINGLETON) + r")[\s>]")


def blocks(lines):
    """Yield (start_line, [tag, ...]) for each run of consecutive /// lines."""
    start, tags = None, []
    for i, line in enumerate(lines):
        if line.lstrip().startswith("///"):
            if start is None:
                start, tags = i, []
            m = OPEN.match(line)
            if m:
                tags.append(m.group(1))
        elif start is not None:
            yield start, tags
            start, tags = None, []
    if start is not None:
        yield start, tags


def inspect(text, where):
    """Return a list of problems. Split out so the self-test can drive it on a string."""
    found = []
    for start, tags in blocks(text.split("\n")):
        for tag in SINGLETON:
            n = tags.count(tag)
            if n > 1:
                found.append("%s:%d  one doc comment carries %d <%s> blocks - the compiler does "
                             "not diagnose this and DocFX renders whichever it picks"
                             % (where, start + 1, n, tag))
    return found


def main():
    problems, scanned = [], 0

    for root in SOURCES:
        if not os.path.isdir(root):
            continue
        for dirpath, dirnames, filenames in os.walk(root):
            dirnames[:] = [d for d in dirnames if d not in SKIP_DIRS]
            for fn in filenames:
                if not fn.endswith(".cs"):
                    continue
                path = os.path.join(dirpath, fn)
                try:
                    text = io.open(path, encoding="utf-8").read()
                except Exception as exc:
                    problems.append("%s  cannot be read: %s" % (path, exc))
                    continue
                scanned += 1
                problems += inspect(text, os.path.relpath(path, ROOT).replace("\\", "/"))

    # A walk that read nothing proves nothing.
    if scanned == 0:
        print("FAIL: scanned 0 source files")
        return 2

    if problems:
        print("%d problem(s) in %d file(s):" % (len(problems), scanned))
        for p in problems:
            print("  " + p)
        return 1

    print("no doc comment carries a duplicated <%s> block (%d files)"
          % (">/<".join(SINGLETON), scanned))
    return 0


# The controls ship with the guard and run wherever it runs - the convention the other scripts
# here already follow. The POSITIVE one is required: without it, a checker that flagged everything
# would pass every negative case and look thorough.
SAMPLES = [
    ("POSITIVE CONTROL - an ordinary member is accepted", """
    /// <summary>Does a thing.</summary>
    /// <param name="a">First.</param>
    /// <param name="b">Second.</param>
    /// <exception cref="ArgumentException">a is empty.</exception>
    /// <exception cref="InvalidOperationException">b is closed.</exception>
    /// <remarks>Worth knowing.</remarks>
    public void Thing(int a, int b) { }
""", False),

    ("two <summary> blocks stacked", """
    /// <summary>The stray one, copied from a neighbour.</summary>
    /// <summary>The real one.</summary>
    public void Thing() { }
""", True),

    ("two <remarks> blocks stacked", """
    /// <summary>Fine.</summary>
    /// <remarks>First.</remarks>
    /// <remarks>Second.</remarks>
    public void Thing() { }
""", True),

    ("multi-line blocks, the real shape of both offenders", """
    /// <summary>
    /// Converts a thing.
    /// </summary>
    /// <remarks>
    /// Something about it.
    /// </remarks>
    /// <summary>
    /// Converts a different thing.
    /// </summary>
    public void Thing() { }
""", True),

    ("two members, one <summary> each - NOT a duplicate", """
    /// <summary>First member.</summary>
    public void One() { }

    /// <summary>Second member.</summary>
    public void Two() { }
""", False),

    ("repeated <param> and <exception> are legitimate", """
    /// <summary>Fine.</summary>
    /// <param name="a">A.</param>
    /// <param name="b">B.</param>
    /// <exception cref="ArgumentException">One.</exception>
    /// <exception cref="ArgumentNullException">Two.</exception>
    public void Thing(int a, int b) { }
""", False),
]


def self_test():
    failures = 0
    for name, text, should_fail in SAMPLES:
        problems = inspect(text, "SAMPLE")
        ok = bool(problems) == should_fail
        failures += 0 if ok else 1
        print(("  ok    " if ok else "  FAIL  ") + name)
        if not ok:
            print("          expected %s, got: %s"
                  % ("a problem" if should_fail else "no problem", problems or "nothing"))
    print()
    if failures:
        print("%d self-test case(s) failed - do not trust this guard" % failures)
        return 1
    print("all %d self-test cases behave as expected" % len(SAMPLES))
    return 0


sys.exit(self_test() if "--self-test" in sys.argv else main())
