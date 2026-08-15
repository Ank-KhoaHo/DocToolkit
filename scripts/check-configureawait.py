#!/usr/bin/env python3
"""Fail when a shipped `await` does not carry ConfigureAwait(false).

WHY A LIBRARY CARES. When a consumer blocks on one of our async methods from a
context that has a SynchronizationContext - WPF, WinForms, classic ASP.NET - an
unconfigured continuation tries to resume on that context while the caller is
holding it, and the call deadlocks. `ConfigureAwait(false)` is what prevents it,
and it is a library's responsibility rather than the caller's: the caller cannot
add it to our awaits.

THIS IS NOT THE MUTATION-TESTING QUESTION, and someone will think it is.
CLAUDE.md and stryker-config.json record ConfigureAwait as an EQUIVALENT MUTANT,
excluded from mutation scoring, because flipping (false) to (true) cannot be
killed by a test running outside a synchronisation context. That is a statement
about TESTABILITY. This check is about a deadlock in a consumer's process. Both
are true; they are not in tension.

WHY DERIVED RATHER THAN REVIEWED. Measured 2026-08-15: 123 of 131 awaits in
src/ already had it. The convention was not in dispute - eight sites had simply
escaped it over time, in exactly the way a rule enforced only by review does.
Nothing in the build had an opinion, so the count could only ever drift up.

TWO AWAITS HIDE IN `await using var x = ...`. The declaration form has a
disposal await that no ConfigureAwait can reach, so it is reported even when the
initialiser is configured - which is how GuardedResourceLoader's stream read was
found, after a first pass that scanned whole statements had passed it. Write it
as a block instead:

    var source = File.OpenRead(path);
    await using (source.ConfigureAwait(false)) { ... }

COMMENTS AND STRINGS ARE STRIPPED FIRST, and that is not defensive coding: the
first run of this script reported a violation in `PdfEditor.cs` that was the
sentence above, sitting in a comment explaining the fix. A checker fooled by
prose about the thing it checks is worse than no checker.

USAGE

    python scripts/check-configureawait.py

Exits non-zero and lists file:line for every unconfigured await.
"""
from __future__ import annotations

import pathlib
import re
import sys

REPO = pathlib.Path(__file__).resolve().parent.parent
SRC = REPO / "src"

# Blank out comments and string/char literals, preserving newlines so line numbers
# survive. Order matters: verbatim and interpolated strings before line comments,
# or a `//` inside a literal path string truncates the rest of a real line.
LITERALS = re.compile(
    r'@"(?:[^"]|"")*"'      # verbatim string
    r"|\"(?:\\.|[^\"\\])*\""  # regular string
    r"|'(?:\\.|[^'\\])*'"     # char
    r"|/\*.*?\*/"             # block comment
    r"|//[^\n]*",             # line comment
    re.S,
)


def strip(text: str) -> str:
    return LITERALS.sub(lambda m: re.sub(r"[^\n]", " ", m.group(0)), text)


def check(path: pathlib.Path) -> list[tuple[int, str]]:
    raw = path.read_text(encoding="utf-8")
    code = strip(raw)
    lines = raw.split("\n")
    bad: list[tuple[int, str]] = []

    for m in re.finditer(r"\bawait\b", code):
        line_no = code[: m.start()].count("\n") + 1
        rest = code[m.start():]

        # `await using` splits into two shapes with different verdicts.
        using = re.match(r"await\s+using\s*(\(|var\b|[A-Za-z_])", rest)
        if using:
            if using.group(1) == "(":
                stmt = rest[: rest.find(")") + 1] if ")" in rest else rest
                if "ConfigureAwait" not in stmt:
                    bad.append((line_no, "await using block, disposal not configured"))
            else:
                bad.append((line_no,
                            "await using DECLARATION - its disposal await cannot be configured; "
                            "use `await using (x.ConfigureAwait(false)) { ... }`"))
            continue

        end = rest.find(";")
        stmt = rest if end == -1 else rest[:end]
        if "ConfigureAwait" not in stmt:
            bad.append((line_no, lines[line_no - 1].strip()[:90]))

    return bad


def main() -> int:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    files = sorted(SRC.rglob("*.cs"))
    if not files:
        sys.exit("error: no .cs files found under src/ - this check is now blind. "
                 "Fix the path rather than letting it report success.")

    total_awaits = 0
    failures: list[str] = []
    for f in files:
        total_awaits += len(re.findall(r"\bawait\b", strip(f.read_text(encoding="utf-8"))))
        for line_no, detail in check(f):
            failures.append(f"{f.relative_to(REPO).as_posix()}:{line_no}  {detail}")

    # Refuse to pass vacuously. If the tree really had no awaits the check proves
    # nothing, and that is indistinguishable from a broken scanner.
    if total_awaits == 0:
        sys.exit("error: found no awaits anywhere under src/ - the scanner is broken. "
                 "A check that inspects nothing must not report success.")

    print(f"{total_awaits} awaits in {len(files)} files under src/")

    if failures:
        print()
        for line in failures:
            print(f"  {line}")
        print(f"\n::error::{len(failures)} await(s) in shipped code do not carry "
              f"ConfigureAwait(false). A consumer blocking on one of these from a "
              f"WPF/WinForms/classic-ASP.NET context can deadlock.")
        return 1

    print("every await carries ConfigureAwait(false)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
