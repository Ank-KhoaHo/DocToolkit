#!/usr/bin/env python3
"""Assert every `*Core` method is genuinely SHARED by both overloads.

WHY THIS EXISTS. `CLAUDE.md` states the convention as: "Every `*Core` method holds the one real
implementation that both the `byte[]` and `Stream` overloads call, so the two can never drift
apart." That is the whole point of the pattern - one implementation means the two public shapes
cannot disagree about what an operation does.

Measured 2026-08-14, it was not literally true. Of 17 `*Core` methods, 15 were called from two or
more places and **two were called exactly once** - `WorkbookEditor.SetCellCore` and
`WorkbookEditor.FormatCore`, both only from the `Stream` overload. Their `byte[]` overloads had
re-implemented the same open/apply/save/wrap sequence inline, which is precisely the drift the
convention exists to prevent: the duplicate opened the workbook through a different helper, so the
two paths had already begun to differ in how they got there.

Nothing detected that. The convention lived in prose, and prose does not fail a build. This does.

WHAT IT CHECKS, and what it deliberately does not.

A `*Core` method with ONE caller is a private method with a misleading name: it promises shared
implementation and delivers a single-use helper. So the rule is simply "at least two call sites",
which is the property the name claims. It says nothing about WHICH callers - a `*Core` shared by
two `Stream` overloads is still shared, and inventing a rule about caller shapes would encode
today's overload set as if it were the convention.

Counting is textual, over the source rather than the compiled assembly, because the failure being
prevented is somebody writing a second implementation - and at that moment the code may not even
build. A definition line is excluded from its own count.

**Counted per FILE, not per name, and that is load-bearing.** Three names are defined in two
different classes each - `ExtractTextCore`, `ReplaceTextCore` and `ReplaceImageCore` all exist in
both `DocxEditor` and `PresentationEditor` - so 14 distinct names are 17 definitions. Keying by
name alone sums their call sites together, and one class's callers would then mask another class
having none. Since every `*Core` is `private static`, its callers are necessarily in its own file,
so per-file counting is both correct and the stricter reading.

Known limit, stated rather than papered over: two OVERLOADS sharing one name in one file (today,
`WorkbookEditor.CreateCore`) are counted as a single unit, so two definitions with one caller each
would pass. Fixing that needs signature-level parsing, which is a large amount of fragility to buy
a case that has not happened.

USAGE

    python scripts/check-core-sharing.py          # names every offender, exits 1 if any
"""
from __future__ import annotations

import pathlib
import re
import sys

REPO = pathlib.Path(__file__).resolve().parent.parent
SRC = REPO / "src"

DEFINITION = re.compile(
    r"^\s*(?:private|internal|public|protected)\s+(?:static\s+)?(?:async\s+)?"
    r"[\w<>\[\],\.\s\?]+?\s(\w+Core)\s*\(",
    re.MULTILINE,
)


def main() -> int:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    files = sorted(SRC.rglob("*.cs"))
    if not files:
        sys.exit("error: no C# sources found under src/ - this check is looking in the wrong place "
                 "and would otherwise pass by finding nothing.")

    sources = {p: p.read_text(encoding="utf-8") for p in files}

    # (file, name) -> how many times that file DEFINES it. Per file, never per name: see the
    # module docstring for why summing across files would let one class mask another.
    definitions: dict[tuple[pathlib.Path, str], int] = {}
    for path, text in sources.items():
        for name in DEFINITION.findall(text):
            definitions[(path, name)] = definitions.get((path, name), 0) + 1

    if not definitions:
        sys.exit("error: found no *Core methods at all. Either the convention was renamed or this "
                 "parser is blind; a check that silently matches nothing is worse than none.")

    offenders: list[tuple[str, str, int]] = []
    for (path, name), defined in sorted(definitions.items(), key=lambda kv: (kv[0][0].name, kv[0][1])):
        text = sources[path]
        calls = 0
        for match in re.finditer(rf"\b{re.escape(name)}\s*\(", text):
            line_start = text.rfind("\n", 0, match.start()) + 1
            line = text[line_start:text.find("\n", match.start())]
            if DEFINITION.match(line + "("):  # the definition itself, not a call
                continue
            calls += 1
        if calls < 2 * defined:
            offenders.append((name, path.relative_to(REPO).as_posix(), calls))

    total = sum(definitions.values())
    print(f"{total} *Core definition(s) across {len(definitions)} (file, name) pair(s), "
          f"in {len(files)} file(s) under src/")

    if offenders:
        for name, where, calls in offenders:
            print(f"  {name:<28} {where}  <-- {calls} call site(s)")
        print(
            "\n::error::A *Core method with fewer than two call sites is not shared, which is the "
            "one thing its name promises. Route the other overload through it - or, if it genuinely "
            "has a single caller, give it a name that does not claim otherwise.")
        return 1

    print("every *Core method is called from at least two places")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
