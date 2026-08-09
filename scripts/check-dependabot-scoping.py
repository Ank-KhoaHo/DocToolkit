#!/usr/bin/env python3
"""Fail when an update block that can reach src/ does not repeat src/'s ignore rules.

A Dependabot `ignore` rule is scoped to the block that declares it. Nothing is
inherited. That is easy to read past, because the file looks like one
configuration rather than four independent ones.

It is not hypothetical. Found 2026-08-03: a grouped run from the /tests/ block
proposed `SixLabors.Fonts [1.0.0] -> [3.0.0]` - straight past the revenue-gated
2.x line that the [1.0.0] pin exists to prevent - plus three shipped floors that
are ignored in their own blocks. The test projects reach src/ through
ProjectReference, and Dependabot follows those references and edits csproj files
belonging to other blocks, where the other blocks' ignore rules do not apply. The
same is true of the root block, which has the solution in view.

The backlog recorded the mitigation as "repeat all four guards in the tests
block" and then said the important part: *a test asserting that would be better
than a comment*. This is that test. The comment is still there; it explains why,
and this stops it from quietly becoming false.

WHAT IT CHECKS, and why this shape. The protected set is not a list maintained
here - it is DERIVED from the ignore rules the src/ blocks actually declare. So
adding a new pin to a src/ block automatically makes it required everywhere else,
which is exactly the failure being guarded: the rule that gets added in one place
and forgotten in the others.

A block may ignore MORE than required. The tests block ignores SixLabors.Fonts
entirely where src/ ignores only its majors; ignoring more is always safe, so
only the dependency name is compared, never the update-types.
"""

import io
import re
import sys

CONFIG = ".github/dependabot.yml"

BLOCK = re.compile(r"^  - package-ecosystem:\s*(\S+)")
DIRECTORY = re.compile(r"^    directory:\s*[\"']?([^\"'\s]+)")
DIRECTORIES_ITEM = re.compile(r"^      - [\"']?([^\"'\s]+)")
KEY = re.compile(r"^    ([\w-]+):")
IGNORED_NAME = re.compile(r"^\s+- dependency-name:\s*[\"']?([^\"'\s]+)")


def blocks(text):
    """Yield (ecosystem, [directories], {ignored names}) for each update block."""
    current = None
    section = None

    for line in text.splitlines():
        start = BLOCK.match(line)
        if start:
            if current:
                yield current
            current = (start.group(1), [], set())
            section = None
            continue

        if current is None:
            continue

        key = KEY.match(line)
        if key:
            section = key.group(1)
            directory = DIRECTORY.match(line)
            if directory:
                current[1].append(directory.group(1))
            continue

        if section == "directories":
            item = DIRECTORIES_ITEM.match(line)
            if item:
                current[1].append(item.group(1))
            continue

        if section == "ignore":
            name = IGNORED_NAME.match(line)
            if name:
                current[2].add(name.group(1))

    if current:
        yield current


def main():
    try:
        text = io.open(CONFIG, encoding="utf-8").read()
    except FileNotFoundError:
        sys.exit(f"::error::{CONFIG} not found; this check cannot run.")

    nuget = [b for b in blocks(text) if b[0] == "nuget"]
    if not nuget:
        sys.exit("::error::No nuget update blocks parsed from "
                 f"{CONFIG}. Either the file changed shape or this check stopped "
                 "matching it; both are worth a red build.")

    # The source of truth: whatever the src/ blocks protect.
    protects_src = [b for b in nuget if all(d.startswith("/src/") for d in b[1])]
    required = set().union(*(b[2] for b in protects_src)) if protects_src else set()

    if not required:
        sys.exit("::error::No ignore rules found on any /src/ block, so this check "
                 "would pass vacuously. If the pins were deliberately removed, remove "
                 "this check in the same commit and say why.")

    print(f"{len(nuget)} nuget block(s); {len(required)} rule(s) protect src/: "
          f"{', '.join(sorted(required))}")
    print()

    failures = []
    for ecosystem, directories, ignored in nuget:
        scope = ", ".join(directories) or "(unscoped)"
        if all(d.startswith("/src/") for d in directories):
            print(f"  {scope} - is src/, declares {len(ignored)} rule(s)")
            continue

        missing = sorted(required - ignored)
        status = "OK" if not missing else f"MISSING {', '.join(missing)}"
        print(f"  {scope} - reaches src/, {status}")

        for name in missing:
            failures.append(
                f"{scope} can reach src/ but does not ignore `{name}`")

    print()
    if not failures:
        print("Every block that can reach src/ repeats the rules protecting it.")
        return 0

    for failure in failures:
        print(f"::error::{failure}")
    print()
    print("A Dependabot ignore rule is scoped to the block that declares it - nothing "
          "is inherited. A block rooted outside src/ still reaches those projects "
          "(tests through ProjectReference, the root through the solution), so every "
          "rule protecting src/ has to be repeated in it. Add the missing "
          "dependency-name entries to that block's ignore list.")
    return 1


if __name__ == "__main__":
    sys.exit(main())
