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

A SECOND invariant is checked here too, added 2026-08-09 after the same failure
arrived from a different direction: a sample that pins a FIXED version needs an
ignore in the root block, because the root block sees DocToolkit.sln and a
solution reaches every project. Both invariants derive their expected set from
the repository rather than from a list kept in this file.
"""

import glob
import io
import pathlib
import re
import sys

CONFIG = ".github/dependabot.yml"

BLOCK = re.compile(r"^  - package-ecosystem:\s*(\S+)")
DIRECTORY = re.compile(r"^    directory:\s*[\"']?([^\"'\s]+)")
DIRECTORIES_ITEM = re.compile(r"^      - [\"']?([^\"'\s]+)")
KEY = re.compile(r"^    ([\w-]+):")
IGNORED_NAME = re.compile(r"^\s+- dependency-name:\s*[\"']?([^\"'\s]+)")
PACKAGE_REF = re.compile(
    r'<PackageReference\s+Include="([^"]+)"\s+Version="([^"]+)"')


PACKAGE_REF_NAME = re.compile(r'<PackageReference\s+Include="([^"]+)"')


def ignore_lists(text):
    """Yield (ecosystem, [directories], [ignored names IN ORDER]) per update block.

    The list, not the set `blocks()` builds - a duplicate is invisible once deduplicated, and
    Dependabot refuses to parse a file containing one.
    """
    current = None
    section = None

    for line in text.splitlines():
        start = BLOCK.match(line)
        if start:
            if current:
                yield current
            current = (start.group(1), [], [])
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
                current[2].append(name.group(1))

    if current:
        yield current


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


def pinned_names(text):
    """The dependency names the src/ blocks deliberately hold back, wildcard excluded."""
    nuget = [b for b in blocks(text) if b[0] == "nuget"]
    src = [b for b in nuget if b[1] and all(d.startswith("/src/") for d in b[1])]
    names = set().union(*(b[2] for b in src)) if src else set()
    names.discard("*")
    return names


def shipped_packages():
    """Every package the packable projects declare, so no other block may bump one.

    Read from src/Directory.Build.props, which is where the split put all fifteen: a
    PrivateAssets="all" ProjectReference suppresses its own PackageReferences from the nuspec,
    so every package the library needs has to be declared there rather than in the per-format
    project that uses it. That makes the props file the complete shipped set.
    """
    props = pathlib.Path("src") / "Directory.Build.props"
    try:
        text = io.open(props, encoding="utf-8").read()
    except FileNotFoundError:
        sys.exit(f"::error::{props} not found, so the shipped package set cannot be derived. "
                 "This check would then require less than it should, which is worse than "
                 "failing.")

    names = {m for m in PACKAGE_REF_NAME.findall(text)}
    if not names:
        sys.exit(f"::error::No PackageReference parsed from {props}. A derivation that finds "
                 "nothing would make this check pass vacuously.")
    return names


def main():
    try:
        text = io.open(CONFIG, encoding="utf-8").read()
    except FileNotFoundError:
        sys.exit(f"::error::{CONFIG} not found; this check cannot run.")

    # dependency-report.yml asks for this rather than keeping its own copy. A second
    # list of pinned names would rot exactly the way the repeated ignore rules do -
    # which is the failure this whole script exists to prevent.
    if "--list-pinned" in sys.argv:
        for name in sorted(pinned_names(text)):
            print(name)
        return 0

    # ── Dependabot REJECTS a duplicate dependency-name, and this check could not see one ──
    #
    # Added 2026-08-24, immediately after causing the failure it now catches. C37 appended the
    # fifteen shipped packages to two ignore lists that already named SixLabors.Fonts, so both
    # gained a duplicate. Dependabot answered
    #
    #     The property '#/updates/2/ignore/18/dependency-name' is a duplicate.
    #
    # and stopped parsing the file AT ALL - which disables every update it would raise, security
    # updates included. It reached `main` because nothing here could see it: `blocks()` reduces
    # each ignore list to a SET, and a YAML parser accepts duplicate list items happily. The only
    # thing that objected was Dependabot's own API check, on the next pull request.
    #
    # Counted from the TEXT, with this file's own line patterns, deliberately. The first version
    # of this check used PyYAML inside a try/except ImportError - and no workflow here installs
    # PyYAML, so in CI it would have skipped itself and reported success. A check that silently
    # does nothing is worse than no check, which is the lesson this script already carries twice.
    for ecosystem, directories, names in ignore_lists(text):
        duplicated = sorted({n for n in names if names.count(n) > 1})
        if duplicated:
            scope = ", ".join(directories) or "(unscoped)"
            sys.exit(f"::error::the {ecosystem} block at {scope} ignores "
                     f"{', '.join(duplicated)} more than once. Dependabot rejects a duplicate "
                     "dependency-name and stops parsing the WHOLE file, which silently disables "
                     "every update including security ones. Merge the rules into one entry.")

    nuget = [b for b in blocks(text) if b[0] == "nuget"]
    if not nuget:
        sys.exit("::error::No nuget update blocks parsed from "
                 f"{CONFIG}. Either the file changed shape or this check stopped "
                 "matching it; both are worth a red build.")

    # The source of truth: whatever the src/ blocks protect, PLUS every package those
    # projects actually declare.
    #
    # The second half was added 2026-08-24, after #366 - titled "Bump the test-dependencies
    # group" - bumped SEVEN shipped runtime dependencies in src/Directory.Build.props:
    # OfficeIMO 3.2.2 -> 3.2.6 across eight packages, AngleSharp, and PdfPig. It carried
    # `prefix: chore`, which is hidden and non-bumping, so a consumer would have restored four
    # patch versions of the engine behind every DOCX, XLSX, PPTX and PDF path with no changelog
    # entry and no version proposed.
    #
    # Both src/ blocks ignore "*" and propose nothing, and their comment says bumps there "stay
    # a deliberate manual act". They did not: the tests block reaches src/ through
    # ProjectReference, Dependabot edits the src csproj on its way past, and that block is the
    # ONLY automated path that touches a shipped dependency.
    #
    # This check already passed, correctly, and its own comment below named the hole - it
    # required the four NAMED rules where src/Directory.Build.props declares fifteen packages.
    # Only SixLabors.Fonts was in both sets, leaving fourteen reachable and bumpable.
    #
    # DERIVED from the props file, never listed: a sixteenth package joins the requirement by
    # being added, which is the point.
    protects_src = [b for b in nuget if all(d.startswith("/src/") for d in b[1])]
    required = set().union(*(b[2] for b in protects_src)) if protects_src else set()
    required |= shipped_packages()

    # A wildcard is not a rule that protects src/ from other blocks - it says "propose
    # nothing for THIS project", which is a statement about that block's own update
    # policy. The src/ blocks carry one because Dependabot cannot write a correct
    # lockfile for a multi-targeted project (see the comment there), and requiring the
    # tests block to repeat it would stop test dependencies updating at all, for a
    # reason that has nothing to do with them.
    #
    # The NAMED rules still have to be repeated. A wildcard in src/ stops Dependabot
    # proposing bumps rooted there; it does nothing about a bump proposed from the tests
    # or root block that edits a src csproj on its way past.
    required.discard("*")

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

    # ---------------------------------------------------------------------------
    # Second invariant: a sample that pins a FIXED version needs an ignore in the
    # root block.
    #
    # Samples have no block of their own, because they reference the published
    # packages as Version="*" and nothing needs to propose those. That was recorded
    # as "samples are deliberately absent", and on 2026-08-09 it turned out to be
    # too strong: the root block is rooted at "/", DocToolkit.sln is in its view,
    # and a solution reaches every project. It proposed
    # Microsoft.Extensions.Hosting 8.0.1 -> 10.0.10 for samples/WorkerService,
    # whose pin exists precisely to stop an unrelated upstream release reddening a
    # sample.
    #
    # Derived from the csproj files rather than listed, for the same reason as
    # above: the next sample to pin something must not depend on somebody
    # remembering this paragraph.
    # ---------------------------------------------------------------------------
    pinned_in_samples = {}
    for csproj in sorted(glob.glob("samples/*/*.csproj")):
        for name, version in PACKAGE_REF.findall(io.open(csproj, encoding="utf-8").read()):
            if version != "*":
                pinned_in_samples.setdefault(name, csproj)

    root = next((b for b in nuget if b[1] == ["/"]), None)
    if pinned_in_samples and root is None:
        failures.append("samples pin fixed versions but there is no root-rooted nuget "
                        "block to carry the ignores")
    elif pinned_in_samples:
        print()
        print(f"{len(pinned_in_samples)} fixed-version pin(s) in samples/, reachable from "
              "the root block through the solution:")
        for name, csproj in sorted(pinned_in_samples.items()):
            ok = name in root[2]
            print(f"  {name} ({csproj}) - {'OK' if ok else 'NOT IGNORED'}")
            if not ok:
                failures.append(
                    f"{csproj} pins `{name}` at a fixed version, but the root block does "
                    "not ignore it - the solution puts that project in its view")

    print()
    if not failures:
        print("Every block that can reach src/ repeats the rules protecting it, and every "
              "fixed-version sample pin is protected.")
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
