#!/usr/bin/env python3
"""Assert a built .nupkg carries every assembly the package is supposed to ship.

WHY THIS EXISTS. Version 0.36.0 was published to nuget.org carrying `DocToolkit.dll`
and nothing else - a 67 KB shell where pre-split 0.35.0 had shipped 266 KB of
working library. The six sibling assemblies the per-concern project split created
(Primitives, Docx, Html, Pdf, Pptx, Xlsx) were simply absent, so every consumer
touching `DocxEditor`, `PageSetup`, `PdfEditor`, `WorkbookEditor` or
`PresentationEditor` got CS0012/CS0103 and could not compile.

A nuget.org version can be unlisted but never deleted or replaced, so that
release is permanent.

THE CAUSE, and why every existing check missed it. `src/DocToolkit/DocToolkit.csproj`
packs the siblings through a `TargetsForTfmSpecificBuildOutput` target that reads
`@(ReferenceCopyLocalPaths)`. Only `ResolveReferences` populates that item, and
`release.yml` packed with the `no-build` switch, which skips it. So the item list
was empty and `BuildOutputInPackage` added nothing.

Nothing else packs that way. A plain `dotnet pack` runs `ResolveReferences` and
emits all seven assemblies, so every local pack and every CI pack was correct -
the release was the only wrong one, on the one run that cannot be undone. The
tests passed, the premise guards passed, the API approval passed, the install
smoke test passed: none of them look inside the produced package.

WHAT THIS CHECKS. The nupkg's `lib/<tfm>/` folders, against the assemblies the
project is expected to ship, for every target framework the package declares.
It reads the EXPECTED list from the csproj's ProjectReference elements rather
than from a list written down here - the same reason
`gen-third-party-notices.py` reads the lockfile and `automerge-eligible.py`
reads the workflows. A hand-maintained list is what goes stale, and staleness
here means shipping a broken package with a green check.

    python scripts/check-package-contents.py artifacts/Ank.DocToolkit.1.2.3.nupkg
    python scripts/check-package-contents.py --self-test

`--self-test` runs the controls, including the NEGATIVE one: a package holding
only the primary assembly must FAIL. A checker that passed everything would be
worse than no checker, which is the shape of the defect it exists to catch.
"""

import io
import pathlib
import re
import sys
import zipfile

REPO = pathlib.Path(__file__).resolve().parent.parent

# The packable projects, and the csproj whose ProjectReferences say what each must carry.
PACKAGES = {
    "Ank.DocToolkit": REPO / "src" / "DocToolkit" / "DocToolkit.csproj",
    "Ank.DocToolkit.Extensions.DependencyInjection":
        REPO / "src" / "DocToolkit.Extensions.DependencyInjection"
        / "DocToolkit.Extensions.DependencyInjection.csproj",
}

PROJECT_REFERENCE = re.compile(r'<ProjectReference\s+Include="([^"]+)"')
LIB_ENTRY = re.compile(r"^lib/(?P<tfm>[^/]+)/(?P<name>[^/]+)\.dll$")


def expected_assemblies(csproj):
    """The assembly names a package built from `csproj` must contain.

    Derived: the project's own assembly, plus one per ProjectReference. A new
    per-format project joins this check by existing, which is the point.
    """
    text = io.open(csproj, encoding="utf-8").read()
    names = {csproj.stem}
    for include in PROJECT_REFERENCE.findall(text):
        names.add(pathlib.PurePath(include.replace("\\", "/")).stem)
    return names


def inspect(nupkg):
    """{tfm: {assembly names}} for every lib/ folder in the package."""
    found = {}
    with zipfile.ZipFile(nupkg) as z:
        for entry in z.namelist():
            m = LIB_ENTRY.match(entry)
            if m:
                found.setdefault(m.group("tfm"), set()).add(m.group("name"))
    return found


def check(nupkg, expected, label):
    """Returns a list of failure strings; empty means the package is complete."""
    failures = []
    found = inspect(nupkg)

    if not found:
        # A package with no lib/ at all would otherwise satisfy every per-tfm loop below
        # by never entering it - passing vacuously, which is the failure this guards.
        return [f"{label}: no lib/<tfm>/*.dll entries at all. Refusing to call that complete."]

    for tfm in sorted(found):
        missing = sorted(expected - found[tfm])
        if missing:
            failures.append(
                f"{label}: lib/{tfm}/ is missing {len(missing)} assembl"
                f"{'y' if len(missing) == 1 else 'ies'}: {', '.join(missing)}. "
                f"It has {', '.join(sorted(found[tfm]))}."
            )

    print(f"  {label}: {len(found)} target framework(s), "
          f"{len(expected)} expected assembl{'y' if len(expected) == 1 else 'ies'} each")
    for tfm in sorted(found):
        mark = "OK " if not (expected - found[tfm]) else "BAD"
        print(f"    {mark} lib/{tfm}/ {len(found[tfm])} present")

    return failures


def self_test():
    """Controls, positive and negative, over synthetic packages."""
    import tempfile

    expected = {"DocToolkit", "DocToolkit.Primitives", "DocToolkit.Docx"}
    cases = [
        ("complete package passes",
         ["lib/net8.0/DocToolkit.dll", "lib/net8.0/DocToolkit.Primitives.dll",
          "lib/net8.0/DocToolkit.Docx.dll"], False),
        # THE 0.36.0 CASE. This must fail, and it is the whole reason the file exists.
        ("primary assembly alone FAILS",
         ["lib/net8.0/DocToolkit.dll"], True),
        ("one framework complete, the other short, FAILS",
         ["lib/net8.0/DocToolkit.dll", "lib/net8.0/DocToolkit.Primitives.dll",
          "lib/net8.0/DocToolkit.Docx.dll", "lib/net10.0/DocToolkit.dll"], True),
        ("no lib/ at all FAILS",
         ["DocToolkit.nuspec"], True),
    ]

    bad = 0
    with tempfile.TemporaryDirectory() as tmp:
        for name, entries, should_fail in cases:
            path = pathlib.Path(tmp) / "t.nupkg"
            with zipfile.ZipFile(path, "w") as z:
                for e in entries:
                    z.writestr(e, b"")
            failed = bool(check(path, expected, name))
            ok = failed == should_fail
            print(f"  {'ok  ' if ok else 'FAIL'} {name} -> {'failed' if failed else 'passed'}")
            if not ok:
                bad += 1

    # The id lookup, which the synthetic cases above never touch - and which had a real bug.
    for name, want in [
        ("Ank.DocToolkit.1.2.3.nupkg", "Ank.DocToolkit"),
        ("Ank.DocToolkit.Extensions.DependencyInjection.1.2.3.nupkg",
         "Ank.DocToolkit.Extensions.DependencyInjection"),
    ]:
        cands = [p for p in PACKAGES if name.startswith(p + ".")]
        got = max(cands, key=len) if cands else None
        if got == want:
            print(f"  ok   {name} -> {got}")
        else:
            print(f"  FAIL {name} -> {got}, expected {want}")
            bad += 1

    # And the derivation itself, against the real csproj.
    real = expected_assemblies(PACKAGES["Ank.DocToolkit"])
    if len(real) < 2:
        print(f"  FAIL expected_assemblies read {len(real)} name(s) from the real csproj; "
              "a derivation that finds nothing would pass every package")
        bad += 1
    else:
        print(f"  ok   expected_assemblies reads {len(real)} names from the real csproj: "
              f"{', '.join(sorted(real))}")

    print()
    if bad:
        print(f"::error::self-test failed {bad} case(s)")
        return 1
    print("self-test passed, negative controls included")
    return 0


def main(argv):
    if "--self-test" in argv:
        return self_test()

    paths = [pathlib.Path(a) for a in argv if not a.startswith("-")]
    if not paths:
        sys.exit("usage: check-package-contents.py <nupkg> [<nupkg> ...] | --self-test")

    failures = []
    checked = 0
    for nupkg in paths:
        if not nupkg.exists():
            sys.exit(f"::error::{nupkg} not found; this check cannot run.")

        # LONGEST match, not the first. "Ank.DocToolkit.Extensions.DependencyInjection.1.2.3.nupkg"
        # also starts with "Ank.DocToolkit.", so a first-match lookup checked the extensions package
        # against the CORE package's expected assemblies and reported seven missing. Caught by
        # running this against both real nupkgs; the self-test could not see it, because it never
        # exercises the id lookup.
        candidates = [p for p in PACKAGES if nupkg.name.startswith(p + ".")]
        package = max(candidates, key=len) if candidates else None
        if package is None:
            print(f"  {nupkg.name}: not a known package id, skipped")
            continue

        failures += check(nupkg, expected_assemblies(PACKAGES[package]), nupkg.name)
        checked += 1

    if checked == 0:
        sys.exit("::error::no known packages were checked. A run that inspects nothing "
                 "must not report success.")

    print()
    if not failures:
        print(f"every assembly each package should ship is present ({checked} package(s))")
        return 0

    for f in failures:
        print(f"::error::{f}")
    print()
    print("This is how 0.36.0 shipped a 67 KB shell. Do not publish. If the pack step regained "
          "the `no-build` switch, that is the cause - see release.yml's Pack step.")
    return 1


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
