#!/usr/bin/env python3
"""Regenerate the factual half of each THIRD-PARTY-NOTICES.txt.

The four constraints this package exists to satisfy are properties of the
RESOLVED dependency graph, and "all dependencies are permissively licensed" is
the one of them that nothing else checks. ci.yml's premise guards assert that no
BANNED package is present; they say nothing about what the packages that ARE
present are licensed under. That claim lived only in a hand-maintained text file
and in a hand-counted sentence in README.md, so it could go stale silently - and
had: both overstated the closure by one, because SixLabors.Fonts is listed as
both a direct and a transitive dependency (it is genuinely both) and got counted
twice, and the extensions notices omitted Microsoft.Extensions.Primitives
entirely.

Only the facts are generated - the package list, versions, licences and totals,
between BEGIN/END GENERATED markers. Everything else in those files is
hand-written rationale (why SixLabors.Fonts is pinned, what was deliberately
excluded and why) that a generator has no business owning.

Sources, both offline:
  - packages.lock.json  the resolved graph, committed, no restore needed to read
  - <cache>/<id>/<ver>/<id>.nuspec   the licence, written by the package author

Reading the licence from the nuspec rather than a table in this script is the
whole point: a table would be one more hand-maintained claim of exactly the kind
that just drifted.

Usage:
    python scripts/gen-third-party-notices.py            # rewrite in place
    python scripts/gen-third-party-notices.py --check    # fail if out of date
"""

from __future__ import annotations

import argparse
import difflib
import json
import os
import re
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent

CORE = REPO / "src" / "DocToolkit"
EXTENSIONS = REPO / "src" / "DocToolkit.Extensions.DependencyInjection"

BEGIN = "--- BEGIN GENERATED (scripts/gen-third-party-notices.py) - do not edit by hand ---"
END = "--- END GENERATED ---"


def nuget_cache() -> Path:
    """The global packages folder, honouring NUGET_PACKAGES as `dotnet` does."""
    env = os.environ.get("NUGET_PACKAGES")
    return Path(env) if env else Path.home() / ".nuget" / "packages"


def resolved_packages(project: Path) -> dict[str, set[str]]:
    """Map package id -> set of resolved versions, unioned across every TFM.

    A package can resolve to different versions per framework -
    System.IO.Packaging is 8.0.1 on net8.0 and 10.0.2 on net10.0 - and an
    air-gapped consumer mirroring this closure needs both, so versions are a set
    rather than a single value.
    """
    lock = json.loads((project / "packages.lock.json").read_text(encoding="utf-8-sig"))
    out: dict[str, set[str]] = {}
    for packages in lock["dependencies"].values():
        for name, info in packages.items():
            version = info.get("resolved")
            if version:
                out.setdefault(name, set()).add(version)
    return out


def direct_packages(project: Path) -> set[str]:
    lock = json.loads((project / "packages.lock.json").read_text(encoding="utf-8-sig"))
    return {
        name
        for packages in lock["dependencies"].values()
        for name, info in packages.items()
        if info.get("type") == "Direct"
    }


def read_nuspec(name: str, version: str) -> tuple[str, str]:
    """Return (licence, url) for one package, read from its cached .nuspec."""
    folder = nuget_cache() / name.lower() / version
    nuspec = folder / f"{name.lower()}.nuspec"
    if not nuspec.exists():
        candidates = sorted(folder.glob("*.nuspec"))
        if not candidates:
            sys.exit(
                f"error: no .nuspec for {name} {version} under {folder}.\n"
                f"       Run `dotnet restore` first - licences are read from the "
                f"restored package, not from a table in this script."
            )
        nuspec = candidates[0]

    text = nuspec.read_text(encoding="utf-8-sig", errors="replace")

    expression = re.search(r'<license\s+type="expression">([^<]+)</license>', text)
    if expression:
        licence = expression.group(1).strip()
    else:
        # No SPDX expression means the author used the deprecated <licenseUrl>, or
        # nothing at all. Either way a human has to look at it, so say so loudly
        # rather than emitting a blank and letting it read as "checked, fine".
        url = re.search(r"<licenseUrl>([^<]+)</licenseUrl>", text)
        licence = f"UNDECLARED (see {url.group(1)})" if url else "UNDECLARED"

    # Prefer the repository over projectUrl: it is where the licence text actually
    # lives, and projectUrl is often a docs site (AngleSharp) or a landing page
    # (System.IO.Packaging -> https://dot.net/).
    repo = re.search(r'<repository[^>]*\burl="([^"]+)"', text)
    project_url = re.search(r"<projectUrl>([^<]+)</projectUrl>", text)
    url = (repo.group(1) if repo else project_url.group(1) if project_url else "")
    url = re.sub(r"\.git$", "", url.strip())

    return licence, url


def render(names: dict[str, set[str]], direct: set[str], totals: bool) -> str:
    lines: list[str] = []
    pairs = 0
    licences: dict[str, int] = {}

    for name in sorted(names, key=str.lower):
        for version in sorted(names[name]):
            licence, url = read_nuspec(name, version)
            kind = "direct" if name in direct else "transitive"
            tail = f" - {url}" if url else ""
            lines.append(f"- {name} {version} - {licence} ({kind}){tail}")
            licences[licence] = licences.get(licence, 0) + 1
            pairs += 1

    if totals:
        breakdown = ", ".join(f"{n} {lic}" for lic, n in sorted(licences.items()))
        n_direct = len([n for n in names if n in direct])
        lines.append("")
        lines.append(
            f"Totals: {pairs} packages ({n_direct} direct, {pairs - n_direct} transitive) "
            f"- {breakdown}."
        )
        # Only worth explaining when it actually happened, otherwise the note reads
        # as a caveat about a package the reader cannot find in the list above.
        if any(len(v) > 1 for v in names.values()):
            lines.append(
                "A package resolving to two versions across the two target frameworks is "
                "counted once per version,"
            )
            lines.append(
                "because an air-gapped consumer has to mirror both onto their own feed."
            )

    return "\n".join(lines)


def check_readme(names: dict[str, set[str]]) -> list[str]:
    """Verify README.md's hand-written package counts against the resolved graph.

    README.md is prose, so it gets checked rather than generated - a marker block
    inside a markdown table would be worse than the problem. It states the same
    closure twice ("18 dependencies: 17 MIT, 1 Apache-2.0" in the constraints
    table, and again under Dependencies), and both were wrong for the same reason
    the notices file was. Numbers that appear in two files and are maintained in
    neither is exactly the shape of claim this script exists to catch.
    """
    pairs = [(n, v) for n, versions in names.items() for v in versions]
    counts: dict[str, int] = {}
    for name, version in pairs:
        licence, _ = read_nuspec(name, version)
        counts[licence] = counts.get(licence, 0) + 1

    expected = (len(pairs), counts.get("MIT", 0), counts.get("Apache-2.0", 0))

    readme = (REPO / "README.md").read_text(encoding="utf-8")
    pattern = re.compile(
        r"(\d+)\s+(?:dependencies|packages)\s*[:—–-]\s*"
        r"(\d+)\s+MIT,\s*(\d+)\s+Apache-2\.0"
    )
    found = pattern.findall(readme)

    if not found:
        return [
            "README.md: could not find a package-count claim to verify. If the "
            "wording changed, update the pattern in scripts/gen-third-party-notices.py "
            "rather than leaving the claim unchecked."
        ]

    problems = []
    for total, mit, apache in found:
        if (int(total), int(mit), int(apache)) != expected:
            problems.append(
                f"README.md says {total} packages ({mit} MIT, {apache} Apache-2.0); "
                f"the resolved graph has {expected[0]} ({expected[1]} MIT, "
                f"{expected[2]} Apache-2.0)."
            )
    return problems


def splice(path: Path, block: str) -> str:
    text = path.read_text(encoding="utf-8")
    if BEGIN not in text or END not in text:
        sys.exit(f"error: {path} is missing the BEGIN/END GENERATED markers.")
    head, rest = text.split(BEGIN, 1)
    _, tail = rest.split(END, 1)
    return f"{head}{BEGIN}\n{block}\n{END}{tail}"


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--check",
        action="store_true",
        help="do not write; exit 1 with a diff if any file is out of date",
    )
    args = parser.parse_args()

    core_names = resolved_packages(CORE)

    # The extensions package documents only what it adds. Everything it inherits
    # through Ank.DocToolkit is already listed in the core package's notices, and
    # duplicating seventeen entries across two files would just create a second
    # copy to drift. Subtracting by NAME is what caught Microsoft.Extensions.
    # Primitives: it reaches the extensions package via Microsoft.Extensions.
    # Options, not via Ank.DocToolkit, so the "see the core notices" pointer never
    # covered it and it was attributed nowhere.
    ext_all = resolved_packages(EXTENSIONS)
    ext_own = {n: v for n, v in ext_all.items() if n not in core_names}

    targets = [
        (CORE / "THIRD-PARTY-NOTICES.txt", render(core_names, direct_packages(CORE), True)),
        (
            EXTENSIONS / "THIRD-PARTY-NOTICES.txt",
            render(ext_own, direct_packages(EXTENSIONS), True),
        ),
    ]

    stale = False
    for path, block in targets:
        current = path.read_text(encoding="utf-8")
        updated = splice(path, block)
        if current == updated:
            print(f"up to date: {path.relative_to(REPO)}")
            continue
        stale = True
        if args.check:
            print(f"\nOUT OF DATE: {path.relative_to(REPO)}")
            sys.stdout.writelines(
                difflib.unified_diff(
                    current.splitlines(keepends=True),
                    updated.splitlines(keepends=True),
                    fromfile=f"{path.name} (committed)",
                    tofile=f"{path.name} (regenerated)",
                )
            )
        else:
            path.write_text(updated, encoding="utf-8", newline="\n")
            print(f"rewrote: {path.relative_to(REPO)}")

    if args.check and stale:
        print(
            "\n::error::THIRD-PARTY-NOTICES.txt no longer matches the resolved "
            "dependency graph. Run `python scripts/gen-third-party-notices.py` and "
            "commit the result."
        )

    # Always checked, in both modes: this one is never rewritten automatically, so
    # a silent pass in write mode would be the same trap it is meant to close.
    readme_problems = check_readme(core_names)
    for problem in readme_problems:
        print(f"::error::{problem}")
    if not readme_problems:
        print("up to date: README.md package counts")

    return 1 if (args.check and stale) or readme_problems else 0


if __name__ == "__main__":
    raise SystemExit(main())
