#!/usr/bin/env python3
"""Fail when a workflow uses an action that is not pinned by a full commit hash (D48).

    python scripts/check-action-pins.py
    python scripts/check-action-pins.py --self-test

WHY THIS EXISTS. A workflow runs other people's code with this repository's token, and the
release workflow holds the token that publishes to nuget.org. A tag such as `@v4` can be moved
to other code after it was reviewed; a 40-character commit hash cannot. Every workflow here was
already pinned that way, by hand, and nothing held it. This holds it:

  - every `uses:` names a full commit hash, except an action in this repository (`./...`) or a
    container image (`docker://...`);
  - the version the hash was read from sits beside it, as `# vX.Y.Z`, so a reader and
    Dependabot's pull requests can say what changed;
  - Dependabot watches the `github-actions` ecosystem, so a pinned hash is updated rather than
    quietly growing years old.

A commented-out `uses:` is not read. A scan that read no `uses:` line at all fails: it checked
nothing.
"""

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
WORKFLOWS = ROOT / ".github" / "workflows"
DEPENDABOT = ROOT / ".github" / "dependabot.yml"

USES = re.compile(r"^\s*(?:-\s*)?uses:\s*(\S+)(.*)$")
PINNED = re.compile(r"^[\w.-]+/[\w./-]+@[0-9a-f]{40}$")
VERSION_COMMENT = re.compile(r"^\s*#\s*v\d+(\.\d+)*\s*$")


def uses_lines(text):
    """(line number, what is used, the rest of the line) for every live `uses:` line."""
    found = []
    for number, line in enumerate(text.splitlines(), 1):
        match = USES.match(line)
        if match:
            found.append((number, match.group(1).strip("\"'"), match.group(2)))
    return found


def pin_problems(text):
    """What is wrong with how one workflow names its actions, one sentence each."""
    problems = []
    for number, used, rest in uses_lines(text):
        if used.startswith("./") or used.startswith("docker://"):
            continue
        if not PINNED.match(used):
            problems.append(f"line {number}: {used} is not pinned by a full commit hash")
        elif not VERSION_COMMENT.match(rest):
            problems.append(f"line {number}: {used.split('@')[0]} has no '# vX.Y.Z' comment beside its hash")
    return problems


def ecosystems(text):
    """The package ecosystems a dependabot.yml watches."""
    return [line.split(":", 1)[1].strip().strip("\"'") for line in text.splitlines()
            if line.strip().startswith("- package-ecosystem:")]


def self_test():
    bad = []

    def expect(name, ok):
        if not ok:
            bad.append(name)

    sample = "\n".join([
        "    steps:",
        "      - uses: actions/checkout@v4",
        "      - uses: actions/setup-python@main",
        "      - uses: some-org/some-action@11d5960",
        "      - uses: actions/checkout@11d5960a326750d5838078e36cf38b85af677262",
        "      - uses: actions/checkout@11d5960a326750d5838078e36cf38b85af677262 # v4.4.0",
        "      - uses: some-org/an-action/in-a-folder@33d7b85da1a04171e211befe2c739303b9a4c066 # v4.39.3",
        "      # - uses: commented/out@v1",
        "      - uses: ./.github/actions/local",
        "        uses: docker://alpine:3.20",
        "      - uses: \"actions/cache@v4\"",
    ])
    problems = pin_problems(sample)
    expect("a tag, a branch, a short hash and a quoted tag are reported, in order",
           [p.split(":")[0] for p in problems] == ["line 2", "line 3", "line 4", "line 5", "line 11"])
    expect("a hash with no version comment is reported as that",
           len(problems) > 3 and "no '# vX.Y.Z' comment" in problems[3])
    expect("a pinned action in a sub-folder passes", not any(p.startswith("line 7") for p in problems))
    expect("a commented-out line is not read", len(uses_lines(sample)) == 9)
    expect("dependabot ecosystems are read",
           ecosystems("updates:\n  - package-ecosystem: nuget\n  - package-ecosystem: \"github-actions\"\n")
           == ["nuget", "github-actions"])

    for name in bad:
        print(f"::error::self-test case '{name}' failed")
    if bad:
        print(f"::error::self-test failed {len(bad)} case(s)")
        return 1
    print("self-test passed")
    return 0


def main(argv):
    if "--self-test" in argv:
        return self_test()
    files = sorted(WORKFLOWS.glob("*.yml")) + sorted(WORKFLOWS.glob("*.yaml"))
    problems, read = [], 0
    for path in files:
        text = path.read_text(encoding="utf-8")
        read += len(uses_lines(text))
        problems += [f"{path.name} {problem}" for problem in pin_problems(text)]
    if read == 0:
        print("::error::no `uses:` line was read in .github/workflows: the scan checked nothing.")
        return 1
    if not DEPENDABOT.exists() or "github-actions" not in ecosystems(DEPENDABOT.read_text(encoding="utf-8")):
        problems.append("dependabot.yml does not watch the github-actions ecosystem: a pinned hash is never updated")
    for problem in problems:
        print(f"::error::{problem}")
    if problems:
        print("A tag can be moved to other code after it was reviewed; a commit hash cannot.")
        return 1
    print(f"{read} `uses:` line(s) in {len(files)} workflow(s): every action is pinned by commit hash.")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
