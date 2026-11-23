#!/usr/bin/env python3
"""Fail when a workflow uses an action that is not pinned by a full commit hash (D48), or lacks
one `on:` line and one `jobs:` line (D54).

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

IT ALSO HOLDS THE TWO KEYS GITHUB READS A WORKFLOW BY (D54): `on:` and `jobs:` each start exactly
one line. A file that lost its `on:` line (a template once cut a comment at the letters `on:`,
which also end `reason:`) is no workflow at all to GitHub: a required check that never starts
blocks every merge, and one that is not required stops without a word. It lives here because
this is the guard that already reads every workflow file.
"""

import re
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
WORKFLOWS = ROOT / ".github" / "workflows"
DEPENDABOT = ROOT / ".github" / "dependabot.yml"

USES = re.compile(r"^\s*(?:-\s*)?uses:\s*(\S+)(.*)$")
PINNED = re.compile(r"^[\w.-]+/[\w./-]+@[0-9a-f]{40}$")
VERSION_COMMENT = re.compile(r"^\s*#\s*v\d+(\.\d+)*\s*$")

# The two keys GitHub reads a workflow by (D54). Each must start exactly one line of the file.
WORKFLOW_KEYS = ("on", "jobs")


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


def key_problems(name, text):
    """What GitHub would refuse a workflow for (D54): `on:` and `jobs:` each start a line, once.
    A value on the same line and a key in quotes count; a key inside another key (indented), in
    a comment, or a longer word that begins with the key does not."""
    problems = []
    for key in WORKFLOW_KEYS:
        pattern = re.compile(rf"""^(?:{key}|"{key}"|'{key}')[ \t]*:""")
        found = [line for line in text.splitlines() if pattern.match(line)]
        if len(found) != 1:
            problems.append(f"{name} has {len(found)} lines that start with `{key}:`, not one: GitHub would refuse it")
    return problems


def scan(files):
    """(every problem in the workflow files, how many `uses:` lines were read)."""
    problems, read = [], 0
    for path in files:
        text = path.read_text(encoding="utf-8")
        read += len(uses_lines(text))
        problems += [f"{path.name} {problem}" for problem in pin_problems(text)]
        problems += key_problems(path.name, text)
    return problems, read


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

    # D54: the two keys GitHub reads a workflow by. Without `on:` it is no workflow at all, and a
    # required check that never starts blocks every merge without a word.
    sound = ("name: sample\non:\n  push:\n    branches: [main]\n  workflow_dispatch:\n"
             "jobs:\n  gate:\n    runs-on: ubuntu-latest\n")
    expect("a workflow with both keys passes", key_problems("s.yml", sound) == [])
    for written in ("on: push\n", "on: [push, pull_request]\n", '"on":\n  push:\n', "'on':\n  push:\n",
                    "on:   # what starts it\n  push:\n"):
        expect(f"the key written as {written.splitlines()[0]!r} counts",
               key_problems("s.yml", f"name: sample\n{written}jobs:\n  gate: {{}}\n") == [])
    expect("a quoted jobs key counts",
           key_problems("s.yml", sound.replace("\njobs:\n", '\n"jobs":\n')) == [])
    no_on = sound.replace("\non:\n", "\n")
    expect("a workflow with no `on:` line is reported, by name and count",
           key_problems("s.yml", no_on) == ["s.yml has 0 lines that start with `on:`, not one: GitHub would refuse it"])
    expect("a workflow with no `jobs:` line is reported",
           key_problems("s.yml", sound.replace("jobs:\n", "")) == [
               "s.yml has 0 lines that start with `jobs:`, not one: GitHub would refuse it"])
    expect("a key that starts two lines is reported",
           key_problems("s.yml", sound + "on:\n  workflow_dispatch:\n") == [
               "s.yml has 2 lines that start with `on:`, not one: GitHub would refuse it"])
    expect("a comment that holds the key does not count",
           len(key_problems("s.yml", no_on.replace("name: sample\n", "name: sample\n# on:\n# left out, for one reason:\n"))) == 1)
    expect("a key inside another key does not count",
           len(key_problems("s.yml", no_on.replace("name: sample\n", "name: sample\nenv:\n  on: yes\n"))) == 1)
    expect("a word that only begins with `on` does not stand in for the key",
           key_problems("s.yml", no_on.replace("name: sample\n", "name: sample\nonly: x\n")) == [
               "s.yml has 0 lines that start with `on:`, not one: GitHub would refuse it"])
    expect("a word that only begins with `jobs` is not a second `jobs:` line",
           key_problems("s.yml", sound.replace("name: sample\n", "name: sample\njobsite: y\n")) == [])
    # The scan of real files runs both rules: a rule the scan never calls holds nothing.
    with tempfile.TemporaryDirectory() as tmp:
        lost = Path(tmp) / "lost.yml"
        lost.write_text(no_on.replace("runs-on: ubuntu-latest\n", "runs-on: ubuntu-latest\n    steps:\n"
                                      "      - uses: actions/checkout@v4\n"), encoding="utf-8")
        found, read = scan([lost])
        expect("the scan of a file reports both an unpinned action and a missing `on:` line",
               read == 1 and len(found) == 2 and found[0].startswith("lost.yml line ")
               and found[1].startswith("lost.yml has 0 lines that start with `on:`"))

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
    problems, read = scan(files)
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
    print(f"{read} `uses:` line(s) in {len(files)} workflow(s): every action is pinned by commit hash, "
          f"and each workflow has one `on:` line and one `jobs:` line.")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
