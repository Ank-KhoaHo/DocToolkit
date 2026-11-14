#!/usr/bin/env python3
"""Run every check this repository can run on one machine, in order, each with a time limit.

    python scripts/check.py             every step: guards, format, build, the Linux test suite
    python scripts/check.py --quick     the guards and the format check only, and it says so
    python scripts/check.py --list      print the steps and run nothing
    python scripts/check.py --self-test the machinery, and that no CI guard is missing here

WHY THIS EXISTS. CI splits this repository's checks across jobs and workflows, and nothing ran
all of them locally with one command. The guards here are the ones CI runs; the self-test (run in
CI) fails when a workflow runs a script this gate neither runs nor names, with a reason, in
NOT_LOCAL - so the two lists cannot drift apart unnoticed.

Passing here is not passing in CI. CI builds on four operating systems, and two of its checks
read output only a CI run produces (coverage, the packed .nupkg). This gate says what it ran.

THE TESTS RUN IN THE LINUX CONTAINER (Dockerfile.linux-test), which needs Docker and nothing
else. That is deliberate: it is CI's own "linux container build & test" image, and on a machine
where application control blocks freshly built assemblies it is the only way the suite runs.

IT IS PYTHON, NOT A SHELL SCRIPT, ON PURPOSE. On Windows the word `bash` starts a different shell
depending on where it is typed, and the shell's `timeout` there is another program entirely.
`subprocess` gives one behaviour everywhere, and a time limit that is the same everywhere.
"""

import os
import re
import subprocess
import sys
import tempfile
import threading
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
WORKFLOWS = ROOT / ".github" / "workflows"

# Set in the environment of every step, so a step that started the gate again is refused.
RUNNING = "DOCTOOLKIT_GATE_RUNNING"

PY = sys.executable          # the interpreter running the gate, not whichever is first on PATH


class Step:
    """One check: a name a person recognises, a command (a list, never a shell string), a time
    limit in seconds, and whether --quick runs it."""

    def __init__(self, name, command, timeout, quick=True):
        self.name, self.command, self.timeout, self.quick = name, command, timeout, quick


def _guard(script, *args, timeout=180):
    return Step(" ".join([script, *args]), [PY, f"scripts/{script}", *args], timeout)


STEPS = [
    # The gate's own self-test first. It once passed here and failed in CI: it had last run
    # before ci.yml began to run it, so it never saw itself in the workflows.
    _guard("check.py", "--self-test"),
    # Then the guards' self-tests: a guard whose own test fails says nothing about the tree.
    _guard("backup-working-half.py", "--self-test"),
    _guard("check-action-pins.py", "--self-test"),
    _guard("check-changelog-curated.py", "--self-test"),
    _guard("check-doc-blocks.py", "--self-test"),
    _guard("check-docs-layout.py", "--self-test"),
    _guard("check-landing-calls.py", "--self-test"),
    _guard("check-mutation-scope.py", "--self-test"),
    _guard("check-no-private-references.py", "--self-test"),
    _guard("check-package-contents.py", "--self-test"),
    _guard("check-packed-install.py", "--self-test"),
    _guard("check-readme-conversions.py", "--self-test"),
    _guard("check-readme-shape.py", "--self-test"),
    _guard("check-render-policy.py", "--self-test"),
    _guard("filter-codeql-sarif.py", "--self-test"),
    _guard("gen-type-index.py", "--self-test"),
    _guard("set-floors.py", "--self-test"),
    # The guards, on this tree.
    _guard("check-action-pins.py"),
    _guard("check-configureawait.py"),
    _guard("check-core-sharing.py"),
    _guard("check-dependabot-scoping.py"),
    _guard("check-doc-blocks.py"),
    _guard("check-doc-snippets.py"),
    _guard("check-docs-layout.py"),
    _guard("check-landing-calls.py"),
    _guard("check-mutation-scope.py"),
    _guard("check-no-private-references.py"),
    _guard("check-package-description.py"),
    _guard("check-readme-conversions.py"),
    _guard("check-readme-coverage.py"),
    _guard("check-readme-shape.py"),
    _guard("check-render-policy.py"),
    _guard("check-workflow-tools.py"),
    _guard("gen-capability-matrix.py", "--check"),
    _guard("gen-guard-inventory.py", "--check"),
    _guard("gen-readme-snippets.py", "--check"),
    _guard("gen-repository-layout.py", "--check"),
    _guard("gen-third-party-notices.py", "--check"),
    _guard("gen-type-index.py", "--check"),
    _guard("repair-lockfiles.py", "--check"),
    Step("dotnet format", ["dotnet", "format", "DocToolkit.sln", "--verify-no-changes"], 900),
    # --no-incremental: an unchanged project re-emits no warnings, so -warnaserror alone can
    # pass on a tree that has them.
    Step("build (Release, warnings are errors)",
         ["dotnet", "build", "DocToolkit.sln", "-c", "Release", "-warnaserror", "--no-incremental"],
         1200, quick=False),
    Step("tests: build the Linux image",
         ["docker", "build", "-f", "Dockerfile.linux-test", "-t", "doctoolkit-linux-test", "."],
         1800, quick=False),
    Step("tests: run the suite in it", ["docker", "run", "--rm", "doctoolkit-linux-test"],
         1800, quick=False),
]

# Scripts a workflow runs that this gate does not, each with the reason. An entry is a hole in
# the gate, so it needs one; the self-test fails on a workflow script that is in neither list.
NOT_LOCAL = {
    "automerge-eligible.py": "decides on a Dependabot PR's own metadata, given by the workflow",
    "check-changelog-curated.py": "its real run reads a release's version; the self-test runs here",
    "check-coverage.py": "reads the cobertura reports of CI's test matrix; the container run makes none",
    "check-guide-output.py": "its byte counts are anchored to the Linux runner's fonts",
    "check-package-contents.py": "reads the packed .nupkg, which only release and CI build",
    "check-packed-install.py": "installs the packed .nupkg, which only release and CI build",
    "extract-changelog-section.py": "a release step: it prints one version's notes",
    "filter-codeql-sarif.py": "its real run reads CodeQL's SARIF; the self-test runs here",
}


def workflow_scripts(text):
    """Every scripts/<name>.py a workflow file runs, by file name."""
    return set(re.findall(r"python3?\s+scripts/([A-Za-z0-9_.-]+\.py)", text))


def missing_from_gate(scripts, steps, not_local):
    """Scripts that no step runs and NOT_LOCAL does not name."""
    run = {Path(s.command[1]).name for s in steps if len(s.command) > 1 and s.command[1].startswith("scripts/")}
    return sorted(name for name in scripts if name not in run and name not in not_local)


def _show(line):
    """Write one line, whatever the console can display: a character it cannot encode costs one
    '?' on screen, never the reader thread (which would block the step on a full pipe)."""
    try:
        sys.stdout.write(line)
    except UnicodeEncodeError:
        encoding = getattr(sys.stdout, "encoding", None) or "ascii"
        sys.stdout.write(line.encode(encoding, errors="replace").decode(encoding))
    sys.stdout.flush()


def run_step(step, cwd=ROOT, show=True):
    """Run one step; show its whole output as it arrives. Returns (ok, reason, output).

    Never the last lines only: the line that explains a failure is often printed first."""
    cwd = Path(cwd)
    env = {**os.environ, "PYTHONIOENCODING": "utf-8", "PYTHONDONTWRITEBYTECODE": "1", RUNNING: "1"}
    try:
        proc = subprocess.Popen(step.command, cwd=cwd, env=env, stdout=subprocess.PIPE,
                                stderr=subprocess.STDOUT, text=True, encoding="utf-8",
                                errors="replace", bufsize=1)
    except OSError as exc:
        return False, f"could not start: {exc}", ""

    lines = []

    def read():
        assert proc.stdout is not None
        with proc.stdout:
            for line in proc.stdout:
                lines.append(line)
                if show:
                    _show(line)

    reader = threading.Thread(target=read, daemon=True)
    reader.start()
    timed_out = False
    try:
        proc.wait(timeout=step.timeout)
    except subprocess.TimeoutExpired:
        timed_out = True
        proc.kill()             # this step's own process; a step must stop what it starts
        proc.wait()
    reader.join(timeout=5.0)
    output = "".join(lines)
    if timed_out:
        return False, f"timed out after {step.timeout} s", output
    if proc.returncode != 0:
        return False, f"exit code {proc.returncode}", output
    return True, "", output


def run(steps, show=True, cwd=ROOT):
    """Run the steps in order, stop at the first failure. 0 only when every step passed; a gate
    with no steps returns 1, because it checked nothing."""
    if not steps:
        print("check: no steps to run, so nothing was checked.", flush=True)
        return 1
    for position, step in enumerate(steps, start=1):
        print(f"\n== {position}/{len(steps)}  {step.name}", flush=True)
        ok, reason, _ = run_step(step, cwd=cwd, show=show)
        if not ok:
            print(f"\nFAILED at step {position}, {step.name}: {reason}", flush=True)
            return 1
    print(f"\nAll {len(steps)} local checks passed.", flush=True)
    return 0


# ---------------------------------------------------------------------------
# Self-test: the machinery with tiny commands, then the drift check against the workflows.

def self_test():
    bad = []

    def expect(name, ok):
        if not ok:
            bad.append(name)

    def tiny(code, timeout=30):
        return Step("tiny", [PY, "-c", code], timeout)

    with tempfile.TemporaryDirectory() as tmp:
        ok, _, out = run_step(tiny("print('one'); print('two')"), cwd=tmp, show=False)
        expect("a passing step passes and keeps all its output", ok and out.splitlines() == ["one", "two"])

        ok, reason, out = run_step(tiny("print('why it failed'); import sys; sys.exit(3)"), cwd=tmp, show=False)
        expect("a failing step fails, says its exit code, and keeps the first line",
               not ok and "exit code 3" in reason and "why it failed" in out)

        ok, reason, _ = run_step(tiny("import time; time.sleep(30)", timeout=1), cwd=tmp, show=False)
        expect("a step past its time limit fails and says so", not ok and "timed out" in reason)

        ok, reason, _ = run_step(Step("missing", ["no-such-program-d48"], 10), cwd=tmp, show=False)
        expect("a command that cannot start fails", not ok and "could not start" in reason)

        ok, _, out = run_step(tiny(f"import os; print(os.environ.get('{RUNNING}'))"), cwd=tmp, show=False)
        expect("every step sees the marker that stops a nested gate", out.strip() == "1")

        expect("an empty gate fails, because it checked nothing", run([], show=False, cwd=tmp) == 1)
        expect("the gate stops at the first failure",
               run([tiny("import sys; sys.exit(1)"), tiny("open('ran', 'w')")], show=False, cwd=tmp) == 1
               and not (Path(tmp) / "ran").exists())

    # The drift check, first on a sample, then on this repository's workflows.
    sample = "run: |\n  python3 scripts/a.py --check\n  python scripts/b.py\n  python3 scripts/c.py x\n"
    expect("workflow scripts are read by name", workflow_scripts(sample) == {"a.py", "b.py", "c.py"})
    expect("a script in neither list is reported",
           missing_from_gate({"a.py", "b.py", "c.py"}, [_guard("a.py")], {"b.py": "why"}) == ["c.py"])

    found = set()
    for workflow in sorted(WORKFLOWS.glob("*.yml")):
        found |= workflow_scripts(workflow.read_text(encoding="utf-8"))
    expect(f"the workflows were read ({len(found)} scripts found)", len(found) >= 20)
    missing = missing_from_gate(found, STEPS, NOT_LOCAL)
    expect(f"every script CI runs is a step here or named in NOT_LOCAL (missing: {missing})", not missing)
    stale = sorted(name for name in NOT_LOCAL if name not in found)
    expect(f"NOT_LOCAL names only scripts a workflow runs (stale: {stale})", not stale)
    absent = sorted({Path(s.command[1]).name for s in STEPS if len(s.command) > 1
                     and s.command[1].startswith("scripts/") and not (ROOT / s.command[1]).exists()})
    expect(f"every step's script exists (absent: {absent})", not absent)

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
    steps = [s for s in STEPS if s.quick] if "--quick" in argv else STEPS
    if "--list" in argv:
        for position, step in enumerate(steps, start=1):
            print(f"{position}. {step.name}  ({step.timeout} s)")
        return 0
    if os.environ.get(RUNNING):
        print("check: the gate is already running in a parent process and will not start inside "
              "itself. A step must never run scripts/check.py.", flush=True)
        return 1
    code = run(steps)
    if "--quick" in argv:
        print("--quick: the build and the test suite did NOT run.", flush=True)
    return code


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
