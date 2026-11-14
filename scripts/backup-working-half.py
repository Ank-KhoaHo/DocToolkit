#!/usr/bin/env python3
"""Zip the maintainer's working material into a folder outside the repository.

    python scripts/backup-working-half.py               into a folder named _backups beside the repo
    python scripts/backup-working-half.py <folder>      into that folder
    python scripts/backup-working-half.py --self-test

WHY THIS EXISTS. This repository is public, so its working material - the backlog, the ticket
specs, the maintainer's guidance files - is gitignored on purpose. It is in no commit and no
push. It exists on one disk, and one `git clean -fdx` removes all of it, with no undo. This
gives it a second copy.

WHEN TO RUN IT: after a batch of tickets lands, and before any `git clean`.

WHAT IT DOES NOT DO. An archive on the same disk protects against a bad command, not against a
dead disk: copy it somewhere else as well. The archive holds every working note and the private
words of .leak-patterns.local, so keep it where those may be kept, never in a published folder.

CI runs only --self-test: the working material never reaches a runner. The self-test builds a
throwaway repository, and also checks that every entry below is gitignored HERE, so a typo in
the list cannot quietly leave a file out of every backup.
"""

import os
import subprocess
import sys
import tempfile
import zipfile
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]

# Every gitignored path that holds working material. A path ending in "/" is a folder, taken
# whole. .gitignore is the source: an entry here that git does not ignore fails the self-test.
WORKING_MATERIAL = (
    "docs/",
    "spike/",
    "CLAUDE.md",
    "AGENTS.md",
    "DEVELOPMENT-CHECKLIST.md",
    "BACKLOG.html",
    ".superpowers/",
    ".leak-patterns.local",
)


def _git(root, *args):
    return subprocess.run(["git", *args], cwd=root, capture_output=True, text=True,
                          encoding="utf-8", errors="replace")


def _tracked(root):
    """Every path git tracks. A git failure stops the backup: read as "nothing is tracked", it
    would quietly change which files are saved."""
    done = _git(root, "ls-files", "-z")
    if done.returncode != 0:
        raise SystemExit(f"backup: git could not list the tracked files in {root}: {done.stderr.strip()}")
    return {path for path in (done.stdout or "").split("\0") if path}


def working_files(root):
    """Every working file on disk, as repository-relative paths with '/', sorted.

    A file git tracks is left out: a tracked working file is a leak for the guards to report,
    not something a backup should paper over."""
    root = Path(root)
    tracked = _tracked(root)
    found = set()
    for entry in WORKING_MATERIAL:
        target = root / entry.rstrip("/")
        candidates = target.rglob("*") if entry.endswith("/") and target.is_dir() else [target]
        for path in candidates:
            if path.is_file():
                relative = path.relative_to(root).as_posix()
                if relative not in tracked:
                    found.add(relative)
    return sorted(found)


def backup(root, destination, now=None):
    """Write one archive into `destination` (created if needed) and return its path.

    `now` fixes the timestamp in the name; the self-test passes one so two backups share it."""
    root, destination = Path(root), Path(destination)
    if destination.resolve() == root.resolve() or destination.resolve().is_relative_to(root.resolve()):
        # Inside the repository it would be removed by the same `git clean` it is meant to
        # survive, and it would be one `git add -f` away from being published.
        raise SystemExit(f"backup: {destination} is inside the repository. Choose a folder outside it.")

    files = working_files(root)
    if not files:
        # An empty archive would look like a backup and hold nothing.
        raise SystemExit("backup: there is no working material on this disk to save.")

    destination.mkdir(parents=True, exist_ok=True)
    stem = f"{root.resolve().name}-working-half-{(now or datetime.now()):%Y-%m-%d-%H%M%S}"
    archive = destination / f"{stem}.zip"
    counter = 2
    while archive.exists():             # two backups in one second must both survive
        archive = destination / f"{stem}-{counter}.zip"
        counter += 1

    with zipfile.ZipFile(archive, "w", compression=zipfile.ZIP_DEFLATED) as held:
        for relative in files:
            held.write(root / relative, arcname=relative)
    return archive


# ---------------------------------------------------------------------------
# Self-test

# The sample spec path is joined from parts: written whole, check-docs-layout.py reads it as a
# reference to a real file under docs/ and reports it missing.
_WORKING = {
    "/".join(("docs", "superpowers", "specs", "D900-sample.md")): "# D900\n",
    "spike/notes.txt": "a spike\n",
    "CLAUDE.md": "guidance\n",
    "AGENTS.md": "agent rules\n",
    "BACKLOG.html": "<p>rows</p>\n",
    ".superpowers/sdd/ledger.md": "ledger\n",
    ".leak-patterns.local": "private-word\n",
}
_NOT_WORKING = {"README.md": "# Widget\n", "src/Widget.cs": "class Widget {}\n", "notes-elsewhere.txt": "x\n"}


def _sample_repo(base):
    root = Path(base) / "widget-forge"
    for name, text in {**_WORKING, **_NOT_WORKING}.items():
        path = root / name
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(text.encode("utf-8"))    # bytes: write_text adds CR on Windows
    _git(root, "init", "-q")
    _git(root, "add", "README.md", "src/Widget.cs")
    return root


def self_test():
    bad = []

    def expect(name, ok):
        if not ok:
            bad.append(name)

    def refused(call):
        try:
            call()
        except SystemExit:
            return True
        return False

    with tempfile.TemporaryDirectory() as tmp:
        root = _sample_repo(tmp)
        expect("only the working material is listed", working_files(root) == sorted(_WORKING))

        destination = Path(tmp) / "backups"
        same_second = datetime(2026, 11, 14, 12, 0, 0)
        first = backup(root, destination, same_second)
        expect("the archive is a zip in the destination, named after the repo",
               first.parent == destination and first.suffix == ".zip"
               and first.name.startswith("widget-forge-working-half-"))
        with zipfile.ZipFile(first) as held:
            expect("the archive holds exactly the working files", sorted(held.namelist()) == sorted(_WORKING))
            expect("every file arrives byte for byte",
                   all(held.read(name) == text.encode("utf-8") for name, text in _WORKING.items()))

        second = backup(root, destination, same_second)
        expect("a second backup does not replace the first", first != second and first.exists() and second.exists())

        expect("a destination inside the repository is refused", refused(lambda: backup(root, root / "backups")))
        expect("nothing is written inside the repository", not (root / "backups").exists())

        _git(root, "add", "-f", "CLAUDE.md")
        expect("a working file that git tracks is left out", "CLAUDE.md" not in working_files(root))

    with tempfile.TemporaryDirectory() as tmp:
        empty = Path(tmp) / "empty"
        (empty / "src").mkdir(parents=True)
        (empty / "src" / "A.cs").write_text("class A {}\n", encoding="utf-8")
        _git(empty, "init", "-q")
        destination = Path(tmp) / "backups"
        expect("nothing to back up is an error, not an empty archive", refused(lambda: backup(empty, destination)))
        expect("no archive is left behind", not destination.exists() or not any(destination.iterdir()))

        not_a_repo = Path(tmp) / "plain"
        not_a_repo.mkdir()
        (not_a_repo / "CLAUDE.md").write_text("x\n", encoding="utf-8")
        expect("a folder git cannot read is refused", refused(lambda: working_files(not_a_repo)))

    # Against this repository's own .gitignore: every entry must really be working material.
    for entry in WORKING_MATERIAL:
        probe = entry + "probe.txt" if entry.endswith("/") else entry
        ignored = _git(ROOT, "check-ignore", "-q", "--no-index", probe).returncode == 0
        expect(f"'{entry}' is gitignored in this repository", ignored)

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
    destination = Path(argv[0]) if argv else ROOT.parent / "_backups"
    archive = backup(ROOT, destination)
    with zipfile.ZipFile(archive) as held:
        count = len(held.namelist())
    print(f"Saved {count} working files to {archive} ({os.path.getsize(archive):,} bytes).")
    print("It is on the same disk as the repository: copy it somewhere else too.")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
