#!/usr/bin/env python3
"""Set the docs guard's floors from what this repository holds today (D48).

    python scripts/set-floors.py            set each floor to three quarters of its real count
    python scripts/set-floors.py --check    print the counts and the floors; change nothing
    python scripts/set-floors.py --self-test

scripts/check-docs-layout.py carries four floors. Each makes a walk that silently found far
fewer files than the repository holds fail, instead of reporting a clean pass. A floor helps
only while it sits near the real count: one left at a young repository's size lets most of an
older one vanish without a word. Run this after the tree grows a lot; never to make a red run
green - a count below its floor means a walk shrank, and that is the thing to explain.

The tracked counts are measured as CI measures them (tracked files only). The backlog counts
need the backlog on disk; run where it is, or they are left as they are.
"""

import json
import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
GUARD = ROOT / "scripts" / "check-docs-layout.py"
SHARE = (3, 4)


def floor_of(count):
    """Three quarters of a count, rounded down, and never below 1: a floor of 0 checks nothing."""
    return max(1, count * SHARE[0] // SHARE[1])


def read_floors(text, name):
    """The dict literal assigned to `name` in the guard's source."""
    match = re.search(rf"^{name} = (\{{[^}}\n]*\}})$", text, re.M)
    if not match:
        raise SystemExit(f"set-floors: no single-line '{name} = {{...}}' in {GUARD.name}")
    return json.loads(match.group(1))


def write_floors(text, name, floors):
    new = f"{name} = " + json.dumps(floors)
    return re.sub(rf"^{name} = \{{[^}}\n]*\}}$", lambda _: new, text, count=1, flags=re.M)


def measure():
    out = subprocess.run([sys.executable, str(GUARD), "--counts"], cwd=ROOT, capture_output=True,
                         text=True, encoding="utf-8", check=True).stdout
    return json.loads(out.strip().splitlines()[-1])


def self_test():
    bad = []

    def expect(name, ok):
        if not ok:
            bad.append(name)

    expect("three quarters, rounded down", floor_of(38) == 28 and floor_of(412) == 309)
    expect("never a floor of zero", floor_of(0) == 1 and floor_of(1) == 1)
    sample = 'A = 1\nTRACKED_FLOORS = {"md_files": 101, "src_files": 307}\nB = 2\n'
    expect("floors are read", read_floors(sample, "TRACKED_FLOORS") == {"md_files": 101, "src_files": 307})
    written = write_floors(sample, "TRACKED_FLOORS", {"md_files": 28, "src_files": 309})
    expect("floors are written in place, nothing else moves",
           written == 'A = 1\nTRACKED_FLOORS = {"md_files": 28, "src_files": 309}\nB = 2\n')
    try:
        read_floors("nothing here\n", "TRACKED_FLOORS")
        refused = False
    except SystemExit:
        refused = True
    expect("a missing floor line is refused, not read as no floors", refused)
    real = GUARD.read_text(encoding="utf-8")
    expect("the guard's own floor lines are readable",
           bool(read_floors(real, "TRACKED_FLOORS")) and bool(read_floors(real, "LOCAL_FLOORS")))

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
    counts = measure()
    raw = GUARD.read_bytes().decode("utf-8")
    text = raw.replace("\r\n", "\n")
    changed = False
    for name, measured in (("TRACKED_FLOORS", counts["tracked"]), ("LOCAL_FLOORS", counts["local"])):
        floors = read_floors(text, name)
        if not measured:
            print(f"{name}: not measured here (no backlog on disk); left at {floors}")
            continue
        wanted = {key: floor_of(measured[key]) for key in floors}
        for key in floors:
            print(f"{name}[{key}]: count {measured[key]}, floor {floors[key]} -> {wanted[key]}")
        if wanted != floors:
            text, changed = write_floors(text, name, wanted), True
    if "--check" in argv:
        return 0
    if changed:
        GUARD.write_bytes((text.replace("\n", "\r\n") if "\r\n" in raw else text).encode("utf-8"))
        print(f"rewrote the floors in {GUARD.relative_to(ROOT).as_posix()}")
    else:
        print("the floors already match")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
