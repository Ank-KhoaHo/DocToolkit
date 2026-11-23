#!/usr/bin/env python3
"""Fail when the icon, the banner or the cover is missing or has the wrong shape (D51).

WHY THIS EXISTS. Every app here has an icon, a banner and a cover. On 2026-11-20 this repo
had all three, but only the banner was held by anything: check-readme-shape.py asks for it
by name as the first thing in each README. The icon (the NuGet package icon) and the cover
(GitHub's social preview) could be deleted or replaced by a picture of the wrong shape, and
every check would stay green.

WHAT IT HOLDS. Exactly these three files, by path:

  - ICON   assets/icon.png                 square;
  - BANNER assets/art/banner.png           960 pixels wide or more, and about three times as
                                           wide as it is high (BANNER_RATIO, within BANNER_SLACK);
  - COVER  assets/art/social-preview.png   exactly 1280 by 640, GitHub's social preview size.

Each must be a PNG file. The size is read from the PNG header, so no image library is needed.

WHAT IT DOES NOT HOLD. assets/banner.png and assets/social-preview.png are older copies
beside assets/art/. They stay, because a published page may link to them, and this check
ignores them. Uploading the cover to GitHub's settings is a manual step: this check sees the
file in the repo, not what GitHub serves.

WHY A SCRIPT OF ITS OWN. check-readme-shape.py reads text; this reads pictures.

USAGE

    python scripts/check-art.py --self-test
    python scripts/check-art.py
    python scripts/check-art.py --root DIR    # check a copy of the tree (for planted faults)
"""
from __future__ import annotations

import os
import struct
import sys
import tempfile
import zlib

ICON = "assets/icon.png"
BANNER = "assets/art/banner.png"
COVER = "assets/art/social-preview.png"

BANNER_MIN_WIDTH = 960
BANNER_RATIO = 3.0
BANNER_SLACK = 0.5          # 2.5 to 3.5 times as wide as high
COVER_SIZE = (1280, 640)

PNG_SIGNATURE = b"\x89PNG\r\n\x1a\n"


def png_size(path):
    """Return (width, height) from a PNG's IHDR chunk, or None when it is not a PNG."""
    with open(path, "rb") as f:
        head = f.read(24)
    if len(head) < 24 or head[:8] != PNG_SIGNATURE or head[12:16] != b"IHDR":
        return None
    return struct.unpack(">II", head[16:24])


def check(root):
    """Return a list of error strings for the three pictures under root."""
    errors = []
    for rel in (ICON, BANNER, COVER):
        path = os.path.join(root, *rel.split("/"))
        if not os.path.isfile(path):
            errors.append(f"{rel} is missing")
            continue
        size = png_size(path)
        if size is None:
            errors.append(f"{rel} is not a PNG file")
            continue
        w, h = size
        if rel == ICON and w != h:
            errors.append(f"{rel} is {w} by {h}: the icon must be square")
        elif rel == BANNER and (w < BANNER_MIN_WIDTH or h == 0
                                or abs(w / h - BANNER_RATIO) > BANNER_SLACK):
            errors.append(f"{rel} is {w} by {h}: the banner must be {BANNER_MIN_WIDTH} wide or more "
                          f"and {BANNER_RATIO - BANNER_SLACK:g} to {BANNER_RATIO + BANNER_SLACK:g} "
                          f"times as wide as high")
        elif rel == COVER and (w, h) != COVER_SIZE:
            errors.append(f"{rel} is {w} by {h}: the cover must be {COVER_SIZE[0]} by {COVER_SIZE[1]}")
    return errors


# ---------------------------------------------------------------------------
# Self-test. Written before the implementation (D51, 2026-11-20); do not edit the
# cases to make them pass.

def _png(width, height):
    """A minimal valid PNG header: signature plus an IHDR chunk."""
    data = struct.pack(">IIBBBBB", width, height, 8, 6, 0, 0, 0)
    chunk = b"IHDR" + data
    return PNG_SIGNATURE + struct.pack(">I", len(data)) + chunk + struct.pack(">I", zlib.crc32(chunk))


_GOOD = {ICON: _png(128, 128), BANNER: _png(1920, 640), COVER: _png(1280, 640)}

# (name, the files that differ from _GOOD (None = missing), the text each error must hold)
_CASES = [
    ("the three right pictures pass", {}, []),
    ("a missing icon fails", {ICON: None}, [f"{ICON} is missing"]),
    ("a missing banner fails", {BANNER: None}, [f"{BANNER} is missing"]),
    ("a missing cover fails", {COVER: None}, [f"{COVER} is missing"]),
    ("an icon that is not square fails", {ICON: _png(128, 120)}, [f"{ICON} is 128 by 120"]),
    ("a large square icon passes", {ICON: _png(512, 512)}, []),
    ("a cover of the wrong size fails", {COVER: _png(1200, 630)}, [f"{COVER} is 1200 by 630"]),
    ("a cover twice the size fails", {COVER: _png(2560, 1280)}, [f"{COVER} is 2560 by 1280"]),
    # D54: a cover that shares one side with 1280 by 640. Without these, a rule that read the
    # width alone, or the height alone, passed every case above.
    ("a cover of the right width and the wrong height fails", {COVER: _png(1280, 630)},
     [f"{COVER} is 1280 by 630"]),
    ("a cover of the right height and the wrong width fails", {COVER: _png(1200, 640)},
     [f"{COVER} is 1200 by 640"]),
    ("a narrow banner fails", {BANNER: _png(900, 300)}, [f"{BANNER} is 900 by 300"]),
    ("a banner at the minimum width passes", {BANNER: _png(960, 320)}, []),
    ("a square banner fails", {BANNER: _png(1280, 1280)}, [f"{BANNER} is 1280 by 1280"]),
    ("a banner four times as wide fails", {BANNER: _png(2560, 640)}, [f"{BANNER} is 2560 by 640"]),
    ("a banner at 2.5 passes", {BANNER: _png(1600, 640)}, []),
    ("a file that is not a PNG fails", {ICON: b"GIF89a" + b"\0" * 40}, [f"{ICON} is not a PNG"]),
    ("an empty file fails", {COVER: b""}, [f"{COVER} is not a PNG"]),
    ("two faults give two errors", {ICON: None, COVER: _png(1, 1)},
     [f"{ICON} is missing", f"{COVER} is 1 by 1"]),
]


def _tree(root, changes):
    files = dict(_GOOD)
    files.update(changes)
    for rel, data in files.items():
        if data is None:
            continue
        path = os.path.join(root, *rel.split("/"))
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "wb") as f:
            f.write(data)


def self_test():
    bad = 0
    for name, changes, expected in _CASES:
        with tempfile.TemporaryDirectory() as root:
            _tree(root, changes)
            errors = check(root)
        exact = len(errors) == len(expected) and all(any(want in e for e in errors) for want in expected)
        if not exact:
            bad += 1
            print(f"::error::self-test case '{name}': expected {expected or 'pass'}, got {errors or 'pass'}")
    if bad:
        print(f"::error::self-test failed {bad} case(s)")
        return 1
    print("self-test passed")
    return 0


def main(argv):
    if "--self-test" in argv:
        return self_test()
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    if "--root" in argv:
        root = argv[argv.index("--root") + 1]
    errors = check(root)
    for line in errors:
        print(f"::error::{line}")
    if errors:
        return 1
    print(f"icon, banner and cover present, each the right shape ({ICON}, {BANNER}, {COVER})")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
