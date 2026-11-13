"""Render the README art in assets/art/ from SVG to PNG with the installed Microsoft Edge (D46).

WHY A BROWSER AND NOT A LIBRARY. make-icon.py writes its PNG byte by byte because every
convenient imaging library is revenue-gated or drags in native binaries - exactly what this
repository's dependency guards keep out. That works for the icon's shapes; it cannot set
type. These images carry text, so a browser renders them, and the browser is one already on
the maintainer's machine rather than a dependency of the repository.

WHY PNG AND NOT THE SVG ITSELF. GitHub serves SVG correctly, but SVG text is drawn in the
viewer's fonts, so the layout shifts on a machine without Segoe UI. A PNG looks the same on
GitHub, on nuget.org and in a link card. The SVG stays as the source of truth.

MEASURED 2026-11-13 (Edge 142, Windows): --force-device-scale-factor=2 with --window-size
equal to the SVG's size gives exactly twice the size (880x330 -> 1760x660), in about 1 s.
--default-background-color=00000000 keeps the rounded corners transparent, so they do not
show as white squares in GitHub's dark mode.

EACH EDGE IS STARTED, AWAITED AND, IF IT OUTLIVES ITS SCREENSHOT, STOPPED BY ITS OWN PID -
never by name: other sessions on this machine run browsers, and orphaned headless Edge
processes have locked this machine's Windows account before.

Not run in CI (no Edge on the Linux runner); the PNGs are committed.

    python assets/render-art.py           # render every PNG
    python assets/render-art.py --check   # re-render to a temp folder; fail if any size differs
"""
import pathlib
import shutil
import struct
import subprocess
import sys
import tempfile

ART = pathlib.Path(__file__).resolve().parent / "art"
EDGE = pathlib.Path(r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe")

# (svg, png, css width, css height, scale)
ITEMS = [
    ("hero.svg", "hero.png", 880, 330, 2),
    ("runtimes.svg", "runtimes.png", 880, 320, 2),
    ("social-preview.svg", "social-preview.png", 1280, 640, 1),
]


def png_size(path):
    data = path.read_bytes()[:24]
    if data[:8] != b"\x89PNG\r\n\x1a\n":
        raise ValueError(f"{path} is not a PNG")
    return struct.unpack(">II", data[16:24])


def render(svg, png, width, height, scale):
    profile = pathlib.Path(tempfile.mkdtemp(prefix="render-art-"))
    if png.exists():
        png.unlink()
    args = [str(EDGE), "--headless=new", "--disable-gpu", "--hide-scrollbars", "--no-first-run",
            f"--user-data-dir={profile}", f"--force-device-scale-factor={scale}",
            f"--window-size={width},{height}", "--default-background-color=00000000",
            f"--screenshot={png}", svg.as_uri()]
    process = subprocess.Popen(args, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    try:
        process.wait(timeout=60)
    except subprocess.TimeoutExpired:
        process.kill()
        process.wait()
        raise SystemExit(f"Edge (PID {process.pid}) did not finish {svg.name} in 60 s; stopped that PID only.")
    finally:
        shutil.rmtree(profile, ignore_errors=True)
    expected = (width * scale, height * scale)
    actual = png_size(png)
    if actual != expected:
        raise SystemExit(f"{png.name}: {actual[0]}x{actual[1]}, expected {expected[0]}x{expected[1]}")
    return actual


def main(argv):
    if not EDGE.exists():
        raise SystemExit(f"Microsoft Edge not found at {EDGE}")
    check = "--check" in argv
    out_dir = pathlib.Path(tempfile.mkdtemp(prefix="render-art-out-")) if check else ART
    failed = False
    for svg, png, width, height, scale in ITEMS:
        size = render(ART / svg, out_dir / png, width, height, scale)
        if check and png_size(ART / png) != size:
            print(f"::error::{png} is {png_size(ART / png)}, a fresh render is {size}; run assets/render-art.py")
            failed = True
        else:
            print(f"{png}: {size[0]}x{size[1]}")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
