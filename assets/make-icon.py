"""Generate the DocToolkit package icon: a 128x128 PNG, written byte-by-byte.

No imaging library on purpose - every convenient .NET/Python one is either revenue-gated or
drags in native binaries, which is exactly what this repo's dependency guards exist to block.
Rendered at 4x and box-downsampled, because hard edges at 128px look cheap.
"""
import zlib, struct, sys

S = 4                    # supersample factor
N = 128                  # final size
W = N * S

NAVY  = (0x1B, 0x2A, 0x4E)
PAGE  = (0xFF, 0xFF, 0xFF)
FOLD  = (0xC6, 0xD4, 0xE8)
LINE  = (0x3D, 0x7E, 0xBF)
ACCENT= (0xF0, 0xB4, 0x29)


def rounded(x, y, x0, y0, x1, y1, r):
    """Is (x, y) inside the rounded rectangle?"""
    if x < x0 or x > x1 or y < y0 or y > y1:
        return False
    cx = x0 + r if x < x0 + r else (x1 - r if x > x1 - r else x)
    cy = y0 + r if y < y0 + r else (y1 - r if y > y1 - r else y)
    return (x - cx) ** 2 + (y - cy) ** 2 <= r * r


def shade(px, py):
    """Colour for one supersampled pixel, in 128-space floats."""
    x, y = px / S, py / S

    # Background: rounded square, full bleed.
    if not rounded(x, y, 0, 0, N - 1, N - 1, 24):
        return None                                    # transparent outside

    # Page body: x 32..96, y 22..106, with the top-right corner cut for the fold.
    PX0, PY0, PX1, PY1, F = 32.0, 22.0, 96.0, 106.0, 22.0
    inside_page = rounded(x, y, PX0, PY0, PX1, PY1, 4)

    if inside_page:
        # The diagonal cut runs from (PX1-F, PY0) to (PX1, PY0+F).
        if (x - (PX1 - F)) + (y - PY0) < F:
            # Above/right of the diagonal -> the folded-back triangle.
            return FOLD if x > PX1 - F - 1 else PAGE
        # Text lines.
        for i, (ly, lx1) in enumerate(((56, 84), (70, 84), (84, 68))):
            if ly <= y <= ly + 7 and 44 <= x <= lx1:
                return ACCENT if i == 2 else LINE
        return PAGE

    return NAVY


def main(path):
    # Render supersampled, accumulate into the final grid.
    acc = [[[0, 0, 0, 0] for _ in range(N)] for _ in range(N)]
    for py in range(W):
        fy = py // S
        for px in range(W):
            c = shade(px, py)
            cell = acc[fy][px // S]
            if c is None:
                cell[3] += 0                            # transparent contributes nothing
            else:
                cell[0] += c[0]; cell[1] += c[1]; cell[2] += c[2]; cell[3] += 255

    raw = bytearray()
    n = S * S
    for row in acc:
        raw.append(0)                                   # filter: none
        for r, g, b, a in row:
            if a == 0:
                raw += bytes((0, 0, 0, 0))
            else:
                # Un-premultiply: colour was only summed for covered samples.
                cov = a // 255
                raw += bytes((r // cov, g // cov, b // cov, a // n))

    def chunk(t, d):
        c = t + d
        return struct.pack(">I", len(d)) + c + struct.pack(">I", zlib.crc32(c) & 0xffffffff)

    png = (b"\x89PNG\r\n\x1a\n"
           + chunk(b"IHDR", struct.pack(">IIBBBBB", N, N, 8, 6, 0, 0, 0))   # 6 = RGBA
           + chunk(b"IDAT", zlib.compress(bytes(raw), 9))
           + chunk(b"IEND", b""))

    with open(path, "wb") as f:
        f.write(png)
    print(f"wrote {path}: {len(png)} bytes, {N}x{N} RGBA")


main(sys.argv[1])
