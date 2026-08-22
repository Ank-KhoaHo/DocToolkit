"""Generate the DocToolkit social preview: a 1280x640 PNG, written byte-by-byte.

Same constraint as assets/make-icon.py, for the same reason: no imaging library, because every
convenient .NET/Python one is either revenue-gated or drags in native binaries - the exact thing
this repository's dependency guards exist to block. stdlib zlib + struct only.

GitHub's documented requirements (docs.github.com, "Customizing your repository's social media
preview"): PNG/JPG/GIF, under 1 MB, "at least 640 by 320 pixels (1280 by 640 pixels for best
display)". There is NO API for the upload - it is manual, under Settings > Social preview.

Palette and the page-with-fold motif are lifted from assets/make-icon.py so the card and the
package icon read as one thing.

Rendered at 3x and box-downsampled; hard edges at this size look cheap otherwise.

Two outputs, ONE design. The README banner is the same artwork on a shorter canvas rather than a
second file to keep in step - see YOFF below.

Usage:
    python make-social-preview.py social-preview.png       # 1280x640, the social preview
    python make-social-preview.py --banner banner.png      # 1280x360, the README banner
"""
import zlib, struct, sys

S = 3                                   # supersample factor
W, H = 1280, 640                        # GitHub's recommended size

# The README banner is the SAME artwork on a shorter canvas, not a second design. Every layout
# constant below is in card coordinates; YOFF slides that content up so a 360-tall canvas frames
# it. Two designs would drift apart, which is the whole reason this is one file.
#
# The numbers are derived, not guessed: content runs from y=196 (top of the wordmark) to y=476
# (bottom of the second subtitle), so 280px of content. On a 360 canvas a 155 shift leaves 41
# above and 39 below. Change a layout constant and re-derive rather than nudging this.
YOFF = 0

NAVY   = (0x1B, 0x2A, 0x4E)
PAGE   = (0xFF, 0xFF, 0xFF)
FOLD   = (0xC6, 0xD4, 0xE8)
LINE   = (0x3D, 0x7E, 0xBF)
ACCENT = (0xF0, 0xB4, 0x29)
MUTED  = (0x9F, 0xB2, 0xCE)

# 5x7 uppercase font. Only the glyphs this card needs, plus digits, so nothing is carried unused.
FONT = {
    'A': "01110 10001 10001 11111 10001 10001 10001", 'B': "11110 10001 10001 11110 10001 10001 11110",
    'C': "01110 10001 10000 10000 10000 10001 01110", 'D': "11110 10001 10001 10001 10001 10001 11110",
    'E': "11111 10000 10000 11110 10000 10000 11111", 'F': "11111 10000 10000 11110 10000 10000 10000",
    'G': "01110 10001 10000 10111 10001 10001 01111", 'H': "10001 10001 10001 11111 10001 10001 10001",
    'I': "11111 00100 00100 00100 00100 00100 11111", 'K': "10001 10010 10100 11000 10100 10010 10001",
    'L': "10000 10000 10000 10000 10000 10000 11111", 'M': "10001 11011 10101 10101 10001 10001 10001",
    'N': "10001 11001 10101 10011 10001 10001 10001", 'O': "01110 10001 10001 10001 10001 10001 01110",
    'P': "11110 10001 10001 11110 10000 10000 10000", 'R': "11110 10001 10001 11110 10100 10010 10001",
    'S': "01111 10000 10000 01110 00001 00001 11110", 'T': "11111 00100 00100 00100 00100 00100 00100",
    'U': "10001 10001 10001 10001 10001 10001 01110", 'V': "10001 10001 10001 10001 10001 01010 00100",
    'W': "10001 10001 10001 10101 10101 11011 10001", 'X': "10001 10001 01010 00100 01010 10001 10001",
    'Y': "10001 10001 01010 00100 00100 00100 00100", 'Z': "11111 00001 00010 00100 01000 10000 11111",
    '#': "01010 01010 11111 01010 11111 01010 01010", '.': "00000 00000 00000 00000 00000 01100 01100",
    '-': "00000 00000 00000 11111 00000 00000 00000", ' ': "00000 00000 00000 00000 00000 00000 00000",
}
GLYPH = {c: [r for r in v.split()] for c, v in FONT.items()}


def text_px(s, x0, y0, scale):
    """Pixels covered by `s` drawn at (x0, y0), 5x7 cells scaled up. Returns a set."""
    out = set()
    for i, ch in enumerate(s):
        g = GLYPH.get(ch)
        if g is None:
            raise KeyError("no glyph for %r - add it to FONT rather than silently dropping it" % ch)
        cx = x0 + i * 6 * scale
        for ry, row in enumerate(g):
            for rx, bit in enumerate(row):
                if bit == '1':
                    for dy in range(scale):
                        for dx in range(scale):
                            out.add((cx + rx * scale + dx, y0 + ry * scale + dy))
    return out


def text_w(s, scale):
    return (len(s) * 6 - 1) * scale


def rounded(x, y, x0, y0, x1, y1, r):
    if x < x0 or x > x1 or y < y0 or y > y1:
        return False
    cx = x0 + r if x < x0 + r else (x1 - r if x > x1 - r else x)
    cy = y0 + r if y < y0 + r else (y1 - r if y > y1 - r else y)
    return (x - cx) ** 2 + (y - cy) ** 2 <= r * r


# --- the icon motif, in its own 128-unit space, mapped onto the card ------------------------
ICON_N, ICON_SIZE, ICON_X, ICON_Y = 128.0, 300.0, 96, 170


def icon_at(x, y):
    """Colour of the page-with-fold motif at card coords, or None outside it."""
    u = (x - ICON_X) * ICON_N / ICON_SIZE
    v = (y - ICON_Y) * ICON_N / ICON_SIZE
    if not (0 <= u < ICON_N and 0 <= v < ICON_N):
        return None
    PX0, PY0, PX1, PY1, F = 32.0, 22.0, 96.0, 106.0, 22.0
    if not rounded(u, v, PX0, PY0, PX1, PY1, 4):
        return None
    if (u - (PX1 - F)) + (v - PY0) < F:
        return FOLD if u > PX1 - F - 1 else PAGE
    for i, (ly, lx1) in enumerate(((56, 84), (70, 84), (84, 68))):
        if ly <= v <= ly + 7 and 44 <= u <= lx1:
            return ACCENT if i == 2 else LINE
    return PAGE


# --- text layout ----------------------------------------------------------------------------
TX = 470
WORDMARK, WM_S, WM_Y = "DOCTOOLKIT", 11, 196
TAG,      TG_S, TG_Y = "HTML TO PDF IN C#", 7, 306
SUB1,     SB_S, SB1_Y = "NO BROWSER. NO NATIVE BINARIES.", 4, 396
SUB2,           SB2_Y = "RUNS WHERE .NET RUNS. MIT.", 448

wm  = text_px(WORDMARK, TX, WM_Y, WM_S)
tag = text_px(TAG,      TX, TG_Y, TG_S)
s1  = text_px(SUB1,     TX, SB1_Y, SB_S)
s2  = text_px(SUB2,     TX, SB2_Y, SB_S)

RULE_Y, RULE_H, RULE_W = 356, 6, text_w(TAG, TG_S)


def shade(px, py):
    x, y = px / S, py / S + YOFF
    xi, yi = int(x), int(y)
    if (xi, yi) in wm:
        return PAGE
    if (xi, yi) in tag:
        return ACCENT
    if (xi, yi) in s1 or (xi, yi) in s2:
        return MUTED
    if RULE_Y <= y < RULE_Y + RULE_H and TX <= x < TX + RULE_W:
        return LINE
    c = icon_at(x, y)
    if c is not None:
        return c
    return NAVY


def main(path):
    acc = [[[0, 0, 0] for _ in range(W)] for _ in range(H)]
    for py in range(H * S):
        fy = py // S
        row = acc[fy]
        for px in range(W * S):
            r, g, b = shade(px, py)
            cell = row[px // S]
            cell[0] += r; cell[1] += g; cell[2] += b

    n = S * S
    raw = bytearray()
    for row in acc:
        raw.append(0)                                   # filter: none
        for r, g, b in row:
            raw += bytes((r // n, g // n, b // n))

    def chunk(t, d):
        c = t + d
        return struct.pack(">I", len(d)) + c + struct.pack(">I", zlib.crc32(c) & 0xffffffff)

    png = (b"\x89PNG\r\n\x1a\n"
           + chunk(b"IHDR", struct.pack(">IIBBBBB", W, H, 8, 2, 0, 0, 0))   # 2 = truecolour RGB
           + chunk(b"IDAT", zlib.compress(bytes(raw), 9))
           + chunk(b"IEND", b""))

    with open(path, "wb") as f:
        f.write(png)
    print("wrote %s: %d bytes, %dx%d RGB (GitHub cap is 1 MB)" % (path, len(png), W, H))


if "--banner" in sys.argv:
    # Same artwork, README aspect. Written as a separate invocation rather than a second script.
    H, YOFF = 360, 155
    args = [a for a in sys.argv[1:] if a != "--banner"]
    main(args[0] if args else "banner.png")
else:
    main(sys.argv[1] if len(sys.argv) > 1 else "social-preview.png")
