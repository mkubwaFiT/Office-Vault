"""
Generate Trove's app icon — a gold magnifier over a stack of document sheets on
an indigo→violet rounded badge — using ONLY the standard library (no Pillow, which
the build excludes). Produces assets/trove.png (256) and assets/trove.ico (multi-size).

Run:  python assets/make_icon.py
"""
import os
import zlib
import math
import struct

BASE = 256
SS = 3                      # supersampling factor for antialiasing
HERE = os.path.dirname(os.path.abspath(__file__))


def _lerp(a, b, t):
    return tuple(int(round(a[i] + (b[i] - a[i]) * t)) for i in range(3))


def _rounded_rect(x, y, x0, y0, x1, y1, r):
    if x < x0 or x > x1 or y < y0 or y > y1:
        return False
    cx = min(max(x, x0 + r), x1 - r)
    cy = min(max(y, y0 + r), y1 - r)
    return (x - cx) ** 2 + (y - cy) ** 2 <= r * r


def _seg_dist(px, py, ax, ay, bx, by):
    dx, dy = bx - ax, by - ay
    if dx == 0 and dy == 0:
        return math.hypot(px - ax, py - ay)
    t = max(0.0, min(1.0, ((px - ax) * dx + (py - ay) * dy) / (dx * dx + dy * dy)))
    return math.hypot(px - (ax + t * dx), py - (ay + t * dy))


def _over(dst, src):
    sr, sg, sb, sa = src
    if sa == 0:
        return dst
    a = sa / 255.0
    return (int(sr * a + dst[0] * (1 - a)),
            int(sg * a + dst[1] * (1 - a)),
            int(sb * a + dst[2] * (1 - a)), 255)


# Palette
BG_TOP, BG_BOT = (79, 70, 229), (124, 58, 237)     # indigo -> violet
SHEET = (248, 250, 252)
SHEET_EDGE = (203, 213, 225)
GOLD = (245, 158, 11)
GOLD_HI = (253, 224, 71)
LENS = (255, 255, 255)


def _render(size):
    s = size * SS
    scale = s / BASE
    px = bytearray(size * size * 4)
    # magnifier geometry (base-256 coords)
    ring_cx, ring_cy, ring_r, ring_w = 150, 116, 52, 15
    h0 = (ring_cx + ring_r * 0.707, ring_cy + ring_r * 0.707)
    h1 = (212, 196)
    for yy in range(size):
        for xx in range(size):
            # supersample
            acc = [0, 0, 0, 0]
            for oy in range(SS):
                for ox in range(SS):
                    X = (xx * SS + ox + 0.5) / scale
                    Y = (yy * SS + oy + 0.5) / scale
                    col = (0, 0, 0, 0)
                    # badge background (rounded square, vertical gradient)
                    if _rounded_rect(X, Y, 18, 18, 238, 238, 52):
                        col = _lerp(BG_TOP, BG_BOT, (Y - 18) / 220.0) + (255,)
                        # document stack (three offset sheets)
                        for dx, dy in ((-16, 18), (-4, 8), (8, -2)):
                            if _rounded_rect(X, Y, 70 + dx, 78 + dy, 150 + dx, 188 + dy, 12):
                                col = _over(col, SHEET_EDGE + (255,))
                            if _rounded_rect(X, Y, 74 + dx, 82 + dy, 146 + dx, 184 + dy, 10):
                                col = _over(col, SHEET + (255,))
                        # magnifier lens fill
                        d = math.hypot(X - ring_cx, Y - ring_cy)
                        if d <= ring_r - ring_w / 2:
                            col = _over(col, LENS + (70,))
                        # magnifier handle (gold capsule)
                        if _seg_dist(X, Y, h0[0], h0[1], h1[0], h1[1]) <= 9:
                            col = _over(col, GOLD + (255,))
                        # magnifier ring
                        if abs(d - ring_r) <= ring_w / 2:
                            hi = 1 if (Y < ring_cy) else 0
                            col = _over(col, (GOLD_HI if hi else GOLD) + (255,))
                    for i in range(4):
                        acc[i] += col[i]
            n = SS * SS
            o = (yy * size + xx) * 4
            px[o:o + 4] = bytes(acc[i] // n for i in range(4))
    return px


def _png_bytes(size, px):
    raw = bytearray()
    for y in range(size):
        raw.append(0)
        raw.extend(px[y * size * 4:(y + 1) * size * 4])
    comp = zlib.compress(bytes(raw), 9)

    def chunk(typ, data):
        return (struct.pack(">I", len(data)) + typ + data
                + struct.pack(">I", zlib.crc32(typ + data) & 0xffffffff))

    return (b"\x89PNG\r\n\x1a\n"
            + chunk(b"IHDR", struct.pack(">IIBBBBB", size, size, 8, 6, 0, 0, 0))
            + chunk(b"IDAT", comp) + chunk(b"IEND", b""))


def _downscale(px, src, dst):
    out = bytearray(dst * dst * 4)
    ratio = src / dst
    for y in range(dst):
        for x in range(dst):
            acc = [0, 0, 0, 0]
            cnt = 0
            for sy in range(int(y * ratio), int((y + 1) * ratio)):
                for sx in range(int(x * ratio), int((x + 1) * ratio)):
                    o = (sy * src + sx) * 4
                    for i in range(4):
                        acc[i] += px[o + i]
                    cnt += 1
            o = (y * dst + x) * 4
            out[o:o + 4] = bytes(acc[i] // max(cnt, 1) for i in range(4))
    return out


def main():
    print("Rendering 256px master…")
    master = _render(256)
    png256 = _png_bytes(256, master)
    with open(os.path.join(HERE, "trove.png"), "wb") as f:
        f.write(png256)

    images = [(256, png256)]
    for sz in (64, 48, 32, 16):
        images.append((sz, _png_bytes(sz, _downscale(master, 256, sz))))

    # Assemble a PNG-payload ICO
    n = len(images)
    offset = 6 + n * 16
    entries, blobs = b"", b""
    for sz, png in images:
        d = 0 if sz >= 256 else sz
        entries += struct.pack("<BBBBHHII", d, d, 0, 0, 1, 32, len(png), offset)
        offset += len(png)
        blobs += png
    with open(os.path.join(HERE, "trove.ico"), "wb") as f:
        f.write(struct.pack("<HHH", 0, 1, n) + entries + blobs)
    print("Wrote assets/trove.png and assets/trove.ico  (sizes: %s)" %
          ", ".join(str(s) for s, _ in images))


if __name__ == "__main__":
    main()
