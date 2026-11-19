"""Draws the program icon at every size Windows uses, so it is sharp everywhere.

Run:  python tools/make_icons.py
Writes archiver/web/app.ico (16-256 px, for the .exe and the window) and
PNG copies used by the screen. Needs Pillow.
"""
from __future__ import annotations

import os

from PIL import Image, ImageDraw

HERE = os.path.dirname(os.path.abspath(__file__))
WEB = os.path.join(HERE, "..", "archiver", "web")

TOP = (24, 128, 214)     # gradient, light blue
BOTTOM = (10, 88, 168)   # gradient, deep blue
WHITE = (255, 255, 255, 255)
SLOT = (12, 94, 176, 255)

ICO_SIZES = [16, 20, 24, 32, 40, 48, 64, 96, 128, 256]
PNG_SIZES = [32, 48, 64, 128, 192, 256, 512]


def draw(size: int) -> Image.Image:
    """One icon, drawn for this exact size (small sizes get bolder shapes)."""
    ss = 8 if size <= 64 else 4  # draw bigger, then shrink smoothly
    n = size * ss
    small = size <= 24

    def px(f: float) -> int:
        return round(f * n)

    # background: rounded square with a vertical gradient
    grad = Image.new("RGBA", (n, n))
    gd = ImageDraw.Draw(grad)
    for y in range(n):
        k = y / (n - 1)
        c = tuple(round(TOP[i] + (BOTTOM[i] - TOP[i]) * k) for i in range(3))
        gd.line([(0, y), (n, y)], fill=c + (255,))
    mask = Image.new("L", (n, n), 0)
    inset = 0 if small else px(0.03)
    ImageDraw.Draw(mask).rounded_rectangle(
        [inset, inset, n - 1 - inset, n - 1 - inset], radius=px(0.2 if small else 0.22), fill=255)
    img = Image.new("RGBA", (n, n), (0, 0, 0, 0))
    img.paste(grad, (0, 0), mask)
    d = ImageDraw.Draw(img)

    # archive box: lid + body + handle slot
    if small:
        lid = (0.14, 0.22, 0.86, 0.40)
        body = (0.20, 0.45, 0.80, 0.82)
        slot = (0.36, 0.54, 0.64, 0.64)
        r_lid, r_body = 0.05, 0.06
    else:
        lid = (0.20, 0.25, 0.80, 0.39)
        body = (0.25, 0.43, 0.75, 0.76)
        slot = (0.40, 0.52, 0.60, 0.585)
        r_lid, r_body = 0.04, 0.06
    d.rounded_rectangle([px(v) for v in lid], radius=px(r_lid), fill=WHITE)
    d.rounded_rectangle([px(v) for v in body], radius=px(r_body), fill=WHITE)
    d.rounded_rectangle([px(v) for v in slot], radius=px((slot[3] - slot[1]) / 2), fill=SLOT)
    return img.resize((size, size), Image.LANCZOS)


def main() -> None:
    os.makedirs(WEB, exist_ok=True)
    images = {s: draw(s) for s in sorted(set(ICO_SIZES + PNG_SIZES))}
    big = images[256]
    big.save(os.path.join(WEB, "app.ico"), format="ICO",
             sizes=[(s, s) for s in ICO_SIZES],
             append_images=[images[s] for s in ICO_SIZES if s != 256])
    for s in PNG_SIZES:
        images[s].save(os.path.join(WEB, f"icon-{s}.png"), optimize=True)
    print("icons written to", os.path.normpath(WEB))


if __name__ == "__main__":
    main()
