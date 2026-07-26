"""Tile a family's tab renders into one contact sheet.

Side by side, a family that was built by one script looks like one thing. This file's
families do not, and a contact sheet is the fastest way to see that.
"""
import os
import re
import sys

from PIL import Image, ImageDraw, ImageFont

SP = os.path.dirname(os.path.abspath(__file__))
SHOTS = os.path.join(SP, "shots")
ORDER = {"1": ["1.1 Ampol Retail", "1.2 Customer", "1.3 Enterprise Data",
               "1.4 TDD Group Functions", "1.5 P&C", "1.6 Finance", "1.7 Infrastructure",
               "1.8 Energy Solutions & B2B", "1.9 Commercial Fuels", "1.10 Z Retail",
               "1.11 BP&T", "1.12 SA&D", "1.13 Cyber Roles"],
         "2": ["2.1 Ampol Retail", "2.2 Customer", "2.3 Enterprise Data",
               "2.4 TDD Group Functions", "2.5 P&C", "2.6 Finance", "2.7 Infrastructure",
               "2.8 Energy Solutions & B2B", "2.9 Commercial Fuels", "2.10 Z Retail",
               "2.11 COE Cyber", "2.12 BP&T", "2.13 SA&D", "2.14 EGI"],
         "3": ["3.1 Cost Bridge", "3.2 Overhead & Leadership", "3.3 Squad Detail",
               "3.4 COE Detail"]}


def stem(n):
    return re.sub(r"[^A-Za-z0-9]+", "_", n).strip("_")


def trim(im, pad=12):
    """Crop the white margin LibreOffice leaves around the printed area."""
    g = im.convert("L")
    bg = Image.new("L", im.size, 255)
    from PIL import ImageChops
    bbox = ImageChops.difference(g, bg).getbbox()
    if not bbox:
        return im
    l, t, r, b = bbox
    return im.crop((max(0, l - pad), max(0, t - pad),
                    min(im.width, r + pad), min(im.height, b + pad)))


def build(fam, cols=3, cell_w=900):
    names = [n for n in ORDER[fam]
             if os.path.exists(os.path.join(SHOTS, stem(n) + ".png"))]
    tiles = []
    for n in names:
        im = trim(Image.open(os.path.join(SHOTS, stem(n) + ".png")).convert("RGB"))
        h = max(1, int(im.height * cell_w / im.width))
        tiles.append((n, im.resize((cell_w, h), Image.LANCZOS)))
    if not tiles:
        return None
    cell_h = max(t.height for _, t in tiles)
    cell_h = min(cell_h, 1100)
    lab = 34
    rows = (len(tiles) + cols - 1) // cols
    W = cols * (cell_w + 16) + 16
    H = rows * (cell_h + lab + 16) + 16
    out = Image.new("RGB", (W, H), "#ffffff")
    d = ImageDraw.Draw(out)
    try:
        font = ImageFont.truetype(
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 22)
    except OSError:
        font = ImageFont.load_default()
    for i, (n, t) in enumerate(tiles):
        r, c = divmod(i, cols)
        x = 16 + c * (cell_w + 16)
        y = 16 + r * (cell_h + lab + 16)
        d.text((x, y + 4), n, fill="#000000", font=font)
        box = t.crop((0, 0, cell_w, min(t.height, cell_h)))
        out.paste(box, (x, y + lab))
        d.rectangle([x, y + lab, x + cell_w, y + lab + box.height],
                    outline="#c0c0c0", width=1)
    p = os.path.join(SHOTS, f"contact_{fam}x.png")
    out.save(p)
    return p, len(tiles)


if __name__ == "__main__":
    for fam in (sys.argv[1:] or ["1", "2", "3"]):
        r = build(fam)
        print(f"  {fam}.x -> {r[0]} ({r[1]} tabs)" if r else f"  {fam}.x: nothing to tile")
