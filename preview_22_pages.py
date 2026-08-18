"""Render PNG previews of the interface pages, for visual checking only.

The previews are throwaway output; the deliverable stays the DOCX file whose
shapes are native Word shapes.
"""

import sys
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

import generate_22_interfaces_doc as gen

gen.build()

SCALE = 2.0
FONT_PATH = "/usr/share/fonts/google-noto/NotoSans-Regular.ttf"
BOLD_PATH = "/usr/share/fonts/google-noto/NotoSans-Bold.ttf"
OUT_DIR = Path("preview_22")
OUT_DIR.mkdir(exist_ok=True)

font_cache = {}


def get_font(size, bold):
    key = (round(size * SCALE), bold)
    if key not in font_cache:
        path = BOLD_PATH if bold and Path(BOLD_PATH).exists() else FONT_PATH
        font_cache[key] = ImageFont.truetype(path, max(int(size * SCALE), 6))
    return font_cache[key]


def rgb(value):
    return tuple(int(value[i:i + 2], 16) for i in (0, 2, 4)) if value else None


pages = {}
for shape in gen.SHAPE_RECTS:
    pages.setdefault(shape["page"], []).append(shape)

wanted = [int(a) for a in sys.argv[1:]] or sorted(p for p in pages if p)
for page in wanted:
    img = Image.new("RGB", (int(gen.PAGE_W * SCALE), int(gen.PAGE_H * SCALE)), "white")
    draw = ImageDraw.Draw(img)
    for shape in pages[page]:
        x0, y0 = shape["x"] * SCALE, shape["y"] * SCALE
        x1, y1 = (shape["x"] + shape["w"]) * SCALE, (shape["y"] + shape["h"]) * SCALE
        fill, line = rgb(shape["fill"]), rgb(shape["line"])
        if fill is None and line is None:
            pass  # transparent text box: draw the wording only
        elif shape["geom"] == "line":
            draw.line([x0, y0, x1, y1], fill=line or (80, 80, 80), width=1)
        elif shape["geom"] == "ellipse":
            draw.ellipse([x0, y0, x1, y1], fill=fill, outline=line, width=1)
        elif shape["geom"] == "roundRect":
            draw.rounded_rectangle([x0, y0, x1, y1], radius=4 * SCALE, fill=fill,
                                   outline=line, width=1)
        else:
            draw.rectangle([x0, y0, x1, y1], fill=fill, outline=line, width=1)
        if shape["text"]:
            lines = shape["text"].split("\n")
            font = get_font(shape["size"], shape["bold"])
            line_h = font.size * 1.22
            total = line_h * len(lines)
            ty = (y0 + y1) / 2 - total / 2
            for row in lines:
                width = draw.textlength(row, font=font)
                if shape["align"] == "center":
                    tx = (x0 + x1) / 2 - width / 2
                elif shape["align"] == "right":
                    tx = x1 - width - 2 * SCALE
                else:
                    tx = x0 + 2.5 * SCALE
                draw.text((tx, ty), row, font=font, fill=rgb(shape["color"]))
                ty += line_h
    path = OUT_DIR / f"page_{page:02d}.png"
    img.save(path)
    print("SAVED", path)
