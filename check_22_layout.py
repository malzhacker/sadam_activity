"""Geometry sanity check for the 22-interface blueprint.

Checks per interface page:
  1. Reserved-zone check - the popup dialog floats in a fixed lower-right zone of
     the screen body, so no other text shape may sit under it.
  2. Collision check - no two text shapes may overlap, apart from the shapes that
     belong to the dialog itself.
Out-of-page placement is already asserted while the document is being built.
"""

import generate_22_interfaces_doc as gen

gen.build()

MODAL = (gen.MODAL_X, gen.MODAL_Y, gen.MODAL_W, gen.MODAL_H)
ZONE = (gen.MODAL_X - 9, gen.MODAL_Y - 9, gen.MODAL_W + 13, gen.MODAL_H + 13)


def rect(shape):
    return (shape["x"], shape["y"], shape["w"], shape["h"])


def overlap(a, b):
    ax, ay, aw, ah = a
    bx, by, bw, bh = b
    dx = min(ax + aw, bx + bw) - max(ax, bx)
    dy = min(ay + ah, by + bh) - max(ay, by)
    return dx * dy if dx > 0 and dy > 0 else 0.0


def in_modal(r):
    x, y, w, h = r
    return (x >= MODAL[0] - 10 and x + w <= MODAL[0] + MODAL[2] + 1
            and y >= MODAL[1] - 10 and y + h <= MODAL[1] + MODAL[3] + 1)


pages = {}
for shape in gen.SHAPE_RECTS:
    pages.setdefault(shape["page"], []).append(shape)

zone_hits, collisions = [], []
for page, shapes in sorted(pages.items()):
    if page == 0:
        continue
    outside = [s for s in shapes if not in_modal(rect(s))]
    for shape in outside:
        if shape["text"].strip() and overlap(rect(shape), ZONE):
            zone_hits.append((page, shape["text"][:32], rect(shape)))
    texted = [s for s in outside if s["text"].strip()]
    for i in range(len(texted)):
        for j in range(i + 1, len(texted)):
            a, b = texted[i], texted[j]
            area = overlap(rect(a), rect(b))
            if not area:
                continue
            smaller = min(a["w"] * a["h"], b["w"] * b["h"])
            if area / smaller > 0.02:
                collisions.append((page, a["text"][:26], b["text"][:26],
                                   round(area / smaller, 2)))

occluded = []
for page, shapes in sorted(pages.items()):
    if page == 0:
        continue
    for i, shape in enumerate(shapes):
        if not shape["text"].strip():
            continue
        area = shape["w"] * shape["h"]
        for later in shapes[i + 1:]:
            if later["fill"] is None or not area:
                continue
            if overlap(rect(shape), rect(later)) / area > 0.3:
                occluded.append((page, shape["text"][:26], "hidden by",
                                 later["name"], rect(later)))
                break

print("OCCLUDED_TEXT=", len(occluded))
for row in occluded[:40]:
    print("   O", row)
print("PAGES_WITH_SHAPES=", len([p for p in pages if p]))
print("TOTAL_SHAPES=", len(gen.SHAPE_RECTS))
print("ZONE_HITS=", len(zone_hits))
for row in zone_hits[:60]:
    print("   Z", row)
print("TEXT_COLLISIONS=", len(collisions))
for row in collisions[:60]:
    print("   C", row)
