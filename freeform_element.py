import random

from pptx.util import Inches, Pt

from element_model import Element
from utils import ppt_color, rand_color


FREEFORM_PATTERNS = ["blob", "zigzag", "badge", "wave"]


def normalized_points(pattern):
    if pattern == "zigzag":
        return [(0.02, 0.25), (0.22, 0.05), (0.42, 0.32), (0.62, 0.08), (0.95, 0.75), (0.72, 0.95), (0.5, 0.65), (0.25, 0.92)]
    if pattern == "badge":
        return [(0.5, 0.0), (0.62, 0.25), (0.92, 0.18), (0.75, 0.5), (0.98, 0.78), (0.62, 0.72), (0.5, 1.0), (0.38, 0.72), (0.02, 0.78), (0.25, 0.5), (0.08, 0.18), (0.38, 0.25)]
    if pattern == "wave":
        return [(0.02, 0.55), (0.18, 0.2), (0.35, 0.5), (0.52, 0.8), (0.7, 0.45), (0.95, 0.58), (0.95, 0.92), (0.02, 0.92)]
    return [(0.12, 0.2), (0.35, 0.04), (0.7, 0.12), (0.95, 0.42), (0.82, 0.82), (0.45, 0.98), (0.08, 0.72), (0.02, 0.38)]


def make_freeform(x, y, w, h):
    pattern = random.choice(FREEFORM_PATTERNS)
    return Element(
        "freeform",
        x,
        y,
        w,
        h,
        {
            "pattern": pattern,
            "points": normalized_points(pattern),
            "fill": rand_color(),
            "line": rand_color(),
            "line_width": random.uniform(0.5, 2.8),
        },
    )


def add_freeform_to_pptx(slide, el):
    points = el.data["points"]
    start = points[0]
    builder = slide.shapes.build_freeform(Inches(el.x + start[0] * el.w), Inches(el.y + start[1] * el.h))
    scaled = [(Inches(el.x + px * el.w), Inches(el.y + py * el.h)) for px, py in points[1:]]
    builder.add_line_segments(scaled, close=True)
    shape = builder.convert_to_shape()
    shape.fill.solid()
    shape.fill.fore_color.rgb = ppt_color(el.data["fill"])
    shape.line.color.rgb = ppt_color(el.data["line"])
    shape.line.width = Pt(el.data["line_width"])


def add_freeform_to_png(draw, el, box):
    x1, y1, x2, y2 = box
    pts = [(int(x1 + px * (x2 - x1)), int(y1 + py * (y2 - y1))) for px, py in el.data["points"]]
    draw.polygon(pts, fill=(*el.data["fill"], 150), outline=(*el.data["line"], 255))
