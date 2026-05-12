import random

from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.shapes import MSO_CONNECTOR
from pptx.util import Inches, Pt

from element_model import Element
from utils import draw_dashed_line, ppt_color, rand_color


CONNECTORS = [
    ("straight", MSO_CONNECTOR.STRAIGHT),
    ("elbow", MSO_CONNECTOR.ELBOW),
    ("curve", MSO_CONNECTOR.CURVE),
]

DASHES = [
    MSO_LINE_DASH_STYLE.SOLID,
    MSO_LINE_DASH_STYLE.DASH,
    MSO_LINE_DASH_STYLE.DASH_DOT,
    MSO_LINE_DASH_STYLE.LONG_DASH,
    MSO_LINE_DASH_STYLE.ROUND_DOT,
]


def make_connector(x, y, w, h):
    name, connector_type = random.choice(CONNECTORS)
    return Element(
        "connector",
        x,
        y,
        w,
        h,
        {
            "connector": name,
            "ppt_type": connector_type,
            "color": rand_color(),
            "width": random.uniform(0.5, 4.0),
            "dash": random.choice(DASHES),
            "flip": random.choice([True, False]),
        },
    )


def add_connector_to_pptx(slide, el):
    x1 = Inches(el.x)
    y1 = Inches(el.y + (el.h if el.data["flip"] else 0))
    x2 = Inches(el.x + el.w)
    y2 = Inches(el.y if el.data["flip"] else el.y + el.h)
    line = slide.shapes.add_connector(el.data["ppt_type"], x1, y1, x2, y2)
    line.line.color.rgb = ppt_color(el.data["color"])
    line.line.width = Pt(el.data["width"])
    line.line.dash_style = el.data["dash"]


def add_connector_to_png(draw, el, box):
    x1, y1, x2, y2 = box
    start = (x1, y2) if el.data["flip"] else (x1, y1)
    end = (x2, y1) if el.data["flip"] else (x2, y2)
    width = max(1, int(el.data["width"] * 1.4))
    fill = (*el.data["color"], 255)
    if el.data["dash"] == MSO_LINE_DASH_STYLE.SOLID:
        draw.line((*start, *end), fill=fill, width=width)
    else:
        draw_dashed_line(draw, start, end, fill, width)
