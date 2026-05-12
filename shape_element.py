import random

from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.util import Inches, Pt

import config
from element_model import Element
from utils import draw_rect_outline, ppt_color, rand_color, weighted_choice


SHAPE_FAMILIES = {
    "basic": [
        ("rect", MSO_AUTO_SHAPE_TYPE.RECTANGLE),
        ("round_rect", MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE),
        ("ellipse", MSO_AUTO_SHAPE_TYPE.OVAL),
        ("triangle", MSO_AUTO_SHAPE_TYPE.ISOSCELES_TRIANGLE),
        ("right_triangle", MSO_AUTO_SHAPE_TYPE.RIGHT_TRIANGLE),
        ("diamond", MSO_AUTO_SHAPE_TYPE.DIAMOND),
        ("parallelogram", MSO_AUTO_SHAPE_TYPE.PARALLELOGRAM),
        ("trapezoid", MSO_AUTO_SHAPE_TYPE.TRAPEZOID),
        ("pentagon", MSO_AUTO_SHAPE_TYPE.PENTAGON),
        ("hexagon", MSO_AUTO_SHAPE_TYPE.HEXAGON),
        ("octagon", MSO_AUTO_SHAPE_TYPE.OCTAGON),
        ("donut", MSO_AUTO_SHAPE_TYPE.DONUT),
        ("arc", MSO_AUTO_SHAPE_TYPE.BLOCK_ARC),
        ("pie", MSO_AUTO_SHAPE_TYPE.PIE),
        ("wave", MSO_AUTO_SHAPE_TYPE.WAVE),
    ],
    "arrows": [
        ("right_arrow", MSO_AUTO_SHAPE_TYPE.RIGHT_ARROW),
        ("left_arrow", MSO_AUTO_SHAPE_TYPE.LEFT_ARROW),
        ("up_arrow", MSO_AUTO_SHAPE_TYPE.UP_ARROW),
        ("down_arrow", MSO_AUTO_SHAPE_TYPE.DOWN_ARROW),
        ("left_right_arrow", MSO_AUTO_SHAPE_TYPE.LEFT_RIGHT_ARROW),
        ("up_down_arrow", MSO_AUTO_SHAPE_TYPE.UP_DOWN_ARROW),
        ("quad_arrow", MSO_AUTO_SHAPE_TYPE.QUAD_ARROW),
        ("bent_arrow", MSO_AUTO_SHAPE_TYPE.BENT_ARROW),
        ("uturn_arrow", MSO_AUTO_SHAPE_TYPE.U_TURN_ARROW),
        ("circular_arrow", MSO_AUTO_SHAPE_TYPE.CIRCULAR_ARROW),
        ("curved_right_arrow", MSO_AUTO_SHAPE_TYPE.CURVED_RIGHT_ARROW),
        ("striped_right_arrow", MSO_AUTO_SHAPE_TYPE.STRIPED_RIGHT_ARROW),
        ("notched_right_arrow", MSO_AUTO_SHAPE_TYPE.NOTCHED_RIGHT_ARROW),
        ("swoosh_arrow", MSO_AUTO_SHAPE_TYPE.SWOOSH_ARROW),
        ("right_arrow_callout", MSO_AUTO_SHAPE_TYPE.RIGHT_ARROW_CALLOUT),
        ("left_right_arrow_callout", MSO_AUTO_SHAPE_TYPE.LEFT_RIGHT_ARROW_CALLOUT),
    ],
    "flowchart": [
        ("flow_process", MSO_AUTO_SHAPE_TYPE.FLOWCHART_PROCESS),
        ("flow_decision", MSO_AUTO_SHAPE_TYPE.FLOWCHART_DECISION),
        ("flow_data", MSO_AUTO_SHAPE_TYPE.FLOWCHART_DATA),
        ("flow_document", MSO_AUTO_SHAPE_TYPE.FLOWCHART_DOCUMENT),
        ("flow_predefined", MSO_AUTO_SHAPE_TYPE.FLOWCHART_PREDEFINED_PROCESS),
        ("flow_terminator", MSO_AUTO_SHAPE_TYPE.FLOWCHART_TERMINATOR),
        ("flow_delay", MSO_AUTO_SHAPE_TYPE.FLOWCHART_DELAY),
        ("flow_display", MSO_AUTO_SHAPE_TYPE.FLOWCHART_DISPLAY),
    ],
    "callouts": [
        ("rect_callout", MSO_AUTO_SHAPE_TYPE.RECTANGULAR_CALLOUT),
        ("round_callout", MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGULAR_CALLOUT),
        ("oval_callout", MSO_AUTO_SHAPE_TYPE.OVAL_CALLOUT),
        ("cloud_callout", MSO_AUTO_SHAPE_TYPE.CLOUD_CALLOUT),
        ("line_callout", MSO_AUTO_SHAPE_TYPE.LINE_CALLOUT_2),
    ],
    "symbols": [
        ("cloud", MSO_AUTO_SHAPE_TYPE.CLOUD),
        ("heart", MSO_AUTO_SHAPE_TYPE.HEART),
        ("lightning", MSO_AUTO_SHAPE_TYPE.LIGHTNING_BOLT),
        ("sun", MSO_AUTO_SHAPE_TYPE.SUN),
        ("moon", MSO_AUTO_SHAPE_TYPE.MOON),
        ("gear6", MSO_AUTO_SHAPE_TYPE.GEAR_6),
        ("gear9", MSO_AUTO_SHAPE_TYPE.GEAR_9),
        ("cube", MSO_AUTO_SHAPE_TYPE.CUBE),
        ("can", MSO_AUTO_SHAPE_TYPE.CAN),
        ("funnel", MSO_AUTO_SHAPE_TYPE.FUNNEL),
        ("no_symbol", MSO_AUTO_SHAPE_TYPE.NO_SYMBOL),
    ],
    "stars": [
        ("star4", MSO_AUTO_SHAPE_TYPE.STAR_4_POINT),
        ("star5", MSO_AUTO_SHAPE_TYPE.STAR_5_POINT),
        ("star8", MSO_AUTO_SHAPE_TYPE.STAR_8_POINT),
        ("star16", MSO_AUTO_SHAPE_TYPE.STAR_16_POINT),
        ("explosion1", MSO_AUTO_SHAPE_TYPE.EXPLOSION1),
        ("explosion2", MSO_AUTO_SHAPE_TYPE.EXPLOSION2),
    ],
    "math": [
        ("plus", MSO_AUTO_SHAPE_TYPE.MATH_PLUS),
        ("minus", MSO_AUTO_SHAPE_TYPE.MATH_MINUS),
        ("multiply", MSO_AUTO_SHAPE_TYPE.MATH_MULTIPLY),
        ("divide", MSO_AUTO_SHAPE_TYPE.MATH_DIVIDE),
        ("equal", MSO_AUTO_SHAPE_TYPE.MATH_EQUAL),
        ("not_equal", MSO_AUTO_SHAPE_TYPE.MATH_NOT_EQUAL),
    ],
    "ribbons": [
        ("down_ribbon", MSO_AUTO_SHAPE_TYPE.DOWN_RIBBON),
        ("up_ribbon", MSO_AUTO_SHAPE_TYPE.UP_RIBBON),
        ("curved_down_ribbon", MSO_AUTO_SHAPE_TYPE.CURVED_DOWN_RIBBON),
        ("left_right_ribbon", MSO_AUTO_SHAPE_TYPE.LEFT_RIGHT_RIBBON),
    ],
}

DASH_STYLES = [
    MSO_LINE_DASH_STYLE.SOLID,
    MSO_LINE_DASH_STYLE.DASH,
    MSO_LINE_DASH_STYLE.DASH_DOT,
    MSO_LINE_DASH_STYLE.LONG_DASH,
    MSO_LINE_DASH_STYLE.ROUND_DOT,
    MSO_LINE_DASH_STYLE.SQUARE_DOT,
]


def random_shape_type():
    family = weighted_choice(config.SHAPE_FAMILY_WEIGHTS)
    name, ppt_type = random.choice(SHAPE_FAMILIES[family])
    return family, name, ppt_type


def make_shape(x, y, w, h):
    family, name, ppt_type = random_shape_type()
    outline_only = random.random() < config.SHAPE_OUTLINE_ONLY_PROBABILITY
    return Element(
        "shape",
        x,
        y,
        w,
        h,
        {
            "family": family,
            "shape": name,
            "ppt_type": ppt_type,
            "fill": rand_color(),
            "line": rand_color(),
            "line_width": random.uniform(0.5, 2.5),
            "outline_only": outline_only,
            "dash": random.choice(DASH_STYLES)
            if random.random() < config.SHAPE_DASHED_LINE_PROBABILITY
            else MSO_LINE_DASH_STYLE.SOLID,
        },
    )


def add_shape_to_pptx(slide, el):
    shp = slide.shapes.add_shape(
        el.data["ppt_type"],
        Inches(el.x),
        Inches(el.y),
        Inches(el.w),
        Inches(el.h),
    )
    if el.data["outline_only"]:
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = ppt_color(el.data["fill"])
    shp.line.color.rgb = ppt_color(el.data["line"])
    shp.line.width = Pt(el.data["line_width"])
    shp.line.dash_style = el.data["dash"]


def add_shape_to_png(draw, el, box):
    fill = None if el.data["outline_only"] else (*el.data["fill"], 165)
    line = (*el.data["line"], 255)
    width = max(1, int(el.data["line_width"] * 1.5))
    dashed = el.data["dash"] != MSO_LINE_DASH_STYLE.SOLID
    if "arrow" in el.data["shape"]:
        draw.rounded_rectangle(box, radius=5, fill=fill)
        draw_rect_outline(draw, box, line, width, dashed)
        x1, y1, x2, y2 = box
        draw.polygon([(x2, (y1 + y2) // 2), (x2 - min(26, x2 - x1), y1), (x2 - min(26, x2 - x1), y2)], fill=line)
    elif el.data["shape"] in {"ellipse", "donut", "sun", "moon"}:
        draw.ellipse(box, fill=fill, outline=line, width=3)
    elif "triangle" in el.data["shape"]:
        x1, y1, x2, y2 = box
        pts = [(x1 + (x2 - x1) // 2, y1), (x1, y2), (x2, y2)]
        draw.polygon(pts, fill=fill, outline=line)
    elif el.data["shape"] in {"diamond", "flow_decision"}:
        x1, y1, x2, y2 = box
        pts = [((x1 + x2) // 2, y1), (x2, (y1 + y2) // 2), ((x1 + x2) // 2, y2), (x1, (y1 + y2) // 2)]
        draw.polygon(pts, fill=fill, outline=line)
    elif "star" in el.data["shape"] or "explosion" in el.data["shape"]:
        x1, y1, x2, y2 = box
        pts = [
            ((x1 + x2) // 2, y1),
            (int(x1 + (x2 - x1) * 0.62), int(y1 + (y2 - y1) * 0.38)),
            (x2, int(y1 + (y2 - y1) * 0.42)),
            (int(x1 + (x2 - x1) * 0.7), int(y1 + (y2 - y1) * 0.62)),
            (int(x1 + (x2 - x1) * 0.78), y2),
            ((x1 + x2) // 2, int(y1 + (y2 - y1) * 0.75)),
            (int(x1 + (x2 - x1) * 0.22), y2),
            (int(x1 + (x2 - x1) * 0.3), int(y1 + (y2 - y1) * 0.62)),
            (x1, int(y1 + (y2 - y1) * 0.42)),
            (int(x1 + (x2 - x1) * 0.38), int(y1 + (y2 - y1) * 0.38)),
        ]
        draw.polygon(pts, fill=fill, outline=line)
    else:
        draw.rounded_rectangle(box, radius=6, fill=fill, outline=line, width=width)
