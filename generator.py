import random

import config
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE

from chart_element import make_chart
from connector_element import make_connector
from design_system import background, body_text, jitter_box, label_text, sample_theme, title_text
from element_model import Element
from freeform_element import make_freeform
from image_element import make_image
from shape_element import make_shape
from table_element import make_table
from svg_element import make_svg, make_svg_image
from utils import rand_color, rand_text


SLIDE_W = 13.333
SLIDE_H = 7.5
MARGIN = 0.55


def mix(a, b, t=0.5):
    return tuple(int(x * (1 - t) + y * t) for x, y in zip(a, b))


def maybe_color(theme, role="accent"):
    if random.random() < config.RANDOM_COLOR_VARIATION_PROBABILITY:
        return rand_color()
    return theme[role]


def maybe_dash():
    if random.random() > config.RANDOM_PROPERTY_VARIATION_PROBABILITY:
        return MSO_LINE_DASH_STYLE.SOLID
    return random.choice(
        [
            MSO_LINE_DASH_STYLE.SOLID,
            MSO_LINE_DASH_STYLE.DASH,
            MSO_LINE_DASH_STYLE.DASH_DOT,
            MSO_LINE_DASH_STYLE.LONG_DASH,
            MSO_LINE_DASH_STYLE.ROUND_DOT,
            MSO_LINE_DASH_STYLE.SQUARE_DOT,
        ]
    )


def contrast_text(bg):
    return (246, 248, 252) if sum(bg) < 390 else (28, 30, 36)


def random_background_color():
    if random.random() < 0.65:
        return rand_color()
    return random.choice(
        [
            (31, 87, 166),
            (126, 58, 173),
            (29, 151, 112),
            (198, 64, 91),
            (232, 163, 45),
            (61, 64, 91),
            (18, 129, 148),
            (111, 88, 201),
        ]
    )


def maybe_colorful_theme(theme):
    if random.random() >= config.COLORFUL_BACKGROUND_PROBABILITY:
        return theme
    bg = random_background_color()
    theme = theme.copy()
    theme["bg"] = bg
    theme["text"] = contrast_text(bg)
    theme["muted"] = mix(theme["text"], bg, 0.45)
    theme["panel"] = mix(bg, (255, 255, 255), random.uniform(0.12, 0.36)) if sum(bg) < 430 else mix(bg, (0, 0, 0), random.uniform(0.08, 0.22))
    return theme


def text_el(box, data, jitter=0.025):
    x, y, w, h = jitter_box(box, jitter)
    return Element("text", x, y, w, h, data)


def panel_el(box, theme, fill=None, line=None, radius=True, outline=False):
    x, y, w, h = jitter_box(box, 0.02)
    el = Element(
        "shape",
        x,
        y,
        w,
        h,
        {
            "family": "panel",
            "shape": "round_rect" if radius else "rect",
            "ppt_type": MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE if radius else MSO_AUTO_SHAPE_TYPE.RECTANGLE,
            "fill": fill or theme["panel"],
            "line": line or theme["panel"],
            "line_width": random.uniform(0.25, 3.0)
            if random.random() < config.RANDOM_PROPERTY_VARIATION_PROBABILITY
            else random.uniform(0.35, 1.3),
            "outline_only": outline,
            "dash": maybe_dash(),
        },
    )
    el.data["line_width"] *= config.BORDER_WIDTH_MULTIPLIER
    return el


def add_background_treatment(elements, theme):
    style = random.choices(
        ["plain", "top_band", "side_band", "corner_panel", "split_bg", "thin_rules", "large_soft_panel"],
        weights=[2.5, 1.0, 0.9, 0.8, 0.7, 0.8, 1.0],
        k=1,
    )[0]
    if style == "top_band":
        elements.append(panel_el((0, 0, SLIDE_W, random.uniform(0.55, 1.1)), theme, fill=theme["panel"], radius=False))
    elif style == "side_band":
        side_w = random.uniform(0.55, 1.35)
        x = 0 if random.random() < 0.5 else SLIDE_W - side_w
        elements.append(panel_el((x, 0, side_w, SLIDE_H), theme, fill=theme["panel"], radius=False))
    elif style == "corner_panel":
        w = random.uniform(4.0, 7.5)
        h = random.uniform(2.3, 4.5)
        x = random.choice([0, SLIDE_W - w])
        y = random.choice([0, SLIDE_H - h])
        elements.append(panel_el((x, y, w, h), theme, fill=theme["panel"], radius=False))
    elif style == "split_bg":
        if random.random() < 0.5:
            elements.append(panel_el((0, 0, random.uniform(4.8, 7.0), SLIDE_H), theme, fill=theme["panel"], radius=False))
        else:
            elements.append(panel_el((0, random.uniform(3.0, 4.4), SLIDE_W, SLIDE_H), theme, fill=theme["panel"], radius=False))
    elif style == "thin_rules":
        for _ in range(random.randint(1, 3)):
            if random.random() < 0.5:
                elements.append(panel_el((random.uniform(0.5, 11.0), random.uniform(0.4, 6.7), random.uniform(1.2, 3.8), 0.025), theme, fill=theme["accent"], radius=False))
            else:
                elements.append(panel_el((random.uniform(0.5, 12.2), random.uniform(0.4, 5.5), 0.025, random.uniform(0.8, 2.8)), theme, fill=theme["accent"], radius=False))
    elif style == "large_soft_panel":
        elements.append(panel_el((random.uniform(0.45, 2.5), random.uniform(0.55, 1.8), random.uniform(5.5, 9.5), random.uniform(3.5, 5.5)), theme, fill=mix(theme["panel"], theme["bg"], 0.35)))


def add_large_background_shapes(elements, theme):
    if random.random() > config.BACKGROUND_SHAPE_PROBABILITY:
        return
    min_count, max_count = config.BACKGROUND_SHAPE_COUNT_RANGE
    for _ in range(random.randint(min_count, max_count)):
        w = random.uniform(3.0, 9.0)
        h = random.uniform(1.4, 5.2)
        x = random.uniform(-1.2, SLIDE_W - w + 1.2)
        y = random.uniform(-0.9, SLIDE_H - h + 0.9)
        el = themed_shape((x, y, w, h), theme)
        el.data["fill"] = mix(rand_color(), theme["bg"], random.uniform(0.15, 0.45))
        el.data["line"] = rand_color()
        el.data["line_width"] = random.uniform(2.0, 8.0)
        el.data["outline_only"] = random.random() < 0.28
        el.data["dash"] = maybe_dash()
        elements.append(el)


def inset(box, dx=0.25, dy=0.25):
    x, y, w, h = box
    return x + dx, y + dy, max(0.2, w - dx * 2), max(0.2, h - dy * 2)


def split_vertical(box, parts=2, gap=0.35):
    x, y, w, h = box
    weights = [random.uniform(0.8, 1.4) for _ in range(parts)]
    total = sum(weights)
    usable = w - gap * (parts - 1)
    out = []
    cursor = x
    for weight in weights:
        bw = usable * weight / total
        out.append((cursor, y, bw, h))
        cursor += bw + gap
    return out


def split_horizontal(box, parts=2, gap=0.28):
    x, y, w, h = box
    weights = [random.uniform(0.8, 1.35) for _ in range(parts)]
    total = sum(weights)
    usable = h - gap * (parts - 1)
    out = []
    cursor = y
    for weight in weights:
        bh = usable * weight / total
        out.append((x, cursor, w, bh))
        cursor += bh + gap
    return out


def grid_regions(box, cols, rows, gap=0.3):
    x, y, w, h = box
    cell_w = (w - gap * (cols - 1)) / cols
    cell_h = (h - gap * (rows - 1)) / rows
    return [
        (x + c * (cell_w + gap), y + r * (cell_h + gap), cell_w, cell_h)
        for r in range(rows)
        for c in range(cols)
    ]


def area(box):
    return box[2] * box[3]


def title_data(theme, text=None, scale=1.0):
    data = title_text(theme, text or rand_text(2, 4).title())
    data["font_size"] = int(random.randint(26, 46) * scale)
    return data


def themed_image(box, theme, svg=False):
    x, y, w, h = jitter_box(box, 0.035)
    el = make_svg_image(x, y, w, h) if svg else make_image(x, y, w, h)
    el.data["focus"] = area(box) > 8
    if "mask_line" in el.data:
        el.data["mask_line"] = maybe_color(theme, "accent")
    if "line" in el.data:
        el.data["line"] = maybe_color(theme, "accent")
    if random.random() < config.RANDOM_PROPERTY_VARIATION_PROBABILITY and "crop" in el.data:
        el.data["crop"] = tuple(random.uniform(0, config.IMAGE_MAX_CROP * 1.5) for _ in range(4))
    return el


def themed_chart(box, theme):
    x, y, w, h = jitter_box(box, 0.03)
    el = make_chart(x, y, w, h)
    el.data["focus"] = area(box) > 8
    return el


def themed_table(box, theme):
    x, y, w, h = jitter_box(box, 0.03)
    el = make_table(x, y, w, h)
    el.data["header"] = maybe_color(theme, "accent")
    el.data["body"] = maybe_color(theme, "panel")
    el.data["band"] = maybe_color(theme, "panel")
    el.data["border"] = maybe_color(theme, "muted")
    el.data["text_color"] = maybe_color(theme, "text")
    if random.random() < config.RANDOM_PROPERTY_VARIATION_PROBABILITY:
        el.data["border_width"] = random.uniform(0.2, 4.0)
        el.data["font_size"] = random.randint(5, 15)
    return el


def themed_shape(box, theme, focus=False):
    x, y, w, h = jitter_box(box, 0.04)
    el = make_shape(x, y, w, h)
    el.data["fill"] = random.choice([maybe_color(theme, "accent"), maybe_color(theme, "accent2"), maybe_color(theme, "panel")])
    el.data["line"] = random.choice([maybe_color(theme, "accent"), maybe_color(theme, "muted")])
    if random.random() < config.RANDOM_PROPERTY_VARIATION_PROBABILITY:
        el.data["outline_only"] = random.random() < 0.35
        el.data["line_width"] = random.uniform(0.25, 4.0) * config.BORDER_WIDTH_MULTIPLIER
        el.data["dash"] = maybe_dash()
    else:
        el.data["line_width"] *= config.BORDER_WIDTH_MULTIPLIER
    el.data["focus"] = focus
    return el


def text_block(box, theme, level="normal"):
    x, y, w, h = inset(box, 0.15, 0.12)
    els = []
    if random.random() < 0.45:
        els.append(text_el((x, y, w, 0.28), label_text(theme, rand_text(1, 2).upper())))
        y += 0.48
        h -= 0.48
    title_h = min(1.05, h * 0.38)
    scale = 0.78 if w < 2.8 else 1.0
    if level == "focus":
        scale += 0.12
    title = title_data(theme, rand_text(1, 4).title(), scale)
    title["font_size"] = min(title["font_size"], 38 if w < 3.5 else 46)
    title["color"] = maybe_color(theme, "text")
    if random.random() < config.RANDOM_PROPERTY_VARIATION_PROBABILITY:
        title["italic"] = random.random() < 0.25
        title["underline"] = random.random() < 0.15
        title["bg_color"] = rand_color() if random.random() < 0.18 else None
    els.append(text_el((x, y, w, title_h), title))
    if h - title_h > 0.55:
        els.append(text_el((x, y + title_h + 0.25, w, h - title_h - 0.25), body_text(theme, (8, 22))))
    return els


def image_block(box, theme, focus=False):
    els = []
    use_panel = random.random() < 0.28
    img_box = inset(box, 0.15, 0.15) if use_panel else box
    if use_panel:
        els.append(panel_el(box, theme, fill=theme["panel"]))
    els.append(themed_image(img_box, theme, svg=random.random() < (0.35 if focus else 0.18)))
    if random.random() < 0.35 and box[3] > 2.0:
        x, y, w, h = box
        els.append(text_el((x + 0.2, y + h - 0.45, w - 0.4, 0.28), label_text(theme, rand_text(1, 3).title())))
    return els


def image_grid_block(box, theme):
    cols = random.choice([2, 3])
    rows = random.choice([1, 2])
    regions = grid_regions(inset(box, 0.05, 0.05), cols, rows, 0.16)
    els = []
    for region in regions[: random.randint(2, len(regions))]:
        els.extend(image_block(region, theme, focus=False))
    return els


def chart_block(box, theme):
    els = []
    if random.random() < 0.55:
        els.append(panel_el(box, theme, fill=theme["panel"]))
        chart_box = inset(box, 0.25, 0.25)
    else:
        chart_box = box
    els.append(themed_chart(chart_box, theme))
    if random.random() < config.RANDOM_PROPERTY_VARIATION_PROBABILITY:
        els[-1].data["has_legend"] = random.choice([True, False])
        els[-1].data["has_title"] = random.choice([True, False])
    return els


def table_block(box, theme):
    return [themed_table(inset(box, 0.05, 0.05), theme)]


def stat_block(box, theme):
    x, y, w, h = inset(box, 0.18, 0.18)
    els = [panel_el(box, theme, fill=random.choice([theme["accent"], theme["panel"]]))]
    stat = title_data(theme, random.choice([f"{random.randint(12, 96)}%", f"{random.randint(2, 12)}x", str(random.randint(100, 900))]), 1.0)
    stat["font_size"] = random.randint(34, 56)
    stat["color"] = theme["bg"] if sum(theme["accent"]) < 360 else theme["text"]
    els.append(text_el((x, y + h * 0.16, w, min(1.15, h * 0.45)), stat))
    els.append(text_el((x, y + h * 0.62, w, min(0.7, h * 0.28)), body_text(theme, (4, 9), color=theme["text"])))
    return els


def metric_row_group(box, theme):
    count = random.choice([2, 3, 4])
    regions = split_vertical(box, count, 0.18)
    els = []
    for region in regions:
        x, y, w, h = region
        use_panel = random.random() < 0.65
        if use_panel:
            els.append(panel_el(region, theme, fill=random.choice([maybe_color(theme, "panel"), mix(maybe_color(theme, "accent"), theme["bg"], 0.72)])))
        stat = title_data(theme, random.choice([f"{random.randint(10, 98)}%", f"{random.randint(2, 9)}.{random.randint(0, 9)}x", str(random.randint(12, 240))]), 0.75)
        stat["font_size"] = random.randint(22, 38)
        stat["color"] = theme["accent"]
        els.append(text_el((x + 0.18, y + 0.18, w - 0.36, min(0.75, h * 0.45)), stat))
        if h > 1.15:
            els.append(text_el((x + 0.18, y + 0.92, w - 0.36, h - 1.02), body_text(theme, (3, 7))))
    return els


def card_group(box, theme):
    x, y, w, h = box
    count = random.choice([2, 3, 4])
    horizontal = w >= h * 1.35
    regions = split_vertical(box, count, 0.22) if horizontal else split_horizontal(box, count, 0.2)
    els = []
    for region in regions:
        rx, ry, rw, rh = region
        els.append(panel_el(region, theme, fill=theme["panel"]))
        if min(rw, rh) > 1.1 and random.random() < 0.75:
            els.append(themed_shape((rx + 0.22, ry + 0.22, 0.55, 0.55), theme))
        els.append(text_el((rx + 0.25, ry + 0.9, max(0.5, rw - 0.5), 0.45), label_text(theme, rand_text(1, 2).title())))
        if rh > 1.8:
            els.append(text_el((rx + 0.25, ry + 1.45, max(0.5, rw - 0.5), rh - 1.65), body_text(theme, (8, 15))))
    return els


def process_group(box, theme):
    x, y, w, h = box
    steps = random.randint(3, 5)
    horizontal = w >= h
    regions = split_vertical(box, steps, 0.22) if horizontal else split_horizontal(box, steps, 0.18)
    els = []
    centers = []
    for idx, region in enumerate(regions):
        rx, ry, rw, rh = region
        size = min(0.62, rw * 0.35, rh * 0.35)
        cx = rx + rw * random.uniform(0.25, 0.5)
        cy = ry + rh * random.uniform(0.25, 0.45)
        els.append(panel_el((cx, cy, size, size), theme, fill=theme["accent"], radius=True))
        els.append(text_el((rx + 0.05, ry + rh * 0.58, rw - 0.1, rh * 0.34), body_text(theme, (3, 7), color=theme["text"])))
        centers.append((cx + size / 2, cy + size / 2))
    for a, b in zip(centers, centers[1:]):
        x1, y1 = a
        x2, y2 = b
        line = make_connector(x1, y1, max(0.1, x2 - x1), max(0.1, y2 - y1))
        line.data["color"] = theme["muted"]
        els.append(line)
    return els


def quote_block(box, theme):
    x, y, w, h = inset(box, 0.2, 0.2)
    quote = title_data(theme, rand_text(5, 10).title(), 0.82)
    quote["italic"] = random.random() < 0.6
    quote["font_size"] = min(36, quote["font_size"])
    return [
        text_el((x, y, w, h * 0.62), quote),
        text_el((x, y + h * 0.72, w * 0.55, h * 0.18), label_text(theme, rand_text(1, 2).upper())),
    ]


def visual_accent(box, theme):
    choice = random.choice(["shape", "freeform", "svg"])
    if choice == "freeform":
        x, y, w, h = jitter_box(box, 0.04)
        el = make_freeform(x, y, w, h)
        el.data["fill"] = maybe_color(theme, "accent2")
        el.data["line"] = maybe_color(theme, "accent")
        return [el]
    if choice == "svg":
        x, y, w, h = jitter_box(box, 0.04)
        el = make_svg(x, y, w, h)
        el.data["primary"] = maybe_color(theme, "accent")
        el.data["secondary"] = maybe_color(theme, "accent2")
        return [el]
    return [themed_shape(box, theme)]


def overlay_text(theme):
    w = random.uniform(0.55, 2.4)
    h = random.uniform(0.18, 0.65)
    x = random.uniform(0.15, SLIDE_W - w - 0.15)
    y = random.uniform(0.15, SLIDE_H - h - 0.15)
    data = random.choice([label_text(theme, rand_text(1, 3).title()), body_text(theme, (1, 5))])
    data["font_size"] = random.randint(5, 18)
    data["color"] = rand_color()
    data["bold"] = random.random() < 0.45
    data["italic"] = random.random() < 0.25
    data["underline"] = random.random() < 0.18
    data["bg_color"] = rand_color() if random.random() < 0.25 else None
    return text_el((x, y, w, h), data, jitter=0)


def overlay_shape(theme):
    w = random.uniform(0.18, 1.4)
    h = random.uniform(0.15, 1.0)
    x = random.uniform(0.05, SLIDE_W - w - 0.05)
    y = random.uniform(0.05, SLIDE_H - h - 0.05)
    el = themed_shape((x, y, w, h), theme)
    el.data["fill"] = rand_color()
    el.data["line"] = rand_color()
    el.data["line_width"] = random.uniform(0.2, 3.5) * config.BORDER_WIDTH_MULTIPLIER
    el.data["outline_only"] = random.random() < 0.25
    el.data["dash"] = maybe_dash()
    return el


def overlay_connector(theme):
    w = random.uniform(0.4, 2.2)
    h = random.uniform(0.1, 1.4)
    x = random.uniform(0.05, SLIDE_W - w - 0.05)
    y = random.uniform(0.05, SLIDE_H - h - 0.05)
    el = make_connector(x, y, w, h)
    el.data["color"] = rand_color()
    el.data["width"] = random.uniform(0.25, 3.0)
    el.data["dash"] = maybe_dash()
    return el


def overlay_svg(theme):
    w = random.uniform(0.35, 1.6)
    h = random.uniform(0.25, 1.2)
    x = random.uniform(0.05, SLIDE_W - w - 0.05)
    y = random.uniform(0.05, SLIDE_H - h - 0.05)
    return visual_accent((x, y, w, h), theme)[0]


def add_overlay_elements(elements, theme):
    min_count, max_count = config.OVERLAY_ELEMENT_COUNT_RANGE
    count = random.randint(min_count, max_count)
    for _ in range(count):
        kind = random.choices(
            ["text", "shape", "connector", "svg"],
            weights=[
                config.OVERLAY_TEXT_PROBABILITY,
                config.OVERLAY_SHAPE_PROBABILITY,
                config.OVERLAY_CONNECTOR_PROBABILITY,
                config.OVERLAY_SVG_PROBABILITY,
            ],
            k=1,
        )[0]
        if kind == "text":
            elements.append(overlay_text(theme))
        elif kind == "shape":
            elements.append(overlay_shape(theme))
        elif kind == "connector":
            elements.append(overlay_connector(theme))
        else:
            elements.append(overlay_svg(theme))


GROUPS = {
    "text": text_block,
    "image": image_block,
    "image_grid": image_grid_block,
    "chart": chart_block,
    "table": table_block,
    "stat": stat_block,
    "metrics": metric_row_group,
    "cards": card_group,
    "process": process_group,
    "quote": quote_block,
    "accent": visual_accent,
}


def sample_content_regions(content_box):
    mode = random.choices(
        ["one_plus_side", "columns", "rows", "grid", "mosaic", "sidebar_stack", "hero_strip", "offset_cluster", "center_focus"],
        weights=[1.2, 1.0, 0.85, 0.95, 1.2, 0.9, 0.8, 1.1, 0.8],
        k=1,
    )[0]
    if mode == "one_plus_side":
        main_first = random.choice([True, False])
        regions = split_vertical(content_box, 2, random.uniform(0.35, 0.6))
        regions.sort(key=area, reverse=True)
        if not main_first:
            regions.reverse()
        return regions
    if mode == "columns":
        return split_vertical(content_box, random.choice([2, 3]), random.uniform(0.25, 0.45))
    if mode == "rows":
        return split_horizontal(content_box, random.choice([2, 3]), random.uniform(0.25, 0.45))
    if mode == "grid":
        return grid_regions(content_box, random.choice([2, 3]), 2, random.uniform(0.22, 0.4))
    if mode == "sidebar_stack":
        left, right = split_vertical(content_box, 2, random.uniform(0.3, 0.55))
        side, main = (left, right) if random.random() < 0.5 else (right, left)
        return [main] + split_horizontal(side, random.choice([2, 3]), random.uniform(0.2, 0.35))
    if mode == "hero_strip":
        top, bottom = split_horizontal(content_box, 2, random.uniform(0.3, 0.45))
        if random.random() < 0.5:
            top, bottom = bottom, top
        return [top] + split_vertical(bottom, random.choice([2, 3, 4]), random.uniform(0.18, 0.35))
    if mode == "offset_cluster":
        x, y, w, h = content_box
        main_w = random.uniform(w * 0.45, w * 0.68)
        main_h = random.uniform(h * 0.42, h * 0.7)
        main = (random.uniform(x, x + w - main_w), random.uniform(y, y + h - main_h), main_w, main_h)
        smalls = []
        for _ in range(random.randint(2, 4)):
            sw = random.uniform(w * 0.18, w * 0.34)
            sh = random.uniform(h * 0.18, h * 0.34)
            smalls.append((random.uniform(x, x + w - sw), random.uniform(y, y + h - sh), sw, sh))
        return [main] + smalls
    if mode == "center_focus":
        x, y, w, h = content_box
        main = (x + w * 0.22, y + h * 0.18, w * 0.56, h * 0.56)
        return [main, (x, y, w * 0.18, h), (x + w * 0.82, y, w * 0.18, h), (x + w * 0.22, y + h * 0.78, w * 0.56, h * 0.22)]

    top, bottom = split_horizontal(content_box, 2, random.uniform(0.28, 0.45))
    if random.random() < 0.5:
        return [top] + split_vertical(bottom, random.choice([2, 3]), random.uniform(0.25, 0.4))
    return split_vertical(top, random.choice([2, 3]), random.uniform(0.25, 0.4)) + [bottom]


def add_title_region(elements, theme):
    style = random.choices(["top", "left", "floating", "none", "center", "right", "bottom"], weights=[2.6, 0.9, 1.0, 0.35, 0.8, 0.55, 0.5], k=1)[0]
    if style == "top":
        box = (MARGIN, 0.35, random.uniform(5.0, 9.0), random.uniform(0.95, 1.35))
        title = title_data(theme, rand_text(2, 4).title())
        title["font_size"] = min(title["font_size"], 42)
        elements.append(text_el(box, title))
        return (MARGIN, 1.85, SLIDE_W - MARGIN * 2, SLIDE_H - 2.35)
    if style == "left":
        box = (0.65, 0.75, 2.85, 5.55)
        data = title_data(theme, rand_text(1, 3).title(), 0.72)
        data["font_size"] = min(data["font_size"], 30)
        elements.append(text_el(box, data))
        return (3.75, 0.65, SLIDE_W - 4.35, SLIDE_H - 1.25)
    if style == "floating":
        box = (random.uniform(0.65, 6.0), random.uniform(0.45, 1.4), random.uniform(3.5, 6.3), 0.9)
        elements.append(text_el(box, title_data(theme, rand_text(2, 4).title(), 0.9)))
        return (MARGIN, 1.45, SLIDE_W - MARGIN * 2, SLIDE_H - 1.95)
    if style == "center":
        box = (random.uniform(2.2, 3.5), 0.65, random.uniform(6.0, 8.8), 1.15)
        title = title_data(theme, rand_text(2, 5).title(), 1.0)
        elements.append(text_el(box, title))
        return (MARGIN, 2.0, SLIDE_W - MARGIN * 2, SLIDE_H - 2.45)
    if style == "right":
        box = (9.4, 0.65, 3.2, 5.8)
        data = title_data(theme, rand_text(1, 3).title(), 0.72)
        data["font_size"] = min(data["font_size"], 30)
        elements.append(text_el(box, data))
        return (0.65, 0.65, 8.2, SLIDE_H - 1.25)
    if style == "bottom":
        box = (0.75, 5.85, random.uniform(5.0, 8.8), 0.9)
        elements.append(text_el(box, title_data(theme, rand_text(2, 4).title(), 0.86)))
        return (MARGIN, MARGIN, SLIDE_W - MARGIN * 2, 5.0)
    return (MARGIN, MARGIN, SLIDE_W - MARGIN * 2, SLIDE_H - MARGIN * 2)


def choose_group(region, used):
    w, h = region[2], region[3]
    options = {
        "text": 1.4,
        "image": 1.3,
        "image_grid": 0.55 if w > 3.2 and h > 1.8 else 0.0,
        "chart": 0.8 if w > 2.8 and h > 1.8 else 0.0,
        "table": 0.45 if w > 3.2 and h > 1.8 else 0.0,
        "stat": 0.85 if w > 1.6 and h > 1.2 else 0.0,
        "metrics": 0.6 if w > 3.0 and h > 1.0 else 0.0,
        "cards": 0.65 if w > 3.0 and h > 2.0 else 0.0,
        "process": 0.6 if w > 4.0 or h > 3.0 else 0.0,
        "quote": 0.55 if w > 3.0 and h > 1.5 else 0.0,
        "accent": 0.45 if w * h < 4.5 else 0.06,
    }
    if used.get("image", 0) >= 2:
        options["image"] = 0.1
    if used.get("chart", 0) >= 1:
        options["chart"] = 0.1
    if used.get("table", 0) >= 1:
        options["table"] = 0.05
    names = [name for name, weight in options.items() if weight > 0]
    return random.choices(names, weights=[options[name] for name in names], k=1)[0]


def make_elements():
    theme = maybe_colorful_theme(sample_theme())
    elements = [background(theme)]
    add_large_background_shapes(elements, theme)
    add_background_treatment(elements, theme)
    content_box = add_title_region(elements, theme)
    regions = sample_content_regions(content_box)
    regions = [r for r in regions if r[2] > 1.15 and r[3] > 0.75]
    random.shuffle(regions)

    largest = max(regions, key=area) if regions else content_box
    used = {}
    group_count = random.randint(2, min(6, max(2, len(regions))))

    for idx, region in enumerate(regions[:group_count]):
        if region == largest and area(region) > 7 and random.random() < 0.75:
            group = random.choices(["image", "image_grid", "chart", "table", "text", "cards", "metrics"], weights=[1.7, 0.7, 1.2, 0.7, 1.0, 0.9, 0.7], k=1)[0]
        else:
            group = choose_group(region, used)
        used[group] = used.get(group, 0) + 1
        if group == "text":
            elements.extend(GROUPS[group](region, theme, "focus" if region == largest else "normal"))
        elif group == "image":
            elements.extend(GROUPS[group](region, theme, focus=region == largest))
        else:
            elements.extend(GROUPS[group](region, theme))

    if random.random() < 0.45:
        accent_box = (
            random.uniform(0.6, 11.0),
            random.uniform(0.8, 6.0),
            random.uniform(0.45, 1.25),
            random.uniform(0.35, 0.95),
        )
        elements.extend(visual_accent(accent_box, theme))

    add_overlay_elements(elements, theme)

    for el in elements:
        el.data.setdefault("theme_topic", theme["topic"])
        el.data.setdefault("archetype", "atomic_composition")
    return elements
