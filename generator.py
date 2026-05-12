import random

import config
from chart_element import make_chart
from connector_element import make_connector
from freeform_element import make_freeform
from image_element import make_image
from layout import place_element, place_focus_element
from shape_element import make_shape
from table_element import make_table
from text_element import make_text
from svg_element import make_svg, make_svg_image
from utils import rand_text


COUNTS = {
    "text": "TEXT_COUNT",
    "shape": "SHAPE_COUNT",
    "table": "TABLE_COUNT",
    "image": "IMAGE_COUNT",
    "connector": "CONNECTOR_COUNT",
    "chart": "CHART_COUNT",
    "freeform": "FREEFORM_COUNT",
    "svg": "SVG_COUNT",
    "svg_image": "SVG_IMAGE_COUNT",
}


def initial_counts():
    return {kind: getattr(config, attr) for kind, attr in COUNTS.items()}


def mark_focus(el):
    el.data["focus"] = True
    return el


def make_focus_text(x, y, w, h):
    el = make_text(x, y, w, h, forced_kind=random.choice(["header", "subhead"]))
    el.data["text"] = rand_text(1, 4)
    el.data["font_size"] = random.randint(18, 32) if w < 2.7 else random.randint(24, 42)
    el.data["bold"] = True
    return mark_focus(el)


FOCUS_FACTORIES = {
    "text": make_focus_text,
    "shape": lambda x, y, w, h: mark_focus(make_shape(x, y, w, h)),
    "table": lambda x, y, w, h: mark_focus(make_table(x, y, w, h)),
    "image": lambda x, y, w, h: mark_focus(make_image(x, y, w, h)),
    "chart": lambda x, y, w, h: mark_focus(make_chart(x, y, w, h)),
    "freeform": lambda x, y, w, h: mark_focus(make_freeform(x, y, w, h)),
    "svg": lambda x, y, w, h: mark_focus(make_svg(x, y, w, h)),
    "svg_image": lambda x, y, w, h: mark_focus(make_svg_image(x, y, w, h)),
}


def choose_focus_kind(counts):
    available = [kind for kind in config.FOCUS_KIND_WEIGHTS if counts.get(kind, 0) > 0]
    if not available:
        return None
    weights = [config.FOCUS_KIND_WEIGHTS[kind] for kind in available]
    return random.choices(available, weights=weights, k=1)[0]


def add_focus_elements(elements, counts):
    min_count, max_count = config.FOCUS_COUNT_RANGE
    target_count = random.randint(min_count, max_count)
    target_count = min(target_count, sum(1 for kind in config.FOCUS_KIND_WEIGHTS if counts.get(kind, 0) > 0))
    size_range = config.FOCUS_SIZE_RANGES_BY_COUNT[target_count]

    for _ in range(target_count):
        kind = choose_focus_kind(counts)
        if kind is None:
            break
        candidate = place_focus_element(FOCUS_FACTORIES[kind], elements, size_range)
        if candidate is None:
            continue
        elements.append(candidate)
        counts[kind] -= 1


def add_many(elements, count, factory, size_range):
    for _ in range(count):
        el = place_element(factory, elements, size_range)
        if el is not None:
            elements.append(el)


def add_text(elements, count):
    for _ in range(count):
        if random.random() < config.TEXT_NARROW_BOX_PROBABILITY:
            kind = random.choices(["body", "caption"], weights=[4, 1], k=1)[0]
            el = place_element(
                lambda x, y, w, h: make_text(x, y, w, h, forced_kind=kind),
                elements,
                (1.45, 2.65, 1.0, 3.0),
            )
            if el is not None:
                elements.append(el)
        else:
            el = place_element(make_text, elements, config.TEXT_SIZE_RANGE)
            if el is not None:
                elements.append(el)


def make_elements():
    elements = []
    counts = initial_counts()
    add_focus_elements(elements, counts)
    add_text(elements, counts["text"])
    add_many(elements, counts["shape"], make_shape, config.SHAPE_SIZE_RANGE)
    add_many(elements, counts["table"], make_table, config.TABLE_SIZE_RANGE)
    add_many(elements, counts["image"], make_image, config.IMAGE_SIZE_RANGE)
    add_many(elements, counts["connector"], make_connector, config.CONNECTOR_SIZE_RANGE)
    add_many(elements, counts["chart"], make_chart, config.CHART_SIZE_RANGE)
    add_many(elements, counts["freeform"], make_freeform, config.FREEFORM_SIZE_RANGE)
    add_many(elements, counts["svg"], make_svg, config.SVG_SIZE_RANGE)
    add_many(elements, counts["svg_image"], make_svg_image, config.SVG_IMAGE_SIZE_RANGE)
    random.shuffle(elements)
    return elements
