import random

import config
from chart_element import make_chart
from connector_element import make_connector
from freeform_element import make_freeform
from image_element import make_image
from layout import place_element
from shape_element import make_shape
from table_element import make_table
from text_element import make_text
from svg_element import make_svg, make_svg_image


def add_many(elements, count, factory, size_range):
    for _ in range(count):
        elements.append(place_element(factory, elements, size_range))


def add_text(elements):
    for _ in range(config.TEXT_COUNT):
        if random.random() < config.TEXT_NARROW_BOX_PROBABILITY:
            kind = random.choices(["body", "subhead", "caption"], weights=[4, 2, 1], k=1)[0]
            elements.append(
                place_element(
                    lambda x, y, w, h: make_text(x, y, w, h, forced_kind=kind),
                    elements,
                    (0.95, 2.15, 1.0, 3.0),
                )
            )
        else:
            elements.append(place_element(make_text, elements, config.TEXT_SIZE_RANGE))


def make_elements():
    elements = []
    add_text(elements)
    add_many(elements, config.SHAPE_COUNT, make_shape, config.SHAPE_SIZE_RANGE)
    add_many(elements, config.TABLE_COUNT, make_table, config.TABLE_SIZE_RANGE)
    add_many(elements, config.IMAGE_COUNT, make_image, config.IMAGE_SIZE_RANGE)
    add_many(elements, config.CONNECTOR_COUNT, make_connector, config.CONNECTOR_SIZE_RANGE)
    add_many(elements, config.CHART_COUNT, make_chart, config.CHART_SIZE_RANGE)
    add_many(elements, config.FREEFORM_COUNT, make_freeform, config.FREEFORM_SIZE_RANGE)
    add_many(elements, config.SVG_COUNT, make_svg, config.SVG_SIZE_RANGE)
    add_many(elements, config.SVG_IMAGE_COUNT, make_svg_image, config.SVG_IMAGE_SIZE_RANGE)
    random.shuffle(elements)
    return elements
