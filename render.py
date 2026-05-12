from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.util import Inches

import config
from chart_element import add_chart_to_png, add_chart_to_pptx
from connector_element import add_connector_to_png, add_connector_to_pptx
from freeform_element import add_freeform_to_png, add_freeform_to_pptx
from image_element import add_image_to_png, add_image_to_pptx
from shape_element import add_shape_to_png, add_shape_to_pptx
from table_element import add_table_to_png, add_table_to_pptx
from text_element import add_text_to_png, add_text_to_pptx
from svg_element import add_svg_image_to_png, add_svg_to_png
from utils import px_box


def add_elements_to_slide(slide, elements, image_path):
    for el in elements:
        if el.kind == "text":
            add_text_to_pptx(slide, el)
        elif el.kind == "shape":
            add_shape_to_pptx(slide, el)
        elif el.kind == "table":
            add_table_to_pptx(slide, el)
        elif el.kind == "image":
            add_image_to_pptx(slide, el, image_path)
        elif el.kind == "connector":
            add_connector_to_pptx(slide, el)
        elif el.kind == "chart":
            add_chart_to_pptx(slide, el)
        elif el.kind == "freeform":
            add_freeform_to_pptx(slide, el)
        elif el.kind == "svg":
            pass
        elif el.kind == "svg_image":
            pass


def add_to_pptx(elements, image_path):
    prs = Presentation()
    prs.slide_width = Inches(config.SLIDE_W)
    prs.slide_height = Inches(config.SLIDE_H)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_elements_to_slide(slide, elements, image_path)
    prs.save(config.OUTPUT_PPTX)


def add_deck_to_pptx(slide_elements, image_path, output_pptx):
    prs = Presentation()
    prs.slide_width = Inches(config.SLIDE_W)
    prs.slide_height = Inches(config.SLIDE_H)
    for elements in slide_elements:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        add_elements_to_slide(slide, elements, image_path)
    prs.save(output_pptx)


def add_to_png(elements, image):
    canvas = Image.new("RGB", (config.PREVIEW_W, config.PREVIEW_H), config.PREVIEW_BACKGROUND)
    draw = ImageDraw.Draw(canvas, "RGBA")

    for el in elements:
        box = px_box(el)
        if el.kind == "text":
            add_text_to_png(draw, el, box)
        elif el.kind == "shape":
            add_shape_to_png(draw, el, box)
        elif el.kind == "table":
            add_table_to_png(draw, el, box)
        elif el.kind == "image":
            add_image_to_png(canvas, image, box)
        elif el.kind == "connector":
            add_connector_to_png(draw, el, box)
        elif el.kind == "chart":
            add_chart_to_png(draw, el, box)
        elif el.kind == "freeform":
            add_freeform_to_png(draw, el, box)
        elif el.kind == "svg":
            add_svg_to_png(draw, el, box)
        elif el.kind == "svg_image":
            add_svg_image_to_png(draw, el, box)

    canvas.save(config.OUTPUT_PNG)
