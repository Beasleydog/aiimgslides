import random

from pptx.enum.dml import MSO_PATTERN_TYPE

import config
from element_model import Element
from utils import rand_color, rand_text, weighted_choice


PALETTES = [
    {"bg": (247, 248, 246), "text": (30, 35, 42), "muted": (115, 124, 135), "accent": (31, 117, 161), "accent2": (222, 133, 62), "panel": (236, 240, 242)},
    {"bg": (252, 249, 244), "text": (36, 32, 29), "muted": (123, 113, 103), "accent": (174, 73, 92), "accent2": (64, 130, 109), "panel": (242, 233, 222)},
    {"bg": (25, 30, 39), "text": (246, 247, 249), "muted": (177, 185, 196), "accent": (96, 165, 250), "accent2": (244, 180, 76), "panel": (43, 51, 64)},
    {"bg": (245, 247, 252), "text": (27, 44, 68), "muted": (103, 117, 137), "accent": (82, 96, 214), "accent2": (36, 178, 149), "panel": (232, 237, 247)},
    {"bg": (250, 250, 248), "text": (28, 28, 28), "muted": (115, 115, 115), "accent": (38, 132, 88), "accent2": (180, 70, 53), "panel": (239, 239, 235)},
    {"bg": (33, 115, 169), "text": (250, 252, 255), "muted": (218, 235, 245), "accent": (255, 198, 74), "accent2": (255, 117, 117), "panel": (47, 136, 190)},
    {"bg": (120, 58, 170), "text": (255, 250, 255), "muted": (230, 210, 240), "accent": (71, 221, 181), "accent2": (255, 211, 88), "panel": (139, 78, 190)},
    {"bg": (28, 143, 112), "text": (248, 255, 252), "muted": (205, 241, 230), "accent": (255, 238, 103), "accent2": (48, 78, 140), "panel": (43, 162, 130)},
    {"bg": (210, 72, 94), "text": (255, 252, 247), "muted": (255, 218, 220), "accent": (51, 63, 101), "accent2": (255, 206, 86), "panel": (225, 94, 113)},
    {"bg": (240, 176, 54), "text": (40, 38, 28), "muted": (94, 82, 58), "accent": (32, 119, 149), "accent2": (168, 54, 94), "panel": (249, 203, 91)},
]

FONT_PAIRS = [
    ("Aptos Display", "Aptos"),
    ("Georgia", "Aptos"),
    ("Trebuchet MS", "Verdana"),
    ("Arial", "Arial"),
]

BACKGROUND_PATTERNS = [
    getattr(MSO_PATTERN_TYPE, name)
    for name in dir(MSO_PATTERN_TYPE)
    if name.isupper() and name != "MIXED"
]

GRADIENT_TYPES = ["linear", "radial", "rectangular", "shape"]


def sample_theme():
    palette = random.choice(PALETTES).copy()
    title_font, body_font = random.choice(FONT_PAIRS)
    return {
        **palette,
        "title_font": title_font,
        "body_font": body_font,
        "topic": rand_text(1, 2).title(),
    }


def background(theme):
    style = weighted_choice(config.BACKGROUND_STYLE_WEIGHTS)
    data = {"style": style, "fill": theme["bg"]}
    if style == "gradient":
        data.update(
            {
                "gradient_type": random.choice(GRADIENT_TYPES),
                "colors": [
                    random.choice([theme["bg"], theme["panel"], rand_color()]),
                    random.choice([theme["accent"], theme["accent2"], rand_color()]),
                ],
                "angle": random.choice([0, 30, 45, 60, 90, 120, 135, 180, 225, 270, 315]),
                "focus": random.choice(["center", "top_left", "top_right", "bottom_left", "bottom_right"]),
            }
        )
    elif style == "pattern":
        data.update(
            {
                "pattern": random.choice(BACKGROUND_PATTERNS),
                "fore_color": random.choice([theme["accent"], theme["accent2"], rand_color()]),
                "back_color": random.choice([theme["bg"], theme["panel"], rand_color()]),
            }
        )
    return Element("background", 0, 0, 13.333, 7.5, data)


def jitter_box(box, amount=0.08):
    x, y, w, h = box
    return (
        x + random.uniform(-amount, amount),
        y + random.uniform(-amount, amount),
        max(0.2, w + random.uniform(-amount, amount)),
        max(0.2, h + random.uniform(-amount, amount)),
    )


def title_text(theme, text=None):
    return {
        "kind": "header",
        "text": text or theme["topic"],
        "font_size": random.randint(34, 50),
        "font_face": theme["title_font"],
        "color": theme["text"],
        "bold": True,
        "italic": False,
        "underline": False,
        "bg_color": None,
        "margin": 0.04,
    }


def body_text(theme, words=(8, 18), color=None):
    return {
        "kind": "body",
        "text": rand_text(*words),
        "font_size": random.randint(10, 15),
        "font_face": theme["body_font"],
        "color": color or theme["muted"],
        "bold": False,
        "italic": False,
        "underline": False,
        "bg_color": None,
        "margin": 0.04,
    }


def label_text(theme, text=None):
    return {
        "kind": "caption",
        "text": text or rand_text(1, 3).title(),
        "font_size": random.randint(9, 12),
        "font_face": theme["body_font"],
        "color": theme["accent"],
        "bold": True,
        "italic": False,
        "underline": False,
        "bg_color": None,
        "margin": 0.03,
    }
