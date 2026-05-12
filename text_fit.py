from functools import lru_cache

import config
from utils import font


MIN_FONT_BY_KIND = {
    "header": 18,
    "subhead": 13,
    "body": 7,
    "caption": 6,
    "kicker": 6,
}


def text_box_px(el):
    return max(4, int(el.w / config.SLIDE_W * config.PREVIEW_W)), max(4, int(el.h / config.SLIDE_H * config.PREVIEW_H))


def point_to_px(points):
    return max(1, int(points / 72 * (config.PREVIEW_W / config.SLIDE_W)))


@lru_cache(maxsize=16384)
def measured_width(px_size, bold, text):
    return font(px_size, bold).getbbox(text)[2]


def wrapped_line_count(text, px_size, bold, max_width):
    words = text.split()
    if not words:
        return 1
    lines = 1
    space_width = measured_width(px_size, bold, " ")
    line_width = 0
    for word in words:
        word_width = measured_width(px_size, bold, word)
        test_width = word_width if line_width == 0 else line_width + space_width + word_width
        if test_width <= max_width:
            line_width = test_width
        else:
            lines += 1
            line_width = word_width
    return lines


def text_fits(el, font_size):
    box_w, box_h = text_box_px(el)
    pad = int(el.data.get("margin", 0.04) / config.SLIDE_W * config.PREVIEW_W)
    usable_w = max(4, box_w - pad * 2)
    usable_h = max(4, box_h - pad * 2)
    px_size = point_to_px(font_size)
    bold = el.data.get("bold", False)
    lines = wrapped_line_count(el.data["text"], px_size, bold, usable_w)
    line_h = int(px_size * 1.22)
    return lines * line_h <= usable_h


def truncate_to_fit(el, font_size):
    words = el.data["text"].split()
    if len(words) <= 1:
        return
    original = words[:]
    for count in range(len(original) - 1, 0, -1):
        el.data["text"] = " ".join(original[:count])
        if text_fits(el, font_size):
            return
    el.data["text"] = original[0]


def fit_text_element(el):
    if el.kind != "text":
        return el
    original = int(el.data["font_size"])
    min_size = MIN_FONT_BY_KIND.get(el.data.get("kind"), 6)
    if text_fits(el, original):
        return el
    if not text_fits(el, min_size):
        el.data["font_size"] = min_size
        truncate_to_fit(el, min_size)
        return el

    low, high = min_size, original
    best = min_size
    while low <= high:
        mid = (low + high) // 2
        if text_fits(el, mid):
            best = mid
            low = mid + 1
        else:
            high = mid - 1
    el.data["font_size"] = best
    return el


def fit_text_elements(elements):
    for el in elements:
        fit_text_element(el)
    return elements
