import random
import re
from functools import lru_cache
from pathlib import Path

from PIL import ImageFont
from pptx.dml.color import RGBColor

import config


_WORDS = None
FALLBACK_WORDS = [
    "signal",
    "vector",
    "canvas",
    "metric",
    "frame",
    "orbit",
    "focus",
    "pattern",
    "stream",
    "module",
]


def rand_color(alpha=False):
    values = [random.randint(25, 230) for _ in range(3)]
    return (*values, random.randint(90, 210)) if alpha else tuple(values)


def load_words():
    global _WORDS
    if _WORDS is not None:
        return _WORDS

    path = Path(config.WORDS_FILE)
    if path.exists():
        text = path.read_text(encoding="utf-8", errors="ignore")
        words = re.findall(r"[A-Za-z][A-Za-z-]*", text)
        _WORDS = [word.lower() for word in words if 2 <= len(word) <= 18]
    else:
        _WORDS = FALLBACK_WORDS

    if not _WORDS:
        _WORDS = FALLBACK_WORDS
    return _WORDS


def rand_text(min_words=1, max_words=6):
    words = load_words()
    return " ".join(random.choice(words) for _ in range(random.randint(min_words, max_words)))


def weighted_choice(weights):
    names = list(weights)
    return random.choices(names, weights=[weights[name] for name in names], k=1)[0]


def ppt_color(rgb):
    return RGBColor(*rgb[:3])


def px_box(el):
    return (
        int(el.x / config.SLIDE_W * config.PREVIEW_W),
        int(el.y / config.SLIDE_H * config.PREVIEW_H),
        int((el.x + el.w) / config.SLIDE_W * config.PREVIEW_W),
        int((el.y + el.h) / config.SLIDE_H * config.PREVIEW_H),
    )


@lru_cache(maxsize=128)
def font(size, bold=False):
    names = [
        "arialbd.ttf" if bold else "arial.ttf",
        "DejaVuSans-Bold.ttf" if bold else "DejaVuSans.ttf",
    ]
    for name in names:
        try:
            return ImageFont.truetype(name, size)
        except OSError:
            pass
    return ImageFont.load_default()


def add_wrapped_text(draw, text, box, fill, size, bold=False):
    x1, y1, x2, y2 = box
    fnt = font(size, bold)
    words = text.split()
    lines = []
    line = ""
    for word in words:
        test = f"{line} {word}".strip()
        if draw.textbbox((0, 0), test, font=fnt)[2] <= x2 - x1:
            line = test
        else:
            if line:
                lines.append(line)
            line = word
    if line:
        lines.append(line)

    line_h = int(size * 1.2)
    for i, line in enumerate(lines):
        y = y1 + i * line_h
        if y + line_h > y2:
            break
        draw.text((x1, y), line, fill=fill, font=fnt)


def draw_dashed_line(draw, start, end, fill, width=1, dash=10, gap=6):
    x1, y1 = start
    x2, y2 = end
    length = ((x2 - x1) ** 2 + (y2 - y1) ** 2) ** 0.5
    if length == 0:
        return
    dx = (x2 - x1) / length
    dy = (y2 - y1) / length
    pos = 0
    while pos < length:
        seg_end = min(pos + dash, length)
        draw.line(
            (
                x1 + dx * pos,
                y1 + dy * pos,
                x1 + dx * seg_end,
                y1 + dy * seg_end,
            ),
            fill=fill,
            width=width,
        )
        pos += dash + gap


def draw_rect_outline(draw, box, fill, width=1, dashed=False):
    if not dashed:
        draw.rectangle(box, outline=fill, width=width)
        return
    x1, y1, x2, y2 = box
    draw_dashed_line(draw, (x1, y1), (x2, y1), fill, width)
    draw_dashed_line(draw, (x2, y1), (x2, y2), fill, width)
    draw_dashed_line(draw, (x2, y2), (x1, y2), fill, width)
    draw_dashed_line(draw, (x1, y2), (x1, y1), fill, width)
