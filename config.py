"""Tweakable settings for the random slide generator."""

# Output files.
OUTPUT_PPTX = "random_slide.pptx"
OUTPUT_PNG = "random_slide.png"

# Every generated image element uses this URL for now.
EXAMPLE_IMAGE_URL = "https://picsum.photos/seed/aiimgslides/900/600"

# Counts per element type. Keep every object type nonzero, but avoid making the
# single slide too dense unless you explicitly want chaotic overlap.
TEXT_COUNT = 5
SHAPE_COUNT = 4
TABLE_COUNT = 1
IMAGE_COUNT = 2
CONNECTOR_COUNT = 2
CHART_COUNT = 1
FREEFORM_COUNT = 1
SVG_COUNT = 1
SVG_IMAGE_COUNT = 1

# Slide geometry in PowerPoint inches. 13.333 x 7.5 is widescreen 16:9.
SLIDE_W = 13.333
SLIDE_H = 7.5

# PNG preview size. Keep this 16:9 to match the PPTX slide.
PREVIEW_W = 1600
PREVIEW_H = 900

# Set to an integer for repeatable output, or None for a fresh random slide.
SEED = None

# How many random positions to test before choosing a final spot for one element.
# Higher values usually reduce overlap, but generation takes a little longer.
PLACEMENT_ATTEMPTS = 140

# Higher values make overlap more expensive during placement.
# 0 means elements ignore each other and can freely stack.
OVERLAP_PENALTY_WEIGHT = 55.0

# Extra overlap penalty when two elements of the same kind collide.
# Useful if, for example, text-on-text overlap is harder to read than text-on-shape.
SAME_KIND_OVERLAP_WEIGHT = 28.0

# Small penalty for placing an element very close to slide edges.
# Raise this to push content inward; set to 0 to allow edge-hugging layouts.
EDGE_PENALTY_WEIGHT = 0.8

# Minimum distance from the slide edge in inches when creating random boxes.
SLIDE_MARGIN = 0.22

# Extra invisible spacing used only during placement. PowerPoint objects like
# charts and wrapped text can visually extend outside their nominal boxes.
COLLISION_PADDING_BY_KIND = {
    "text": 0.18,
    "shape": 0.08,
    "table": 0.12,
    "image": 0.08,
    "connector": 0.10,
    "chart": 0.35,
    "freeform": 0.10,
    "svg": 0.10,
    "svg_image": 0.12,
}

# Soft size ranges by element type, in inches: (min_w, max_w, min_h, max_h).
# Text boxes intentionally include narrow widths so paragraphs wrap.
TEXT_SIZE_RANGE = (0.7, 3.8, 0.35, 2.4)
SHAPE_SIZE_RANGE = (0.35, 2.4, 0.25, 1.8)
TABLE_SIZE_RANGE = (2.0, 4.8, 1.0, 2.5)
IMAGE_SIZE_RANGE = (1.0, 3.0, 0.75, 2.1)
CONNECTOR_SIZE_RANGE = (0.8, 3.4, 0.4, 1.8)
CHART_SIZE_RANGE = (2.0, 4.5, 1.4, 2.7)
FREEFORM_SIZE_RANGE = (0.8, 2.5, 0.6, 1.9)
SVG_SIZE_RANGE = (1.0, 2.8, 0.8, 2.0)
SVG_IMAGE_SIZE_RANGE = (1.4, 3.0, 1.0, 2.2)

# Background color used by the generated PNG preview.
PREVIEW_BACKGROUND = (248, 248, 246)

# Text mix. Increase paragraph weight for denser lorem-like blocks, or header
# weight for bigger presentation-style labels.
TEXT_KIND_WEIGHTS = {
    "kicker": 1.0,      # tiny label, 1-3 words
    "header": 2.0,      # large headline, 2-7 words
    "subhead": 1.5,     # medium phrase, 5-14 words
    "body": 2.2,        # paragraph, usually wraps in narrow boxes
    "caption": 1.0,     # small annotation, 4-12 words
}

# Font size ranges by text kind, in PowerPoint points.
TEXT_FONT_SIZES = {
    "kicker": (8, 14),
    "header": (30, 64),
    "subhead": (16, 30),
    "body": (7, 13),
    "caption": (7, 12),
}

# Probability a text box is generated with a deliberately narrow width.
# Raise this when you want more visible line wrapping.
TEXT_NARROW_BOX_PROBABILITY = 0.65

# Shape family weights. These sample multiple concrete PowerPoint shapes per
# family, so raising "arrows" produces straight, bent, curved, and callout arrows.
SHAPE_FAMILY_WEIGHTS = {
    "basic": 2.5,
    "arrows": 2.5,
    "flowchart": 1.4,
    "callouts": 1.2,
    "symbols": 1.0,
    "stars": 0.9,
    "math": 0.6,
    "ribbons": 0.5,
}

# How often a shape has no fill and only a visible outline.
SHAPE_OUTLINE_ONLY_PROBABILITY = 0.18

# How often shape outlines use dash patterns instead of solid lines.
SHAPE_DASHED_LINE_PROBABILITY = 0.35

# Table structure and style variation. Wider ranges create more chaotic tables.
TABLE_ROW_RANGE = (2, 7)
TABLE_COL_RANGE = (2, 6)
TABLE_FONT_SIZE_RANGE = (6, 13)
TABLE_BORDER_WIDTH_RANGE = (0.25, 3.0)
TABLE_DASHED_BORDER_PROBABILITY = 0.35
TABLE_HEADER_ROW_PROBABILITY = 0.75
TABLE_BANDED_ROWS_PROBABILITY = 0.45
TABLE_BANDED_COLS_PROBABILITY = 0.25

# Image crop values use PowerPoint's native picture crop properties.
IMAGE_CROP_PROBABILITY = 0.75
IMAGE_MAX_CROP = 0.22
IMAGE_MASK_OVERLAY_PROBABILITY = 0.35

# python-pptx cannot embed SVG in this environment, but desktop PowerPoint can.
# When enabled, index.py opens the generated PPTX with PowerPoint COM and inserts
# generated SVG assets as real SVG pictures.
INSERT_SVG_WITH_POWERPOINT = True

# Benchmark settings for the fast path: generate many slides into one deck and
# call PowerPoint Export once. This preserves PowerPoint rendering.
DECK_EXPORT_BENCHMARK_SLIDES = 100
