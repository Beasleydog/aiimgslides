import base64
import random
import tempfile
from pathlib import Path

from element_model import Element
from utils import rand_color


SVG_PATTERNS = ["ribbon_path", "rings", "burst", "swoop"]
SVG_IMAGE_CLIPS = ["blob", "ticket", "arch", "hex"]


def svg_color(rgb):
    return "#%02x%02x%02x" % rgb[:3]


def make_svg(x, y, w, h):
    pattern = random.choice(SVG_PATTERNS)
    primary = rand_color()
    secondary = rand_color()
    return Element(
        "svg",
        x,
        y,
        w,
        h,
        {
            "pattern": pattern,
            "primary": primary,
            "secondary": secondary,
            "svg": svg_markup(pattern, primary, secondary),
        },
    )


def make_svg_image(x, y, w, h):
    return Element(
        "svg_image",
        x,
        y,
        w,
        h,
        {
            "clip": random.choice(SVG_IMAGE_CLIPS),
            "line": rand_color(),
            "line_width": random.randint(8, 18),
        },
    )


def svg_markup(pattern, primary, secondary):
    p = svg_color(primary)
    s = svg_color(secondary)
    if pattern == "rings":
        body = f"""
        <circle cx="145" cy="110" r="74" fill="none" stroke="{p}" stroke-width="22"/>
        <circle cx="245" cy="118" r="62" fill="none" stroke="{s}" stroke-width="18" opacity="0.82"/>
        <circle cx="196" cy="72" r="28" fill="{p}" opacity="0.6"/>
        """
    elif pattern == "burst":
        body = f"""
        <path d="M200 10 L226 86 L306 52 L268 126 L386 132 L278 168 L332 240 L236 194 L200 290 L166 196 L70 240 L124 168 L14 132 L132 126 L94 52 L174 86 Z"
              fill="{p}" opacity="0.78"/>
        <path d="M200 82 L226 145 L294 146 L238 184 L258 252 L200 212 L142 252 L162 184 L106 146 L174 145 Z"
              fill="{s}" opacity="0.72"/>
        """
    elif pattern == "swoop":
        body = f"""
        <path d="M22 196 C96 26 198 270 286 70 S372 92 386 204"
              fill="none" stroke="{p}" stroke-width="28" stroke-linecap="round"/>
        <path d="M42 226 C124 112 214 288 360 126"
              fill="none" stroke="{s}" stroke-width="12" stroke-linecap="round" opacity="0.85"/>
        """
    else:
        body = f"""
        <path d="M24 84 C94 16 172 18 224 78 C278 142 334 96 376 42 L376 210 C304 260 230 244 178 190 C116 128 74 168 24 226 Z"
              fill="{p}" opacity="0.82"/>
        <path d="M54 118 C120 54 178 66 226 118 C270 166 318 142 354 104"
              fill="none" stroke="{s}" stroke-width="18" stroke-linecap="round"/>
        """
    return f'<svg xmlns="http://www.w3.org/2000/svg" width="400" height="300" viewBox="0 0 400 300">{body}</svg>'


def clip_path(clip):
    if clip == "ticket":
        return "M36 38 H364 Q342 86 364 134 Q342 182 364 262 H36 Q58 214 36 166 Q58 118 36 38 Z"
    if clip == "arch":
        return "M42 276 V132 C42 50 102 18 200 18 C298 18 358 50 358 132 V276 Z"
    if clip == "hex":
        return "M98 24 H302 L382 150 L302 276 H98 L18 150 Z"
    return "M40 102 C82 26 146 16 200 58 C266 110 330 48 368 118 C406 188 318 286 214 260 C126 238 82 298 38 230 C6 180 12 140 40 102 Z"


def svg_image_markup(el, image_path=None, encoded_image=None):
    encoded = encoded_image or base64.b64encode(Path(image_path).read_bytes()).decode("ascii")
    path = clip_path(el.data["clip"])
    line = svg_color(el.data["line"])
    width = el.data["line_width"]
    return f"""
    <svg xmlns="http://www.w3.org/2000/svg" width="400" height="300" viewBox="0 0 400 300">
      <defs>
        <clipPath id="imgClip"><path d="{path}"/></clipPath>
      </defs>
      <image width="400" height="300" preserveAspectRatio="xMidYMid slice"
             href="data:image/jpeg;base64,{encoded}" clip-path="url(#imgClip)"/>
      <path d="{path}" fill="none" stroke="{line}" stroke-width="{width}" stroke-linejoin="round"/>
    </svg>
    """


def add_svg_to_png(draw, el, box):
    x1, y1, x2, y2 = box
    p = (*el.data["primary"], 165)
    s = (*el.data["secondary"], 210)
    if el.data["pattern"] == "rings":
        draw.ellipse((x1, y1, x1 + (x2 - x1) * 0.72, y2), outline=p, width=12)
        draw.ellipse((x1 + (x2 - x1) * 0.28, y1 + 8, x2, y2 - 8), outline=s, width=10)
    elif el.data["pattern"] == "burst":
        cx, cy = (x1 + x2) // 2, (y1 + y2) // 2
        pts = [(cx, y1), (x2, cy), (cx, y2), (x1, cy)]
        draw.polygon(pts, fill=p, outline=s)
    else:
        draw.rounded_rectangle(box, radius=18, fill=p, outline=s, width=5)


def add_svg_image_to_png(draw, el, box):
    x1, y1, x2, y2 = box
    line = (*el.data["line"], 255)
    if el.data["clip"] == "hex":
        pts = [
            (x1 + (x2 - x1) * 0.22, y1),
            (x1 + (x2 - x1) * 0.78, y1),
            (x2, (y1 + y2) // 2),
            (x1 + (x2 - x1) * 0.78, y2),
            (x1 + (x2 - x1) * 0.22, y2),
            (x1, (y1 + y2) // 2),
        ]
        draw.polygon(pts, fill=(210, 210, 210, 120), outline=line)
    else:
        draw.rounded_rectangle(box, radius=24, fill=(210, 210, 210, 120), outline=line, width=4)


def insert_svg_elements(pptx_path, slide_elements, image_path=None):
    import win32com.client

    app = win32com.client.Dispatch("PowerPoint.Application")
    try:
        insert_svg_elements_with_app(app, pptx_path, slide_elements, image_path)
    finally:
        app.Quit()


def insert_svg_elements_with_app(app, pptx_path, slide_elements, image_path=None):
    pptx_path = Path(pptx_path).resolve()
    presentation = None
    with tempfile.TemporaryDirectory() as tmp:
        try:
            presentation = app.Presentations.Open(str(pptx_path), WithWindow=False)
            insert_svg_elements_into_presentation(presentation, slide_elements, image_path, Path(tmp))
            presentation.Save()
        finally:
            if presentation is not None:
                presentation.Close()


def insert_svg_elements_into_presentation(presentation, slide_elements, image_path=None, tmp_path=None):
    tmp_path = Path(tmp_path) if tmp_path is not None else Path(tempfile.mkdtemp())
    needs_encoded_image = image_path is not None and any(
        el.kind == "svg_image"
        for elements in slide_elements
        for el in elements
    )
    encoded_image = base64.b64encode(Path(image_path).read_bytes()).decode("ascii") if needs_encoded_image else None

    for slide_index, elements in enumerate(slide_elements, start=1):
        slide = presentation.Slides(slide_index)
        svg_elements = [e for e in elements if e.kind in {"svg", "svg_image"}]
        for svg_index, el in enumerate(svg_elements, start=1):
            svg_path = tmp_path / f"slide_{slide_index}_svg_{svg_index}.svg"
            if el.kind == "svg_image":
                if image_path is None:
                    continue
                svg_path.write_text(svg_image_markup(el, encoded_image=encoded_image), encoding="utf-8")
            else:
                svg_path.write_text(el.data["svg"], encoding="utf-8")
            slide.Shapes.AddPicture(
                str(svg_path.resolve()),
                False,
                True,
                el.x * 72,
                el.y * 72,
                el.w * 72,
                el.h * 72,
            )
