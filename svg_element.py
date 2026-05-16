import base64
import math
import random
import tempfile
from pathlib import Path

from element_model import Element
from utils import rand_color


SVG_PATTERNS = [
    "organic_blob",
    "layered_wave",
    "radial_burst",
    "loop_rings",
    "corner_ribbon",
    "dotted_cluster",
    "contour_lines",
    "bracket_frame",
]
SVG_IMAGE_CLIPS = ["blob", "ticket", "arch", "hex"]


def svg_color(rgb):
    values = []

    def collect(value):
        if isinstance(value, (list, tuple)):
            for item in value:
                collect(item)
        else:
            values.append(value)

    collect(rgb)
    return "#%02x%02x%02x" % tuple(int(v) for v in values[:3])


def make_svg(x, y, w, h):
    pattern = random.choice(SVG_PATTERNS)
    primary = rand_color()
    secondary = rand_color()
    params = svg_params(pattern)
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
            "line_width": params.get("line_width", 1.0),
            "params": params,
            "svg": svg_markup(pattern, primary, secondary, params),
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


def svg_params(pattern):
    if pattern == "organic_blob":
        return {
            "bulge": random.uniform(0.10, 0.24),
            "pinch": random.uniform(0.08, 0.20),
            "tilt": random.uniform(-0.08, 0.08),
            "line_width": random.uniform(5, 18),
        }
    if pattern == "layered_wave":
        return {
            "waves": random.randint(2, 4),
            "amplitude": random.uniform(24, 58),
            "phase": random.uniform(-35, 35),
            "line_width": random.uniform(10, 24),
        }
    if pattern == "radial_burst":
        return {
            "points": random.choice([10, 12, 14, 16, 18]),
            "inner": random.uniform(52, 92),
            "outer": random.uniform(118, 145),
            "line_width": random.uniform(4, 11),
        }
    if pattern == "loop_rings":
        return {
            "rings": random.randint(2, 4),
            "offset": random.uniform(28, 52),
            "line_width": random.uniform(10, 24),
        }
    if pattern == "corner_ribbon":
        return {
            "fold": random.uniform(42, 82),
            "depth": random.uniform(42, 92),
            "line_width": random.uniform(3, 9),
        }
    if pattern == "dotted_cluster":
        return {
            "dots": random.randint(9, 18),
            "spread": random.uniform(0.52, 0.78),
            "line_width": random.uniform(1, 4),
        }
    if pattern == "contour_lines":
        return {
            "lines": random.randint(3, 6),
            "amplitude": random.uniform(16, 42),
            "line_width": random.uniform(4, 10),
        }
    return {
        "inset": random.uniform(26, 54),
        "notch": random.uniform(22, 48),
        "line_width": random.uniform(5, 14),
    }


def _blob_path(cx, cy, rx, ry, bulge, pinch, tilt):
    left = cx - rx
    right = cx + rx
    top = cy - ry
    bottom = cy + ry
    return (
        f"M{left + rx * 0.14:.1f} {cy - ry * (0.20 + tilt):.1f} "
        f"C{left + rx * bulge:.1f} {top + ry * 0.08:.1f} {cx - rx * pinch:.1f} {top - ry * 0.02:.1f} {cx + rx * 0.22:.1f} {top + ry * 0.10:.1f} "
        f"C{right + rx * 0.12:.1f} {top + ry * 0.28:.1f} {right - rx * 0.02:.1f} {cy + ry * 0.18:.1f} {right - rx * 0.16:.1f} {cy + ry * 0.40:.1f} "
        f"C{right - rx * 0.32:.1f} {bottom + ry * 0.04:.1f} {cx + rx * 0.08:.1f} {bottom - ry * 0.02:.1f} {cx - rx * 0.30:.1f} {bottom - ry * 0.10:.1f} "
        f"C{left - rx * 0.06:.1f} {bottom - ry * 0.26:.1f} {left + rx * 0.04:.1f} {cy + ry * 0.08:.1f} {left + rx * 0.14:.1f} {cy - ry * (0.20 + tilt):.1f} Z"
    )


def _star_points(cx, cy, inner, outer, count):
    pts = []
    for index in range(count * 2):
        angle = -1.5708 + index * 3.14159 / count
        radius = outer if index % 2 == 0 else inner
        pts.append((cx + radius * math.cos(angle), cy + radius * math.sin(angle)))
    return " ".join(f"{x:.1f},{y:.1f}" for x, y in pts)


def svg_markup(pattern, primary, secondary, params=None):
    params = params or svg_params(pattern)
    p = svg_color(primary)
    s = svg_color(secondary)
    lw = params.get("line_width", 1.0)
    if pattern == "organic_blob":
        path1 = _blob_path(190, 150, 145, 92, params["bulge"], params["pinch"], params["tilt"])
        path2 = _blob_path(220, 144, 92, 54, params["pinch"], params["bulge"], -params["tilt"])
        body = f"""
        <path d="{path1}" fill="{p}" opacity="0.78"/>
        <path d="{path2}" fill="{s}" opacity="0.46"/>
        <path d="{path1}" fill="none" stroke="{s}" stroke-width="{lw:.1f}" opacity="0.68"/>
        """
    elif pattern == "radial_burst":
        points = _star_points(200, 150, params["inner"], params["outer"], params["points"])
        body = f"""
        <polygon points="{points}" fill="{p}" opacity="0.72"/>
        <circle cx="200" cy="150" r="{params['inner'] * 0.52:.1f}" fill="{s}" opacity="0.64"/>
        <polygon points="{points}" fill="none" stroke="{s}" stroke-width="{lw:.1f}" opacity="0.62"/>
        """
    elif pattern == "layered_wave":
        lines = []
        for index in range(params["waves"]):
            y = 72 + index * (160 / max(1, params["waves"] - 1))
            amp = params["amplitude"] * (0.72 + index * 0.12)
            phase = params["phase"] + index * 18
            color = p if index % 2 == 0 else s
            opacity = 0.82 - index * 0.08
            lines.append(
                f'<path d="M24 {y:.1f} C96 {y - amp + phase:.1f} 164 {y + amp:.1f} 234 {y:.1f} S330 {y - amp:.1f} 376 {y + phase:.1f}" '
                f'fill="none" stroke="{color}" stroke-width="{lw:.1f}" stroke-linecap="round" opacity="{opacity:.2f}"/>'
            )
        body = "\n        ".join(lines)
    elif pattern == "loop_rings":
        rings = []
        for index in range(params["rings"]):
            cx = 150 + index * params["offset"]
            cy = 130 + (index % 2) * 34
            rx = 72 - index * 7
            ry = 58 - index * 4
            color = p if index % 2 == 0 else s
            rings.append(
                f'<ellipse cx="{cx:.1f}" cy="{cy:.1f}" rx="{rx:.1f}" ry="{ry:.1f}" fill="none" stroke="{color}" stroke-width="{lw:.1f}" opacity="{0.86 - index * 0.09:.2f}"/>'
            )
        body = "\n        ".join(rings)
    elif pattern == "corner_ribbon":
        fold = params["fold"]
        depth = params["depth"]
        body = f"""
        <path d="M28 28 H372 V{depth:.1f} C292 {depth + fold:.1f} 226 {depth - fold * 0.28:.1f} 154 {depth + fold * 0.46:.1f} C98 {depth + fold * 0.82:.1f} 58 {depth + fold * 0.18:.1f} 28 {depth + fold * 0.52:.1f} Z" fill="{p}" opacity="0.76"/>
        <path d="M372 28 L{372 - fold:.1f} {28 + fold:.1f} H372 Z" fill="{s}" opacity="0.68"/>
        <path d="M44 {depth + fold * 0.36:.1f} C122 {depth - fold * 0.12:.1f} 178 {depth + fold * 0.76:.1f} 254 {depth + fold * 0.26:.1f} S344 {depth + fold * 0.08:.1f} 374 {depth + fold * 0.5:.1f}" fill="none" stroke="{s}" stroke-width="{lw:.1f}" stroke-linecap="round"/>
        """
    elif pattern == "dotted_cluster":
        dots = []
        spread = params["spread"]
        for index in range(params["dots"]):
            col = index % 6
            row = index // 6
            x = 92 + col * 42 * spread + random.uniform(-8, 8)
            y = 70 + row * 54 * spread + random.uniform(-8, 8)
            r = random.uniform(8, 22)
            color = p if index % 3 else s
            dots.append(f'<circle cx="{x:.1f}" cy="{y:.1f}" r="{r:.1f}" fill="{color}" opacity="{random.uniform(0.48, 0.86):.2f}"/>')
        body = "\n        ".join(dots)
    elif pattern == "contour_lines":
        lines = []
        for index in range(params["lines"]):
            inset = 28 + index * 22
            amp = params["amplitude"] - index * 2
            color = p if index % 2 == 0 else s
            lines.append(
                f'<path d="M{inset:.1f} {158 - index * 11:.1f} C{92 + index * 10:.1f} {86 - amp:.1f} {176 + index * 7:.1f} {222 + amp:.1f} {252 - index * 4:.1f} {144 + index * 8:.1f} S{344 - index * 12:.1f} {84 + amp:.1f} {376 - inset * 0.28:.1f} {154 + index * 10:.1f}" fill="none" stroke="{color}" stroke-width="{lw:.1f}" stroke-linecap="round" opacity="{0.82 - index * 0.08:.2f}"/>'
            )
        body = "\n        ".join(lines)
    else:
        inset = params["inset"]
        notch = params["notch"]
        body = f"""
        <path d="M{inset:.1f} {inset:.1f} H{400 - inset:.1f} V{inset + notch:.1f} H{400 - inset - notch:.1f} V{300 - inset - notch:.1f} H{400 - inset:.1f} V{300 - inset:.1f} H{inset:.1f} V{300 - inset - notch:.1f} H{inset + notch:.1f} V{inset + notch:.1f} H{inset:.1f} Z" fill="{p}" opacity="0.34"/>
        <path d="M{inset:.1f} {inset:.1f} H{400 - inset:.1f} M{400 - inset:.1f} {300 - inset:.1f} H{inset:.1f}" fill="none" stroke="{s}" stroke-width="{lw:.1f}" stroke-linecap="round"/>
        <path d="M{inset:.1f} {300 - inset:.1f} V{inset:.1f} M{400 - inset:.1f} {inset:.1f} V{300 - inset:.1f}" fill="none" stroke="{s}" stroke-width="{lw * 0.72:.1f}" stroke-linecap="round" opacity="0.72"/>
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
    if el.data["pattern"] == "loop_rings":
        draw.ellipse((x1, y1, x1 + (x2 - x1) * 0.72, y2), outline=p, width=12)
        draw.ellipse((x1 + (x2 - x1) * 0.28, y1 + 8, x2, y2 - 8), outline=s, width=10)
    elif el.data["pattern"] == "radial_burst":
        cx, cy = (x1 + x2) // 2, (y1 + y2) // 2
        pts = [(cx, y1), (x2, cy), (cx, y2), (x1, cy)]
        draw.polygon(pts, fill=p, outline=s)
    elif el.data["pattern"] in {"layered_wave", "contour_lines"}:
        for offset in (0.32, 0.50, 0.68):
            y = int(y1 + (y2 - y1) * offset)
            draw.arc((x1, y - 28, x2, y + 28), 180, 360, fill=p, width=8)
    elif el.data["pattern"] == "dotted_cluster":
        for index in range(12):
            cx = int(x1 + (x2 - x1) * (0.18 + (index % 4) * 0.2))
            cy = int(y1 + (y2 - y1) * (0.22 + (index // 4) * 0.24))
            r = max(3, int(min(x2 - x1, y2 - y1) * 0.055))
            draw.ellipse((cx - r, cy - r, cx + r, cy + r), fill=p if index % 2 else s)
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
