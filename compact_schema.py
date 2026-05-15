import json


TYPE_TO_CODE = {
    "text": "tx",
    "shape": "sh",
    "table": "tb",
    "image": "im",
    "connector": "cn",
    "freeform": "ff",
    "svg": "sv",
    "svg_image": "si",
}

CODE_TO_TYPE = {value: key for key, value in TYPE_TO_CODE.items()}


def _round(value, digits=3):
    try:
        return round(float(value), digits)
    except Exception:
        return value


def _bbox_list(obj):
    box = obj.get("bbox", {}) if isinstance(obj, dict) else {}
    return [_round(box.get(key, 0.0)) for key in ("x", "y", "w", "h")]


def _bbox_dict(values):
    values = list(values or [])
    while len(values) < 4:
        values.append(0.0)
    return {key: _round(values[index], 4) for index, key in enumerate(("x", "y", "w", "h"))}


def _compact_props(kind, props):
    props = props if isinstance(props, dict) else {}
    if kind == "text":
        return {
            "t": props.get("text", ""),
            "fs": props.get("font_size"),
            "ff": props.get("font_face"),
            "c": props.get("color"),
            "b": int(bool(props.get("bold", False))),
            "i": int(bool(props.get("italic", False))),
            "u": int(bool(props.get("underline", False))),
            "bg": props.get("bg_color"),
            "k": props.get("kind"),
        }
    if kind == "shape":
        return {
            "sh": props.get("shape"),
            "f": props.get("fill"),
            "l": props.get("line"),
            "lw": _round(props.get("line_width", 1.0), 2),
            "ol": int(bool(props.get("outline_only", False))),
        }
    if kind == "table":
        return {
            "r": props.get("rows"),
            "c": props.get("cols"),
            "cells": props.get("cells"),
            "hd": props.get("header"),
            "bd": props.get("body"),
            "tc": props.get("text_color"),
            "fs": props.get("font_size"),
        }
    if kind == "image":
        return {"cr": props.get("crop")}
    if kind == "connector":
        return {
            "ct": props.get("connector"),
            "c": props.get("color"),
            "w": _round(props.get("width", 1.0), 2),
        }
    if kind in {"freeform", "svg"}:
        return {
            "p": props.get("pattern"),
            "f": props.get("fill") or props.get("primary"),
            "l": props.get("line") or props.get("secondary"),
            "lw": _round(props.get("line_width", 1.0), 2),
        }
    if kind == "svg_image":
        return {"cr": props.get("crop")}
    return {}


def _expand_props(kind, props):
    props = props if isinstance(props, dict) else {}
    if kind == "text":
        return {
            "kind": props.get("k") or "body",
            "text": str(props.get("t", "")),
            "font_size": int(_round(props.get("fs", 18), 0) or 18),
            "font_face": props.get("ff") or "Aptos",
            "color": props.get("c") or [20, 20, 20],
            "bold": bool(props.get("b", False)),
            "italic": bool(props.get("i", False)),
            "underline": bool(props.get("u", False)),
            "bg_color": props.get("bg"),
            "margin": 0.04,
        }
    if kind == "shape":
        return {
            "family": "basic",
            "shape": props.get("sh") or "rect",
            "fill": props.get("f") or [230, 230, 230],
            "line": props.get("l") or [40, 40, 40],
            "line_width": float(_round(props.get("lw", 1.0), 2) or 1.0),
            "outline_only": bool(props.get("ol", False)),
        }
    if kind == "table":
        rows = max(1, int(_round(props.get("r", 3), 0) or 3))
        cols = max(1, int(_round(props.get("c", 3), 0) or 3))
        cells = props.get("cells") if isinstance(props.get("cells"), list) else []
        return {
            "rows": rows,
            "cols": cols,
            "cells": cells,
            "header": props.get("hd") or [230, 230, 230],
            "body": props.get("bd") or [245, 245, 245],
            "band": props.get("bd") or [235, 235, 235],
            "border": [70, 70, 70],
            "border_width": 1.0,
            "border_dash": "solid",
            "font_size": int(_round(props.get("fs", 10), 0) or 10),
            "text_color": props.get("tc") or [25, 25, 25],
            "has_header": True,
            "band_rows": False,
            "band_cols": False,
        }
    if kind == "image":
        return {"crop": props.get("cr") or [0, 0, 0, 0], "mask_overlay": False, "mask_line": [0, 0, 0]}
    if kind == "connector":
        return {
            "connector": props.get("ct") or "straight",
            "color": props.get("c") or [40, 40, 40],
            "width": float(_round(props.get("w", 1.0), 2) or 1.0),
            "dash": "solid",
            "flip": False,
        }
    if kind == "freeform":
        return {
            "pattern": props.get("p") or "poly",
            "fill": props.get("f") or [230, 230, 230],
            "line": props.get("l") or [40, 40, 40],
            "line_width": float(_round(props.get("lw", 1.0), 2) or 1.0),
        }
    if kind == "svg":
        return {
            "pattern": props.get("p") or "icon",
            "primary": props.get("f") or [230, 230, 230],
            "secondary": props.get("l") or [40, 40, 40],
        }
    if kind == "svg_image":
        return {"crop": props.get("cr") or [0, 0, 0, 0], "mask_overlay": False, "mask_line": [0, 0, 0]}
    return props


def full_to_compact(scene):
    if not isinstance(scene, dict):
        return scene
    slide = scene.get("slide", {}) if isinstance(scene.get("slide"), dict) else {}
    background = scene.get("background") if isinstance(scene.get("background"), dict) else None
    bg_props = background.get("properties", {}) if background else {}
    compact = {
        "v": 1,
        "s": [_round(slide.get("width", 13.333)), _round(slide.get("height", 7.5))],
        "bg": [
            bg_props.get("style", "solid"),
            bg_props.get("fill") or bg_props.get("back_color") or [255, 255, 255],
        ],
        "o": [],
    }
    for obj in scene.get("objects", []) if isinstance(scene.get("objects"), list) else []:
        kind = obj.get("type")
        code = TYPE_TO_CODE.get(kind)
        if not code:
            continue
        compact["o"].append([code, *_bbox_list(obj), _compact_props(kind, obj.get("properties", {}))])
    return compact


def compact_to_full(scene, image_file=None):
    if not isinstance(scene, dict) or "o" not in scene:
        return scene
    slide_size = scene.get("s") if isinstance(scene.get("s"), list) else [13.333, 7.5]
    width = _round(slide_size[0] if len(slide_size) > 0 else 13.333, 4)
    height = _round(slide_size[1] if len(slide_size) > 1 else 7.5, 4)
    bg = scene.get("bg") if isinstance(scene.get("bg"), list) else ["solid", [255, 255, 255]]
    bg_style = bg[0] if len(bg) > 0 else "solid"
    bg_fill = bg[1] if len(bg) > 1 else [255, 255, 255]
    full = {
        "version": 1,
        "slide": {"width": width, "height": height, "image_file": image_file},
        "background": {
            "id": "object_000",
            "type": "background",
            "z_order": 0,
            "bbox": {"x": 0, "y": 0, "w": width, "h": height},
            "properties": {"style": bg_style, "fill": bg_fill},
        },
        "objects": [],
    }
    for index, item in enumerate(scene.get("o", []), start=1):
        if not isinstance(item, list) or len(item) < 5:
            continue
        code = str(item[0]).lower()
        kind = CODE_TO_TYPE.get(code, code)
        if kind not in TYPE_TO_CODE:
            continue
        props = item[5] if len(item) > 5 and isinstance(item[5], dict) else {}
        full["objects"].append(
            {
                "id": f"object_{index:03}",
                "type": kind,
                "z_order": index,
                "bbox": _bbox_dict(item[1:5]),
                "properties": _expand_props(kind, props),
            }
        )
    return full


def compact_json_len(target_path):
    with open(target_path, "r", encoding="utf-8") as handle:
        return len(json.dumps(full_to_compact(json.load(handle)), separators=(",", ":")))


def compact_schema_reward(scene):
    if not isinstance(scene, dict):
        return -1.0
    points = 0.0
    total = 6.0
    points += scene.get("v") == 1
    points += isinstance(scene.get("s"), list) and len(scene["s"]) >= 2
    points += isinstance(scene.get("bg"), list) and len(scene["bg"]) >= 2
    objects = scene.get("o")
    points += isinstance(objects, list)
    if isinstance(objects, list):
        valid = 0
        for item in objects:
            valid += (
                isinstance(item, list)
                and len(item) >= 6
                and item[0] in CODE_TO_TYPE
                and all(isinstance(value, (int, float)) for value in item[1:5])
                and isinstance(item[5], dict)
            )
        points += valid / max(1, len(objects))
        points += 1.0 if 1 <= len(objects) <= 40 else 0.0
    return max(-1.0, min(1.0, points / total * 2.0 - 1.0))
