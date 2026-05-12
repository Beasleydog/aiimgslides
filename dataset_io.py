import json
from pathlib import Path

import config


def json_safe(value):
    if value is None or isinstance(value, (bool, int, float, str)):
        return value
    if isinstance(value, Path):
        return str(value)
    if hasattr(value, "name") and hasattr(value, "value"):
        return {"name": value.name, "value": value.value}
    if isinstance(value, tuple):
        return [json_safe(item) for item in value]
    if isinstance(value, list):
        return [json_safe(item) for item in value]
    if isinstance(value, dict):
        return {str(key): json_safe(item) for key, item in value.items()}
    return str(value)


def element_to_dict(el, index):
    return {
        "id": f"object_{index:03}",
        "type": el.kind,
        "z_order": index,
        "bbox": {
            "x": round(el.x, 4),
            "y": round(el.y, 4),
            "w": round(el.w, 4),
            "h": round(el.h, 4),
        },
        "properties": json_safe(el.data),
    }


def slide_to_dict(elements, image_file=None):
    background = next((el for el in elements if el.kind == "background"), None)
    objects = [el for el in elements if el.kind != "background"]
    return {
        "version": 1,
        "slide": {
            "width": config.SLIDE_W,
            "height": config.SLIDE_H,
            "image_file": image_file,
        },
        "background": element_to_dict(background, 0) if background is not None else None,
        "objects": [element_to_dict(el, index) for index, el in enumerate(objects, start=1)],
    }


def save_slide_json(path, elements, image_file=None):
    data = slide_to_dict(elements, image_file=image_file)
    Path(path).write_text(json.dumps(data, indent=2, sort_keys=True), encoding="utf-8")
