import json
import math
from collections import Counter
from difflib import SequenceMatcher
from pathlib import Path


IMAGE_TYPES = {"image", "svg_image"}
LIGHT_PROPERTY_TYPES = IMAGE_TYPES
VALID_OBJECT_TYPES = {
    "text",
    "shape",
    "table",
    "image",
    "connector",
    "freeform",
    "svg",
    "svg_image",
}

IGNORED_PROPERTY_KEYS = {
    "archetype",
    "focus",
    "ppt_type",
    "svg",
    "theme_topic",
    "url",
}

PROPERTY_WEIGHTS = {
    "text": {
        "text": 2.0,
        "kind": 0.7,
        "font_size": 1.0,
        "font_face": 0.35,
        "color": 1.0,
        "bold": 0.45,
        "italic": 0.25,
        "underline": 0.25,
        "bg_color": 0.55,
        "margin": 0.15,
    },
    "shape": {
        "family": 0.6,
        "shape": 1.2,
        "fill": 1.0,
        "line": 1.0,
        "line_width": 0.7,
        "outline_only": 0.45,
        "dash": 0.45,
    },
    "connector": {
        "connector": 0.9,
        "color": 1.0,
        "width": 0.8,
        "dash": 0.45,
        "flip": 0.35,
    },
    "table": {
        "rows": 1.0,
        "cols": 1.0,
        "header": 0.75,
        "body": 0.75,
        "band": 0.45,
        "border": 0.75,
        "border_width": 0.5,
        "border_dash": 0.3,
        "font_size": 0.5,
        "text_color": 0.6,
        "has_header": 0.35,
        "band_rows": 0.25,
        "band_cols": 0.25,
        "cells": 0.5,
    },
    "freeform": {
        "pattern": 1.0,
        "fill": 1.0,
        "line": 1.0,
        "line_width": 0.7,
    },
    "svg": {
        "pattern": 1.0,
        "primary": 1.0,
        "secondary": 1.0,
    },
    "background": {
        "style": 1.5,
        "fill": 1.2,
        "gradient_type": 0.7,
        "colors": 1.2,
        "angle": 0.45,
        "focus": 0.35,
        "pattern": 0.7,
        "fore_color": 1.0,
        "back_color": 1.0,
    },
}


def load_json(value):
    if isinstance(value, (str, Path)):
        return json.loads(Path(value).read_text(encoding="utf-8"))
    return value


def clamp(value, lo=-1.0, hi=1.0):
    return max(lo, min(hi, value))


def unit_to_reward(score):
    return clamp(score, 0.0, 1.0) * 2.0 - 1.0


def enum_name(value):
    if isinstance(value, dict):
        return value.get("name", value.get("value"))
    return value


def as_number(value):
    return value if isinstance(value, (int, float)) and not isinstance(value, bool) else None


def is_color(value):
    return (
        isinstance(value, list)
        and len(value) in {3, 4}
        and all(isinstance(item, (int, float)) for item in value[:3])
        and all(0 <= item <= 255 for item in value[:3])
    )


def numeric_reward(target, actual, tolerance=None):
    target_num = as_number(target)
    actual_num = as_number(actual)
    if target_num is None or actual_num is None:
        return -0.6
    if tolerance is None:
        tolerance = max(abs(target_num) * 0.35, 1.0)
    error = abs(target_num - actual_num)
    return unit_to_reward(math.exp(-error / max(tolerance, 1e-6)))


def color_reward(target, actual):
    if not is_color(target) or not is_color(actual):
        return -0.6
    dist = math.sqrt(sum((float(a) - float(b)) ** 2 for a, b in zip(target[:3], actual[:3])))
    max_dist = math.sqrt(3 * 255**2)
    return unit_to_reward(math.exp(-3.0 * dist / max_dist))


def token_f1(target, actual):
    target_tokens = str(target).lower().split()
    actual_tokens = str(actual).lower().split()
    if not target_tokens and not actual_tokens:
        return 1.0
    if not target_tokens or not actual_tokens:
        return 0.0
    target_counts = Counter(target_tokens)
    actual_counts = Counter(actual_tokens)
    overlap = sum((target_counts & actual_counts).values())
    precision = overlap / max(1, len(actual_tokens))
    recall = overlap / max(1, len(target_tokens))
    if precision + recall == 0:
        return 0.0
    return 2 * precision * recall / (precision + recall)


def text_reward(target, actual):
    if not isinstance(actual, str):
        return -0.6
    seq = SequenceMatcher(None, target, actual).ratio()
    tok = token_f1(target, actual)
    length = math.exp(-abs(len(actual) - len(target)) / max(8, len(target)))
    return unit_to_reward(0.45 * seq + 0.45 * tok + 0.10 * length)


def categorical_reward(target, actual):
    target = enum_name(target)
    actual = enum_name(actual)
    if target == actual:
        return 1.0
    if str(target).lower() == str(actual).lower():
        return 0.85
    return -0.45


def list_reward(target, actual, key=""):
    if not isinstance(actual, list):
        return -0.6
    if is_color(target):
        return color_reward(target, actual)
    if not target:
        return 1.0 if not actual else -0.25
    scores = []
    for index, target_item in enumerate(target):
        actual_item = actual[index] if index < len(actual) else None
        scores.append(value_reward(target_item, actual_item, key=key))
    extra = max(0, len(actual) - len(target))
    return clamp(sum(scores) / len(scores) - min(0.35, extra * 0.05))


def value_reward(target, actual, key=""):
    if actual is None:
        return -0.6
    if is_color(target):
        return color_reward(target, actual)
    if isinstance(target, bool):
        return categorical_reward(target, actual)
    if isinstance(target, (int, float)) and not isinstance(target, bool):
        return numeric_reward(target, actual)
    if isinstance(target, str):
        if key == "text" or len(target.split()) > 3:
            return text_reward(target, actual)
        return categorical_reward(target, actual)
    if isinstance(target, list):
        return list_reward(target, actual, key=key)
    if isinstance(target, dict):
        if "name" in target or "value" in target:
            return categorical_reward(target, actual)
        if not isinstance(actual, dict):
            return -0.6
        return property_reward("__dict__", target, actual)
    return categorical_reward(target, actual)


def property_reward(kind, target_props, actual_props):
    if not isinstance(target_props, dict):
        return 0.0
    if not isinstance(actual_props, dict):
        return -0.65

    weights = PROPERTY_WEIGHTS.get(kind, {})
    keys = [
        key
        for key in target_props
        if key not in IGNORED_PROPERTY_KEYS and (not weights or key in weights)
    ]
    if not keys:
        return 0.0

    weighted_scores = []
    total_weight = 0.0
    for key in keys:
        weight = weights.get(key, 0.35)
        weighted_scores.append(weight * value_reward(target_props[key], actual_props.get(key), key=key))
        total_weight += weight

    missing_keys = [key for key in keys if key not in actual_props]
    extra_keys = [
        key
        for key in actual_props
        if key not in IGNORED_PROPERTY_KEYS and key not in target_props
    ]
    penalty = min(0.35, len(missing_keys) * 0.035 + len(extra_keys) * 0.015)
    return clamp(sum(weighted_scores) / max(total_weight, 1e-6) - penalty)


def scorable_property_keys(kind, target_props):
    if not isinstance(target_props, dict):
        return []
    weights = PROPERTY_WEIGHTS.get(kind, {})
    return [
        key
        for key in target_props
        if key not in IGNORED_PROPERTY_KEYS and (not weights or key in weights)
    ]


def property_recall_reward(matches, target_objects, actual_objects):
    present_weight = 0.0
    total_weight = 0.0
    matched_any = False
    for match in matches:
        actual_index = match.get("input_index")
        if actual_index is None:
            continue
        matched_any = True
        target_obj = target_objects[match["target_index"]]
        kind = target_obj.get("type")
        if kind in LIGHT_PROPERTY_TYPES:
            continue
        target_props = target_obj.get("properties", {})
        actual_props = actual_objects[actual_index].get("properties", {})
        weights = PROPERTY_WEIGHTS.get(kind, {})
        for key in scorable_property_keys(kind, target_props):
            weight = weights.get(key, 0.35)
            total_weight += weight
            if isinstance(actual_props, dict) and key in actual_props:
                present_weight += weight
    if not matched_any:
        has_target_properties = any(
            obj.get("type") not in LIGHT_PROPERTY_TYPES
            and scorable_property_keys(obj.get("type"), obj.get("properties", {}))
            for obj in target_objects
        )
        if has_target_properties:
            return -1.0
    if total_weight == 0:
        return 1.0
    return unit_to_reward(present_weight / total_weight)


def bbox_tuple(bbox):
    return (
        float(bbox.get("x", 0.0)),
        float(bbox.get("y", 0.0)),
        float(bbox.get("w", 0.0)),
        float(bbox.get("h", 0.0)),
    )


def valid_bbox(bbox):
    try:
        x, y, w, h = bbox_tuple(bbox)
    except Exception:
        return False
    return all(math.isfinite(v) for v in (x, y, w, h)) and w > 0 and h > 0


def bbox_area(bbox):
    if not valid_bbox(bbox):
        return 0.0
    _, _, w, h = bbox_tuple(bbox)
    return w * h


def bbox_iou(a, b):
    ax, ay, aw, ah = bbox_tuple(a)
    bx, by, bw, bh = bbox_tuple(b)
    if aw <= 0 or ah <= 0 or bw <= 0 or bh <= 0:
        return 0.0
    left = max(ax, bx)
    top = max(ay, by)
    right = min(ax + aw, bx + bw)
    bottom = min(ay + ah, by + bh)
    inter = max(0.0, right - left) * max(0.0, bottom - top)
    union = aw * ah + bw * bh - inter
    return inter / union if union > 0 else 0.0


def bbox_reward(target_bbox, actual_bbox, slide):
    width = float(slide.get("width", 13.333) or 13.333)
    height = float(slide.get("height", 7.5) or 7.5)
    tx, ty, tw, th = bbox_tuple(target_bbox)
    ax, ay, aw, ah = bbox_tuple(actual_bbox)

    center_dist = math.hypot((tx + tw / 2) - (ax + aw / 2), (ty + th / 2) - (ay + ah / 2))
    center_score = math.exp(-center_dist / max(0.01, 0.12 * math.hypot(width, height)))
    size_dist = abs(tw - aw) / max(width, 0.01) + abs(th - ah) / max(height, 0.01)
    size_score = math.exp(-4.0 * size_dist)
    iou_score = bbox_iou(target_bbox, actual_bbox)
    return unit_to_reward(0.45 * iou_score + 0.35 * center_score + 0.20 * size_score)


def type_reward(target_type, actual_type):
    if target_type == actual_type:
        return 1.0
    if target_type in IMAGE_TYPES and actual_type in IMAGE_TYPES:
        return 0.65
    return -0.45


def schema_reward(data):
    if not isinstance(data, dict):
        return -1.0
    points = 0.0
    total = 7.0
    points += isinstance(data.get("slide"), dict)
    points += data.get("version") == 1
    slide = data.get("slide", {}) if isinstance(data.get("slide"), dict) else {}
    points += as_number(slide.get("width")) is not None
    points += as_number(slide.get("height")) is not None
    points += data.get("background") is None or isinstance(data.get("background"), dict)
    objects = data.get("objects")
    points += isinstance(objects, list)
    if isinstance(objects, list):
        valid = 0
        for obj in objects:
            valid += (
                isinstance(obj, dict)
                and isinstance(obj.get("bbox"), dict)
                and valid_bbox(obj.get("bbox", {}))
                and isinstance(obj.get("properties"), dict)
                and obj.get("type") in VALID_OBJECT_TYPES
            )
        points += valid / max(1, len(objects))
    return unit_to_reward(points / total)


def type_distribution_reward(target_objects, actual_objects):
    target_counts = Counter(obj.get("type") for obj in target_objects)
    actual_counts = Counter(obj.get("type") for obj in actual_objects)
    if not target_counts and not actual_counts:
        return 1.0
    overlap = sum(min(target_counts[key], actual_counts[key]) for key in target_counts | actual_counts)
    total = max(sum(target_counts.values()), sum(actual_counts.values()), 1)
    return unit_to_reward(overlap / total)


def count_reward(target_count, actual_count):
    diff = abs(target_count - actual_count)
    scale = max(target_count, actual_count, 1)
    return unit_to_reward(math.exp(-2.0 * diff / scale))


def object_pair_reward(target_obj, actual_obj, slide):
    kind = target_obj.get("type")
    t_reward = type_reward(kind, actual_obj.get("type"))
    b_reward = bbox_reward(target_obj.get("bbox", {}), actual_obj.get("bbox", {}), slide)

    if kind in LIGHT_PROPERTY_TYPES:
        p_reward = 0.0
        weights = (0.35, 0.65, 0.0)
    else:
        p_reward = property_reward(kind, target_obj.get("properties", {}), actual_obj.get("properties", {}))
        weights = (0.22, 0.43, 0.35)

    reward = weights[0] * t_reward + weights[1] * b_reward + weights[2] * p_reward
    return {
        "reward": clamp(reward),
        "type": t_reward,
        "bbox": b_reward,
        "properties": p_reward,
    }


def match_affinity(target_obj, actual_obj, slide):
    return 0.55 * type_reward(target_obj.get("type"), actual_obj.get("type")) + 0.45 * bbox_reward(target_obj.get("bbox", {}), actual_obj.get("bbox", {}), slide)


def match_objects(target_objects, actual_objects, slide):
    candidates = []
    for target_index, target_obj in enumerate(target_objects):
        for actual_index, actual_obj in enumerate(actual_objects):
            candidates.append((match_affinity(target_obj, actual_obj, slide), target_index, actual_index))
    candidates.sort(reverse=True)

    used_targets = set()
    used_actuals = set()
    matches = []
    for affinity, target_index, actual_index in candidates:
        if target_index in used_targets or actual_index in used_actuals:
            continue
        if affinity < -0.35:
            continue
        used_targets.add(target_index)
        used_actuals.add(actual_index)
        matches.append(
            {
                "target_index": target_index,
                "input_index": actual_index,
                "scores": object_pair_reward(target_objects[target_index], actual_objects[actual_index], slide),
            }
        )

    for target_index in range(len(target_objects)):
        if target_index not in used_targets:
            matches.append({"target_index": target_index, "input_index": None, "scores": {"reward": -0.75}})

    matches.sort(key=lambda item: item["target_index"])
    extra = [index for index in range(len(actual_objects)) if index not in used_actuals]
    return matches, extra


def object_precision_reward(matches, actual_count):
    matched = len([match for match in matches if match["input_index"] is not None])
    return unit_to_reward(matched / max(1, actual_count))


def object_recall_reward(matches, target_count):
    matched = len([match for match in matches if match["input_index"] is not None])
    return unit_to_reward(matched / max(1, target_count))


def duplicate_penalty(actual_objects):
    duplicates = 0
    for i, left in enumerate(actual_objects):
        for right in actual_objects[i + 1 :]:
            if left.get("type") == right.get("type") and bbox_iou(left.get("bbox", {}), right.get("bbox", {})) > 0.88:
                duplicates += 1
    return min(0.35, duplicates * 0.04)


def huge_box_penalty(actual_objects, target_objects, slide):
    slide_area = float(slide.get("width", 13.333) or 13.333) * float(slide.get("height", 7.5) or 7.5)
    target_huge_by_type = Counter(
        obj.get("type")
        for obj in target_objects
        if bbox_area(obj.get("bbox", {})) > slide_area * 0.72
    )
    actual_huge_by_type = Counter(
        obj.get("type")
        for obj in actual_objects
        if bbox_area(obj.get("bbox", {})) > slide_area * 0.72
    )
    excess = sum(max(0, actual_huge_by_type[key] - target_huge_by_type.get(key, 0)) for key in actual_huge_by_type)
    return min(0.30, excess * 0.08)


def invalid_bbox_penalty(actual_objects, slide):
    width = float(slide.get("width", 13.333) or 13.333)
    height = float(slide.get("height", 7.5) or 7.5)
    bad = 0
    for obj in actual_objects:
        if not valid_bbox(obj.get("bbox", {})):
            bad += 1
            continue
        x, y, w, h = bbox_tuple(obj.get("bbox", {}))
        if x < -2.5 or y < -2.5 or x + w > width + 2.5 or y + h > height + 2.5 or w > width * 1.35 or h > height * 1.35:
            bad += 1
    return min(0.35, bad * 0.06)


def property_bloat_penalty(target_objects, actual_objects):
    target_keys = sum(len(obj.get("properties", {})) for obj in target_objects if isinstance(obj.get("properties"), dict))
    actual_keys = sum(len(obj.get("properties", {})) for obj in actual_objects if isinstance(obj.get("properties"), dict))
    if actual_keys <= max(8, target_keys * 1.45):
        return 0.0
    excess_ratio = (actual_keys - max(8, target_keys * 1.45)) / max(8, target_keys)
    return min(0.25, excess_ratio * 0.12)


def anti_hack_penalty(target_objects, actual_objects, matches, extra_input_indexes, slide):
    target_count = len(target_objects)
    actual_count = len(actual_objects)
    excess_count = max(0, actual_count - target_count)
    extra_ratio_penalty = min(0.45, 0.035 * len(extra_input_indexes) + 0.055 * excess_count)
    return min(
        0.85,
        extra_ratio_penalty
        + duplicate_penalty(actual_objects)
        + huge_box_penalty(actual_objects, target_objects, slide)
        + invalid_bbox_penalty(actual_objects, slide)
        + property_bloat_penalty(target_objects, actual_objects),
    )


def background_reward(target, actual):
    if target is None and actual is None:
        return 1.0
    if target is None or actual is None:
        return -0.65
    if not isinstance(target, dict) or not isinstance(actual, dict):
        return -0.65
    return property_reward("background", target.get("properties", {}), actual.get("properties", {}))


def grade_json(target_json, input_json):
    target = load_json(target_json)
    actual = load_json(input_json)
    slide = target.get("slide", {}) if isinstance(target, dict) else {}
    target_objects = target.get("objects", []) if isinstance(target, dict) else []
    actual_objects = actual.get("objects", []) if isinstance(actual, dict) else []
    if not isinstance(actual_objects, list):
        actual_objects = []

    matches, extra_input_indexes = match_objects(target_objects, actual_objects, slide)
    object_rewards = [match["scores"]["reward"] for match in matches]
    object_score = sum(object_rewards) / max(1, len(target_objects))
    bg_score = background_reward(target.get("background"), actual.get("background")) if isinstance(target, dict) and isinstance(actual, dict) else -0.65
    schema_score = schema_reward(actual)
    count_score = count_reward(len(target_objects), len(actual_objects))
    type_score = type_distribution_reward(target_objects, actual_objects)
    precision_score = object_precision_reward(matches, len(actual_objects))
    recall_score = object_recall_reward(matches, len(target_objects))
    property_recall_score = property_recall_reward(matches, target_objects, actual_objects)
    hack_penalty = anti_hack_penalty(target_objects, actual_objects, matches, extra_input_indexes, slide)

    reward = clamp(
        0.04 * schema_score
        + 0.06 * bg_score
        + 0.07 * count_score
        + 0.07 * type_score
        + 0.07 * precision_score
        + 0.09 * recall_score
        + 0.25 * property_recall_score
        + 0.35 * object_score
        - hack_penalty
    )

    return {
        "reward": reward,
        "score": (reward + 1.0) / 2.0,
        "schema_reward": schema_score,
        "background_reward": bg_score,
        "count_reward": count_score,
        "type_distribution_reward": type_score,
        "precision_reward": precision_score,
        "recall_reward": recall_score,
        "property_recall_reward": property_recall_score,
        "object_reward": object_score,
        "anti_hack_penalty": hack_penalty,
        "matched_objects": len([match for match in matches if match["input_index"] is not None]),
        "target_objects": len(target_objects),
        "input_objects": len(actual_objects),
        "extra_input_indexes": extra_input_indexes,
        "matches": matches,
    }


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument("target")
    parser.add_argument("input")
    args = parser.parse_args()
    print(json.dumps(grade_json(args.target, args.input), indent=2))
