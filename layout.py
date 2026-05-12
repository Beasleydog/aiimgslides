import random

import config


def overlap_area(a, b):
    ax, ay, aw, ah = padded_box(a)
    bx, by, bw, bh = padded_box(b)
    left = max(ax, bx)
    top = max(ay, by)
    right = min(ax + aw, bx + bw)
    bottom = min(ay + ah, by + bh)
    return max(0, right - left) * max(0, bottom - top)


def padded_box(el):
    pad = config.COLLISION_PADDING_BY_KIND.get(el.kind, 0)
    return el.x - pad, el.y - pad, el.w + pad * 2, el.h + pad * 2


def edge_penalty(el):
    left = el.x
    right = config.SLIDE_W - (el.x + el.w)
    top = el.y
    bottom = config.SLIDE_H - (el.y + el.h)
    nearest_edge = min(left, right, top, bottom)
    return max(0, config.SLIDE_MARGIN * 2 - nearest_edge)


def score_element(candidate, placed):
    score = edge_penalty(candidate) * config.EDGE_PENALTY_WEIGHT
    candidate_area = max(candidate.w * candidate.h, 0.01)
    for existing in placed:
        area = overlap_area(candidate, existing)
        existing_area = max(existing.w * existing.h, 0.01)
        normalized_overlap = area / min(candidate_area, existing_area)
        score += area * config.OVERLAP_PENALTY_WEIGHT
        score += normalized_overlap * config.OVERLAP_PENALTY_WEIGHT
        if candidate.kind == existing.kind:
            score += area * config.SAME_KIND_OVERLAP_WEIGHT
            score += normalized_overlap * config.SAME_KIND_OVERLAP_WEIGHT
    return score


def random_box(size_range):
    min_w, max_w, min_h, max_h = size_range
    w = random.uniform(min_w, max_w)
    h = random.uniform(min_h, max_h)
    margin = config.SLIDE_MARGIN
    return (
        random.uniform(margin, config.SLIDE_W - w - margin),
        random.uniform(margin, config.SLIDE_H - h - margin),
        w,
        h,
    )


def place_element(factory, placed, size_range):
    best = None
    best_score = float("inf")
    for _ in range(config.PLACEMENT_ATTEMPTS):
        candidate = factory(*random_box(size_range))
        score = score_element(candidate, placed)
        if score < best_score:
            best = candidate
            best_score = score
    return best
