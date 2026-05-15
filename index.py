import json
import random
import re
import shutil
import subprocess
import sys
import time
from pathlib import Path

import config
from dataset_io import save_slide_json
from generator import make_elements
from image_element import download_image
from render import add_deck_to_pptx


# Dataset generation config.
# Curriculum levels create the same number of slides for each object count:
# level 1 has 1 foreground object, level 2 has 2, ..., up to max level.
MAX_CURRICULUM_ELEMENTS = 25
SLIDES_PER_CURRICULUM_LEVEL = 40
NUM_SLIDES = MAX_CURRICULUM_ELEMENTS * SLIDES_PER_CURRICULUM_LEVEL
EXPORT_BATCH_SIZE = 100
OUTPUT_DIR = Path("output")
IMAGE_FORMAT = "JPG"
SEED = config.SEED
POWERPOINT_EXPORT_TIMEOUT_SECONDS = 600

# Keep this false for clean dataset folders. Turn it on only when debugging PPTX.
KEEP_WORK_PPTX = False


def slide_number(path):
    match = re.search(r"Slide(\d+)", path.stem, re.IGNORECASE)
    return int(match.group(1)) if match else 0


def reset_output_dir():
    if OUTPUT_DIR.exists():
        shutil.rmtree(OUTPUT_DIR)
    OUTPUT_DIR.mkdir(parents=True)


def curriculum_plan():
    return [
        level
        for level in range(1, MAX_CURRICULUM_ELEMENTS + 1)
        for _ in range(SLIDES_PER_CURRICULUM_LEVEL)
    ]


def write_json_files(slide_elements, levels, start_index=1):
    json_paths = []
    for index, (elements, level) in enumerate(zip(slide_elements, levels), start=start_index):
        json_path = OUTPUT_DIR / f"slide_{index:06}.json"
        content_count = sum(1 for el in elements if el.kind != "background")
        save_slide_json(
            json_path,
            elements,
            image_file=f"slide_{index:06}.jpg",
            curriculum={
                "level": level,
                "target_content_count": level,
                "actual_content_count": content_count,
                "slides_per_level": SLIDES_PER_CURRICULUM_LEVEL,
                "max_level": MAX_CURRICULUM_ELEMENTS,
            },
        )
        json_paths.append(json_path)
    return json_paths


def move_exported_images(exported, start_index=1):
    image_paths = []
    for index, src in enumerate(sorted(exported, key=slide_number), start=start_index):
        dest = OUTPUT_DIR / f"slide_{index:06}.jpg"
        shutil.move(str(src), dest)
        image_paths.append(dest)
    return image_paths


def batched(items, size):
    for start in range(0, len(items), size):
        yield start, items[start : start + size]


def kill_powerpoint():
    subprocess.run(
        ["taskkill", "/IM", "POWERPNT.EXE", "/F"],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        check=False,
    )


def export_batch_with_powerpoint(batch_pptx, batch_export_dir, batch_elements, image_path, batch_index):
    elements_path = batch_pptx.with_suffix(".svg-elements.json")
    svg_elements = [
        [
            {"kind": el.kind, "x": el.x, "y": el.y, "w": el.w, "h": el.h, "data": el.data}
            for el in elements
            if el.kind in {"svg", "svg_image"}
        ]
        for elements in batch_elements
    ]
    elements_path.write_text(json.dumps(svg_elements), encoding="utf-8")

    cmd = [
        sys.executable,
        "powerpoint_export_worker.py",
        "--pptx",
        str(batch_pptx),
        "--export-dir",
        str(batch_export_dir),
        "--elements",
        str(elements_path),
        "--image",
        str(image_path),
        "--format",
        IMAGE_FORMAT,
    ]
    try:
        result = subprocess.run(
            cmd,
            check=True,
            timeout=POWERPOINT_EXPORT_TIMEOUT_SECONDS,
            capture_output=True,
            text=True,
        )
    except subprocess.TimeoutExpired as exc:
        kill_powerpoint()
        raise RuntimeError(f"PowerPoint export timed out for batch {batch_index}") from exc
    except subprocess.CalledProcessError as exc:
        print(exc.stdout or "")
        print(exc.stderr or "")
        raise

    return sorted(batch_export_dir.glob(f"*.{IMAGE_FORMAT.lower()}"), key=slide_number)


def main():
    if SEED is not None:
        random.seed(SEED)

    start = time.time()
    reset_output_dir()
    work_dir = OUTPUT_DIR / "_work"
    export_dir = work_dir / "export"
    work_dir.mkdir(parents=True)

    image = download_image()
    image_path = work_dir / "_example_image.jpg"
    image.save(image_path, quality=92)

    try:
        image_paths = []
        levels = curriculum_plan()
        for batch_index, (start, batch_levels) in enumerate(batched(levels, EXPORT_BATCH_SIZE), start=1):
            batch_start_index = start + 1
            batch_elements = [
                make_elements(target_content_count=level, minimum_content_count=level)
                for level in batch_levels
            ]
            batch_pptx = work_dir / f"slides_{batch_index:03}.pptx"
            batch_export_dir = export_dir / f"batch_{batch_index:03}"
            add_deck_to_pptx(batch_elements, image_path, batch_pptx)
            write_json_files(batch_elements, batch_levels, start_index=batch_start_index)
            exported = export_batch_with_powerpoint(
                batch_pptx,
                batch_export_dir,
                batch_elements,
                image_path,
                batch_index,
            )
            image_paths.extend(move_exported_images(exported, start_index=batch_start_index))
            print(f"exported batch {batch_index}: slides {batch_start_index}-{batch_start_index + len(batch_levels) - 1}")
    finally:
        if not KEEP_WORK_PPTX:
            shutil.rmtree(work_dir, ignore_errors=True)

    elapsed = time.time() - start
    print(f"wrote {len(image_paths)} image/json pairs to {OUTPUT_DIR.resolve()}")
    print(f"elapsed {elapsed:.3f}s ({elapsed / max(1, len(image_paths)):.3f}s/slide)")


if __name__ == "__main__":
    main()
