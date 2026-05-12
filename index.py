import random
import re
import shutil
import time
from pathlib import Path

import win32com.client

import config
from dataset_io import save_slide_json
from export_powerpoint import export_slides_with_svg_with_app
from generator import make_elements
from image_element import download_image
from render import add_deck_to_pptx


# Dataset generation config.
NUM_SLIDES = 1000
OUTPUT_DIR = Path("output")
IMAGE_FORMAT = "JPG"
SEED = config.SEED

# Keep this false for clean dataset folders. Turn it on only when debugging PPTX.
KEEP_WORK_PPTX = False


def slide_number(path):
    match = re.search(r"Slide(\d+)", path.stem, re.IGNORECASE)
    return int(match.group(1)) if match else 0


def reset_output_dir():
    if OUTPUT_DIR.exists():
        shutil.rmtree(OUTPUT_DIR)
    OUTPUT_DIR.mkdir(parents=True)


def write_json_files(slide_elements):
    json_paths = []
    for index, elements in enumerate(slide_elements, start=1):
        json_path = OUTPUT_DIR / f"slide_{index:06}.json"
        save_slide_json(json_path, elements, image_file=f"slide_{index:06}.jpg")
        json_paths.append(json_path)
    return json_paths


def move_exported_images(exported):
    image_paths = []
    for index, src in enumerate(sorted(exported, key=slide_number), start=1):
        dest = OUTPUT_DIR / f"slide_{index:06}.jpg"
        shutil.move(str(src), dest)
        image_paths.append(dest)
    return image_paths


def main():
    if SEED is not None:
        random.seed(SEED)

    start = time.perf_counter()
    reset_output_dir()
    work_dir = OUTPUT_DIR / "_work"
    export_dir = work_dir / "export"
    work_dir.mkdir(parents=True)

    image = download_image()
    image_path = work_dir / "_example_image.jpg"
    pptx_path = work_dir / "slides.pptx"
    image.save(image_path, quality=92)

    app = None
    try:
        slide_elements = [make_elements() for _ in range(NUM_SLIDES)]
        add_deck_to_pptx(slide_elements, image_path, pptx_path)
        write_json_files(slide_elements)

        app = win32com.client.Dispatch("PowerPoint.Application")
        try:
            app.DisplayAlerts = 0
        except Exception:
            pass
        exported = export_slides_with_svg_with_app(app, pptx_path, export_dir, slide_elements, image_path, IMAGE_FORMAT)
        image_paths = move_exported_images(exported)
    finally:
        if app is not None:
            app.Quit()
        if not KEEP_WORK_PPTX:
            shutil.rmtree(work_dir, ignore_errors=True)

    elapsed = time.perf_counter() - start
    print(f"wrote {len(image_paths)} image/json pairs to {OUTPUT_DIR.resolve()}")
    print(f"elapsed {elapsed:.3f}s ({elapsed / max(1, len(image_paths)):.3f}s/slide)")


if __name__ == "__main__":
    main()
