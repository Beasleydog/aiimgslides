import shutil
import time
from pathlib import Path

import win32com.client

from export_powerpoint import export_slides_with_svg_with_app
from generator import make_elements
from image_element import download_image
from render import add_deck_to_pptx


RUNS = 20
OUTPUT_DIR = Path("twenty_image_output")


def main():
    OUTPUT_DIR.mkdir(exist_ok=True)
    export_dir = OUTPUT_DIR / "images"
    if export_dir.exists():
        shutil.rmtree(export_dir)
    export_dir.mkdir(parents=True, exist_ok=True)

    image = download_image()
    image_path = OUTPUT_DIR / "_example_image.jpg"
    pptx_path = OUTPUT_DIR / "20_slides.pptx"
    image.save(image_path, quality=92)

    try:
        start = time.perf_counter()
        slide_elements = [make_elements() for _ in range(RUNS)]
        add_deck_to_pptx(slide_elements, image_path, pptx_path)

        app = win32com.client.Dispatch("PowerPoint.Application")
        try:
            try:
                app.DisplayAlerts = 0
            except Exception:
                pass
            exported = export_slides_with_svg_with_app(app, pptx_path, export_dir, slide_elements, image_path, "JPG")
        finally:
            app.Quit()
    finally:
        image_path.unlink(missing_ok=True)

    elapsed = time.perf_counter() - start
    print(f"exported {len(exported)} images to {export_dir}")
    print(f"wrote {pptx_path}")
    print(f"elapsed {elapsed:.3f}s")


if __name__ == "__main__":
    main()
