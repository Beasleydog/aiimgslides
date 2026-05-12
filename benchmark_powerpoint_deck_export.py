import shutil
import time
from pathlib import Path

import win32com.client

import config
from export_powerpoint import export_slides_with_app
from generator import make_elements
from image_element import download_image
from render import add_deck_to_pptx
from svg_element import insert_svg_elements_with_app


RUNS = config.DECK_EXPORT_BENCHMARK_SLIDES
BENCH_DIR = Path("benchmark_deck_output")


def benchmark():
    BENCH_DIR.mkdir(exist_ok=True)
    export_dir = BENCH_DIR / "export"
    if export_dir.exists():
        shutil.rmtree(export_dir)

    image = download_image()
    image_path = BENCH_DIR / "_example_image.jpg"
    pptx_path = BENCH_DIR / f"{RUNS}_slides.pptx"
    image.save(image_path, quality=92)

    try:
        build_start = time.perf_counter()
        slide_elements = [make_elements() for _ in range(RUNS)]
        add_deck_to_pptx(slide_elements, image_path, pptx_path)
        build_elapsed = time.perf_counter() - build_start

        app = win32com.client.Dispatch("PowerPoint.Application")
        try:
            svg_start = time.perf_counter()
            if config.INSERT_SVG_WITH_POWERPOINT:
                insert_svg_elements_with_app(app, pptx_path, slide_elements, image_path)
            svg_elapsed = time.perf_counter() - svg_start

            export_start = time.perf_counter()
            exported = export_slides_with_app(app, pptx_path, export_dir, "JPG")
            export_elapsed = time.perf_counter() - export_start
        finally:
            app.Quit()
    finally:
        image_path.unlink(missing_ok=True)

    print(f"slides: {RUNS}")
    print(f"pptx build: {build_elapsed:.3f}s ({build_elapsed / RUNS:.3f}s/slide)")
    print(f"svg com insert: {svg_elapsed:.3f}s ({svg_elapsed / RUNS:.3f}s/slide)")
    print(f"powerpoint export once: {export_elapsed:.3f}s ({export_elapsed / RUNS:.3f}s/image)")
    print(f"total build+svg+export: {build_elapsed + svg_elapsed + export_elapsed:.3f}s ({(build_elapsed + svg_elapsed + export_elapsed) / RUNS:.3f}s/image)")
    print(f"exported images: {len(exported)}")
    if exported:
        print(f"first: {exported[0]}")


if __name__ == "__main__":
    benchmark()
