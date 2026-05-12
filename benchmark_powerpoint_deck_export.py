import shutil
import time
from pathlib import Path

import win32com.client

import config
from export_powerpoint import export_slides_with_svg_with_app
from generator import make_elements
from image_element import download_image
from render import add_deck_to_pptx


RUNS = config.DECK_EXPORT_BENCHMARK_SLIDES
BENCH_DIR = Path("benchmark_deck_output")


def benchmark():
    BENCH_DIR.mkdir(exist_ok=True)
    export_dir = BENCH_DIR / "export"
    pptx_path = BENCH_DIR / f"{RUNS}_slides.pptx"
    if export_dir.exists():
        shutil.rmtree(export_dir)
    old_pptx_dir = BENCH_DIR / "pptx"
    if old_pptx_dir.exists():
        shutil.rmtree(old_pptx_dir)

    image = download_image()
    image_path = BENCH_DIR / "_example_image.jpg"
    image.save(image_path, quality=92)

    app = None
    try:
        build_start = time.perf_counter()
        slide_elements = [make_elements() for _ in range(RUNS)]
        add_deck_to_pptx(slide_elements, image_path, pptx_path)
        build_elapsed = time.perf_counter() - build_start

        app = win32com.client.Dispatch("PowerPoint.Application")
        try:
            app.DisplayAlerts = 0
        except Exception:
            pass

        export_start = time.perf_counter()
        exported = export_slides_with_svg_with_app(app, pptx_path, export_dir, slide_elements, image_path, "JPG")
        com_elapsed = time.perf_counter() - export_start
    finally:
        if app is not None:
            app.Quit()
        image_path.unlink(missing_ok=True)

    print(f"slides: {RUNS}")
    print(f"pptx build: {build_elapsed:.3f}s ({build_elapsed / RUNS:.3f}s/slide)")
    print(f"powerpoint svg+export once: {com_elapsed:.3f}s ({com_elapsed / RUNS:.3f}s/image)")
    print(f"total: {build_elapsed + com_elapsed:.3f}s ({(build_elapsed + com_elapsed) / RUNS:.3f}s/image)")
    print(f"exported images: {len(exported)}")
    if exported:
        print(f"first: {exported[0]}")


if __name__ == "__main__":
    benchmark()
