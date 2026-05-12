import shutil
import time
from pathlib import Path

import win32com.client

import config
from export_powerpoint import export_slides_with_app
from index import main


RUNS = 10
BENCH_DIR = Path("benchmark_batch_output")


def benchmark():
    BENCH_DIR.mkdir(exist_ok=True)
    timings = []

    old_pptx = config.OUTPUT_PPTX
    old_png = config.OUTPUT_PNG
    app = None

    try:
        jobs = []
        for i in range(1, RUNS + 1):
            pptx_path = BENCH_DIR / f"slide_{i:02}.pptx"
            preview_path = BENCH_DIR / f"preview_{i:02}.png"
            export_dir = BENCH_DIR / f"export_{i:02}"

            if export_dir.exists():
                shutil.rmtree(export_dir)

            config.OUTPUT_PPTX = str(pptx_path)
            config.OUTPUT_PNG = str(preview_path)
            main()
            jobs.append((i, pptx_path, export_dir))

        startup_start = time.perf_counter()
        app = win32com.client.Dispatch("PowerPoint.Application")
        startup_elapsed = time.perf_counter() - startup_start

        for i, pptx_path, export_dir in jobs:
            start = time.perf_counter()
            exported = export_slides_with_app(app, pptx_path, export_dir, "JPG")
            elapsed = time.perf_counter() - start
            timings.append(elapsed)
            print(f"{i:02}: {elapsed:.3f}s -> {exported[0] if exported else 'no image'}")
    finally:
        config.OUTPUT_PPTX = old_pptx
        config.OUTPUT_PNG = old_png
        if app is not None:
            app.Quit()

    avg = sum(timings) / len(timings)
    print()
    print(f"powerpoint startup: {startup_elapsed:.3f}s")
    print(f"runs: {len(timings)}")
    print(f"avg export only: {avg:.3f}s")
    print(f"min export only: {min(timings):.3f}s")
    print(f"max export only: {max(timings):.3f}s")
    print(f"avg incl startup amortized: {(sum(timings) + startup_elapsed) / len(timings):.3f}s")


if __name__ == "__main__":
    benchmark()
