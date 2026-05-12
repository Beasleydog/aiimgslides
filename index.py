import random
from pathlib import Path

import config
from generator import make_elements
from image_element import download_image
from render import add_to_png, add_to_pptx
from svg_element import insert_svg_elements


def main():
    if config.SEED is not None:
        random.seed(config.SEED)

    image = download_image()
    image_path = Path("_example_image.jpg")
    image.save(image_path, quality=92)

    elements = make_elements()
    try:
        add_to_pptx(elements, image_path)
        if config.INSERT_SVG_WITH_POWERPOINT:
            try:
                insert_svg_elements(config.OUTPUT_PPTX, [elements], image_path)
            except Exception as exc:
                print(f"skipped SVG PowerPoint insertion: {exc}")
        add_to_png(elements, image)
    finally:
        image_path.unlink(missing_ok=True)

    print(f"wrote {config.OUTPUT_PPTX} and {config.OUTPUT_PNG}")


if __name__ == "__main__":
    main()
