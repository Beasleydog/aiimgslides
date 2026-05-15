import argparse
import json
import os
from types import SimpleNamespace
from pathlib import Path

import win32com.client

from export_powerpoint import export_slides_with_svg_with_app


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--pptx", required=True)
    parser.add_argument("--export-dir", required=True)
    parser.add_argument("--elements", required=True)
    parser.add_argument("--image", required=True)
    parser.add_argument("--format", default="JPG")
    args = parser.parse_args()

    slide_elements = [
        [SimpleNamespace(**el) for el in elements]
        for elements in json.loads(Path(args.elements).read_text(encoding="utf-8"))
    ]

    app = None
    try:
        app = win32com.client.Dispatch("PowerPoint.Application")
        try:
            app.DisplayAlerts = 0
        except Exception:
            pass
        exported = export_slides_with_svg_with_app(
            app,
            Path(args.pptx),
            Path(args.export_dir),
            slide_elements,
            Path(args.image),
            args.format,
        )
        for path in exported:
            print(os.fspath(path))
    finally:
        if app is not None:
            app.Quit()


if __name__ == "__main__":
    main()
