import os
import tempfile
from pathlib import Path

import config


def exported_files(output_path, image_format):
    suffix = f".{image_format.lower()}"
    return sorted(
        path.resolve()
        for path in Path(output_path).iterdir()
        if path.is_file() and path.suffix.lower() == suffix
    )


def export_slides(pptx_path=config.OUTPUT_PPTX, output_folder="output_folder", image_format="JPG"):
    import win32com.client

    output_path = Path(output_folder)
    output_path.mkdir(parents=True, exist_ok=True)

    app = None
    presentation = None
    try:
        app = win32com.client.Dispatch("PowerPoint.Application")
        presentation = app.Presentations.Open(os.path.abspath(pptx_path), WithWindow=False)
        presentation.Export(os.path.abspath(output_path), image_format)
    finally:
        if presentation is not None:
            presentation.Close()
        if app is not None:
            app.Quit()

    return exported_files(output_path, image_format)


def export_slides_with_app(app, pptx_path, output_folder, image_format="JPG"):
    output_path = Path(output_folder)
    output_path.mkdir(parents=True, exist_ok=True)

    presentation = None
    try:
        presentation = app.Presentations.Open(os.path.abspath(pptx_path), WithWindow=False)
        presentation.Export(os.path.abspath(output_path), image_format)
    finally:
        if presentation is not None:
            presentation.Close()

    return exported_files(output_path, image_format)


def export_slides_with_svg_with_app(
    app,
    pptx_path,
    output_folder,
    slide_elements,
    image_path,
    image_format="JPG",
    save_pptx=True,
):
    from svg_element import insert_svg_elements_into_presentation

    output_path = Path(output_folder)
    output_path.mkdir(parents=True, exist_ok=True)

    presentation = None
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmp:
        try:
            presentation = app.Presentations.Open(os.path.abspath(pptx_path), WithWindow=False)
            if config.INSERT_SVG_WITH_POWERPOINT:
                insert_svg_elements_into_presentation(presentation, slide_elements, image_path, Path(tmp))
                if save_pptx:
                    presentation.Save()
            presentation.Export(os.path.abspath(output_path), image_format)
        finally:
            if presentation is not None:
                presentation.Close()

    return exported_files(output_path, image_format)


if __name__ == "__main__":
    exported = export_slides()
    for path in exported:
        print(path)
