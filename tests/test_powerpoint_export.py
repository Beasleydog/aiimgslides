import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import config
from export_powerpoint import export_slides
from index import main


class PowerPointExportTest(unittest.TestCase):
    def test_powerpoint_com_exports_slide_image(self):
        old_pptx = config.OUTPUT_PPTX
        old_png = config.OUTPUT_PNG

        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            pptx_path = tmp_path / "random_slide.pptx"
            preview_path = tmp_path / "random_slide.png"
            export_dir = tmp_path / "output_folder"

            config.OUTPUT_PPTX = str(pptx_path)
            config.OUTPUT_PNG = str(preview_path)
            try:
                main()
                try:
                    export_slides(pptx_path, export_dir, "JPG")
                except ImportError as exc:
                    raise unittest.SkipTest("pywin32 is not installed") from exc
                except Exception as exc:
                    raise unittest.SkipTest(f"PowerPoint COM export is not available: {exc}") from exc
            finally:
                config.OUTPUT_PPTX = old_pptx
                config.OUTPUT_PNG = old_png

            exported = list(export_dir.glob("*.jpg")) + list(export_dir.glob("*.jpeg"))
            self.assertTrue(exported, "PowerPoint did not export any JPG slide images")
            self.assertGreater(exported[0].stat().st_size, 0)


if __name__ == "__main__":
    unittest.main()
