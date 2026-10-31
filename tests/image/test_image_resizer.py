"""Re-running image_resizer on a folder must not fail or reprocess its own outputs."""

import sys
import tempfile
import unittest
from pathlib import Path

from PIL import Image

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "image"))

import image_resizer  # noqa: E402


def _make_jpeg(path: Path) -> None:
    Image.effect_noise((400, 400), 80).convert("RGB").save(path, "JPEG", quality=95)


class ProcessFolderRerunTests(unittest.TestCase):
    def test_second_run_overwrites_the_previous_output_and_leaves_no_temp_files(self):
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            _make_jpeg(folder / "photo.jpg")

            image_resizer.process_folder(folder, 30)
            first = (folder / "photo_resize.jpg").stat().st_size
            image_resizer.process_folder(folder, 30)

            names = sorted(p.name for p in folder.iterdir())
            self.assertEqual(names, ["photo.jpg", "photo_resize.jpg"])
            self.assertGreater(first, 0)

    def test_failed_resize_removes_its_temp_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            _make_jpeg(folder / "photo.jpg")
            self.assertFalse(image_resizer.resize_image_to_max_size(folder / "photo.jpg", 0.001))
            self.assertEqual([p.name for p in folder.iterdir()], ["photo.jpg"])


if __name__ == "__main__":
    unittest.main()
