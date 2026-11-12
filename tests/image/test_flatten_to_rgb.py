"""`flatten_to_rgb` composites every transparent mode onto white, the one rule JPG and PDF output share."""

import sys
import unittest
from pathlib import Path

from PIL import Image

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "image"))

from _image_to_jpg_converter import flatten_to_rgb  # noqa: E402

WHITE = (255, 255, 255)


class FlattenToRgbTests(unittest.TestCase):
    def test_transparent_pixels_become_white_in_every_alpha_mode(self) -> None:
        for mode, transparent in (("RGBA", (10, 20, 30, 0)), ("LA", (40, 0))):
            with self.subTest(mode=mode):
                flat = flatten_to_rgb(Image.new(mode, (2, 2), transparent))
                self.assertEqual(flat.mode, "RGB")
                self.assertEqual(flat.getpixel((0, 0)), WHITE)

    def test_palette_image_with_transparency_becomes_white(self) -> None:
        palette = Image.new("RGBA", (2, 2), (10, 20, 30, 0)).convert("P")
        palette.info["transparency"] = 0
        self.assertEqual(flatten_to_rgb(palette).getpixel((0, 0)), WHITE)

    def test_opaque_pixels_keep_their_colour(self) -> None:
        flat = flatten_to_rgb(Image.new("RGBA", (2, 2), (10, 20, 30, 255)))
        self.assertEqual(flat.getpixel((0, 0)), (10, 20, 30))

    def test_other_modes_are_converted_without_a_background(self) -> None:
        self.assertEqual(flatten_to_rgb(Image.new("L", (2, 2), 7)).getpixel((0, 0)), (7, 7, 7))
        rgb = Image.new("RGB", (2, 2), (1, 2, 3))
        self.assertIs(flatten_to_rgb(rgb), rgb)


if __name__ == "__main__":
    unittest.main()
