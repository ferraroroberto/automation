"""wifi_connect.ps1 must hex-encode the SSID as UTF-8, like wifi_bat_generator.py."""

import unittest
from pathlib import Path

WIFI_DIR = Path(__file__).resolve().parents[3] / "system" / "wifi"


class SsidHexEncodingTests(unittest.TestCase):
    def test_ps1_encodes_the_ssid_as_utf8(self):
        # ASCII turns a curly apostrophe into '?', so the profile <hex> stops matching the network.
        source = (WIFI_DIR / "wifi_connect.ps1").read_text(encoding="utf-8")
        self.assertIn("[System.Text.Encoding]::UTF8.GetBytes($Ssid)", source)
        self.assertNotIn("Encoding]::ASCII", source)

    def test_python_generator_encodes_the_ssid_as_utf8(self):
        source = (WIFI_DIR / "wifi_bat_generator.py").read_text(encoding="utf-8")
        self.assertIn('ssid.encode("utf-8")', source)


if __name__ == "__main__":
    unittest.main()
