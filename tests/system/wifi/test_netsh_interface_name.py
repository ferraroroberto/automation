"""parse_interface_name must find the Wi-Fi interface whatever the adapter is called."""

import sys
import unittest
from pathlib import Path

WIFI_DIR = Path(__file__).resolve().parents[3] / "tools" / "system" / "wifi"
sys.path.insert(0, str(WIFI_DIR))
from _netsh import parse_interface_name  # noqa: E402

SHOW_INTERFACES = """
There is 1 interface on the system:

    Name                   : {name}
    Description            : Some Wireless Adapter
    GUID                   : 00000000-0000-0000-0000-000000000000
    State                  : connected
    SSID                   : HomeNet
"""


class ParseInterfaceNameTests(unittest.TestCase):
    def test_any_adapter_name_is_returned(self):
        for name in ("Wi-Fi", "WLAN 2", "Ethernet 3", "Wifi integrado"):
            with self.subTest(name=name):
                self.assertEqual(parse_interface_name(SHOW_INTERFACES.format(name=name)), name)

    def test_a_colon_inside_the_name_survives(self):
        self.assertEqual(parse_interface_name(SHOW_INTERFACES.format(name="Wi-Fi: office")), "Wi-Fi: office")

    def test_the_first_interface_wins(self):
        output = SHOW_INTERFACES.format(name="Wi-Fi") + SHOW_INTERFACES.format(name="Wi-Fi 2")
        self.assertEqual(parse_interface_name(output), "Wi-Fi")

    def test_no_name_line_is_none(self):
        self.assertIsNone(parse_interface_name("The Wireless AutoConfig Service (wlansvc) is not running."))
        self.assertIsNone(parse_interface_name(""))

    def test_description_and_ssid_lines_are_not_mistaken_for_a_name(self):
        self.assertIsNone(parse_interface_name("    Description : Name : x\n    SSID : Name\n"))


if __name__ == "__main__":
    unittest.main()
