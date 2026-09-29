#!/usr/bin/env python3
"""Guard: get_device_type() must classify routers/APs and Apple devices from
the resolved vendor name, not by matching vendor substrings against the raw
MAC hex (issue #145).

Before the fix, ``get_device_type`` tested whether strings like ``"netgear"``
or ``"cisco"`` occurred inside the lowercased MAC address itself — a hex
string that can never contain those letters, so the "Router/AP" branch was
unreachable and every router fell through to "Computer/Network Device" in
every scan. The Apple branch used a separately hand-maintained 4-prefix tuple
that had already drifted out of sync with ``MAC_OUI_DATABASE`` (9+ Apple
prefixes there vs. 4 in the tuple), so real Apple MACs outside those 4
prefixes were also missed.

Run from the repo root:

    & .\\.venv\\Scripts\\python.exe -m unittest discover -s system/wifi -p "test_*.py"
"""

from __future__ import annotations

import sys
import unittest
from pathlib import Path

WIFI_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(WIFI_DIR))

import network_scanner as ns  # noqa: E402


class TestGetDeviceType(unittest.TestCase):
    def test_netgear_mac_classified_as_router_ap(self) -> None:
        # "00:14:6c" is a real Netgear OUI in MAC_OUI_DATABASE.
        self.assertEqual(
            ns.get_device_type("00:14:6c:aa:bb:cc", "192.168.1.1"),
            "Router/AP",
        )

    def test_cisco_mac_classified_as_router_ap(self) -> None:
        self.assertEqual(
            ns.get_device_type("00:50:cc:aa:bb:cc", "192.168.1.1"),
            "Router/AP",
        )

    def test_apple_mac_outside_old_hardcoded_tuple_classified_as_apple(self) -> None:
        # "60:33:4b" is in MAC_OUI_DATABASE as Apple Inc but was NOT one of
        # the 4 prefixes in the old hardcoded startswith() tuple.
        self.assertEqual(
            ns.get_device_type("60:33:4b:aa:bb:cc", "192.168.1.2"),
            "Apple Device",
        )

    def test_unknown_mac_falls_through_to_generic_device(self) -> None:
        self.assertEqual(
            ns.get_device_type("ff:ff:ff:ff:ff:ff", "192.168.1.50"),
            "Computer/Network Device",
        )


if __name__ == "__main__":
    unittest.main()
