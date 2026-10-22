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

The Gaming, IoT, Smart and VM classes had the same drift (issue #175): hand-kept
prefix tuples that disagreed with the database (``00:03:ff`` is Microsoft
Corporation there, yet was tagged a gaming console). They now come from the
vendor name too, so the tests below sample real database MACs per class.

Run from the repo root:

    & .\\.venv\\Scripts\\python.exe -m unittest discover -s system/wifi -p "test_*.py"
"""

from __future__ import annotations

import sys
import unittest
from pathlib import Path

WIFI_DIR = Path(__file__).resolve().parents[3] / "system" / "wifi"
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


class TestVendorDerivedClasses(unittest.TestCase):
    """Gaming/IoT/Smart/VM classes follow the vendor name in MAC_OUI_DATABASE."""

    def classify(self, oui: str) -> str:
        return ns.get_device_type(f"{oui}:11:22:33", "192.168.1.5")

    def test_every_sampled_oui_is_in_the_database(self) -> None:
        # Guards the samples below against a database edit silently turning
        # them into "Unknown" lookups.
        for oui in (
            "00:03:0f", "00:0d:93", "00:19:c5", "00:0d:4c", "00:11:dc",
            "b8:27:eb", "dc:a6:32", "e4:5f:01", "00:1a:11", "3c:5a:b4",
            "00:fc:8b", "18:74:2e", "00:50:56", "00:0c:29", "00:05:69",
            "08:00:27", "52:54:00", "00:03:ff", "0e:0e:0e",
        ):
            with self.subTest(oui=oui):
                self.assertIn(oui, ns.MAC_OUI_DATABASE)

    def test_sony_and_xbox_are_gaming_consoles(self) -> None:
        for oui in ("00:03:0f", "00:0d:93", "00:19:c5", "00:0d:4c", "00:11:dc"):
            with self.subTest(oui=oui):
                self.assertEqual(self.classify(oui), "Gaming Console")

    def test_microsoft_corporation_oui_00_03_ff_is_not_a_gaming_console(self) -> None:
        # Tagged "Gaming Console" by the old hand-kept tuple, but the database
        # says Microsoft Corporation (not Xbox).
        self.assertEqual(ns.get_manufacturer("00:03:ff:11:22:33"), "Microsoft Corporation")
        self.assertEqual(self.classify("00:03:ff"), "Computer/Network Device")

    def test_raspberry_pi_is_iot_device(self) -> None:
        for oui in ("b8:27:eb", "dc:a6:32", "e4:5f:01"):
            with self.subTest(oui=oui):
                self.assertEqual(self.classify(oui), "IoT Device")

    def test_google_and_amazon_are_smart_devices(self) -> None:
        for oui in ("00:1a:11", "3c:5a:b4", "00:fc:8b", "18:74:2e"):
            with self.subTest(oui=oui):
                self.assertEqual(self.classify(oui), "Smart Device")

    def test_google_cloud_is_not_a_smart_device(self) -> None:
        self.assertEqual(self.classify("0e:0e:0e"), "Computer/Network Device")

    def test_hypervisor_vendors_are_virtual_machines(self) -> None:
        # 00:05:69 is a VMware OUI outside the old four-prefix tuple.
        for oui in ("00:50:56", "00:0c:29", "00:05:69", "08:00:27", "52:54:00"):
            with self.subTest(oui=oui):
                self.assertEqual(self.classify(oui), "Virtual Machine")

    def test_routers_and_apple_keep_precedence_over_later_classes(self) -> None:
        self.assertEqual(self.classify("00:14:6c"), "Router/AP")
        self.assertEqual(self.classify("60:33:4b"), "Apple Device")


if __name__ == "__main__":
    unittest.main()
