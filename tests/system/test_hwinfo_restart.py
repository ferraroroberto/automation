"""Offline tests for system/hwinfo_restart.py's Stream Deck plugin stop (fake processes only)."""

import sys
import unittest
from pathlib import Path
from unittest import mock

sys.path.append(str(Path(__file__).resolve().parents[2]))

from system import hwinfo_restart as hr  # noqa: E402 -- needs the repo root on sys.path first

PLUGIN_EXE = r"C:\Users\x\AppData\Roaming\Elgato\StreamDeck\Plugins\com.exension.hwinfo.sdPlugin\hwinfo-plugin.exe"
PLUGIN_PARENT = r"C:\Users\x\AppData\Roaming\Elgato\StreamDeck\Plugins\com.exension.hwinfo.sdPlugin\hwinfo.exe"


class _Proc:
    def __init__(self, pid, name, exe, world):
        self.pid, self.info, self.world, self.terminated = pid, {"name": name, "exe": exe}, world, False

    def terminate(self):
        self.terminated = True
        self.world.remove(self)

    kill = terminate


def _iter(world):
    return lambda attrs=None: list(world)


class StopPluginTests(unittest.TestCase):
    def _stop(self, world):
        with mock.patch.object(hr.psutil, "process_iter", _iter(world)), mock.patch.object(hr.time, "sleep"):
            return hr.stop_plugin()

    def test_only_the_plugin_and_its_parent_are_stopped(self):
        world = []
        plugin, parent = _Proc(10, "hwinfo-plugin.exe", PLUGIN_EXE, world), _Proc(11, "hwinfo.exe", PLUGIN_PARENT, world)
        hwinfo64 = _Proc(12, "HWiNFO64.exe", r"C:\Program Files\HWiNFO64\HWiNFO64.exe", world)
        other = _Proc(13, "hwinfo.exe", r"C:\Tools\hwinfo.exe", world)  # right name, wrong folder
        elsewhere = _Proc(14, "notepad.exe", PLUGIN_EXE, world)  # right folder, wrong name
        unreadable = _Proc(15, "hwinfo-plugin.exe", None, world)  # exe path not readable
        world.extend([plugin, parent, hwinfo64, other, elsewhere, unreadable])
        self.assertEqual(self._stop(world), {10, 11})
        self.assertTrue(plugin.terminated and parent.terminated)
        self.assertFalse(any(p.terminated for p in (hwinfo64, other, elsewhere, unreadable)))

    def test_no_match_stops_nothing(self):
        world = []
        world.append(_Proc(1, "HWiNFO64.exe", r"C:\Program Files\HWiNFO64\HWiNFO64.exe", world))
        self.assertEqual(self._stop(world), set())
        self.assertFalse(world[0].terminated)

    def test_not_running_is_logged_and_is_not_an_error(self):
        with self.assertLogs("system.hwinfo_restart", "INFO") as logs:
            self.assertEqual(self._stop([]), set())
        self.assertIn("was not running", logs.output[0])


class RelaunchTests(unittest.TestCase):
    def _wait(self, world):
        with mock.patch.object(hr.psutil, "process_iter", _iter(world)), mock.patch.object(hr.time, "sleep"):
            return hr.wait_plugin_relaunch({10}, timeout=0.01, interval=0)

    def test_a_new_plugin_process_is_a_relaunch(self):
        world = []
        world.append(_Proc(20, "hwinfo-plugin.exe", PLUGIN_EXE, world))
        with self.assertLogs("system.hwinfo_restart", "INFO") as logs:
            self.assertTrue(self._wait(world))
        self.assertIn("plugin relaunched", logs.output[0])

    def test_no_new_process_warns_that_the_relaunch_is_not_confirmed(self):
        with self.assertLogs("system.hwinfo_restart", "WARNING") as logs:
            self.assertFalse(self._wait([]))
        self.assertIn("plugin relaunch not confirmed", logs.output[0])


if __name__ == "__main__":
    unittest.main()
