"""A local Tuya command that returns an error dict is a failure, not a success."""

import sys
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "tools" / "smart_life"))

import light_control  # noqa: E402

_STATUS = {"dps": {"20": False}}


def _connected(set_value_result):
    device = mock.Mock()
    device.set_value.return_value = set_value_result
    return (device, "lamp", _STATUS, None)


class LocalSetPowerTests(unittest.TestCase):
    def test_error_dict_exits_non_zero_for_every_command(self):
        error = {"Error": "Network Error: Unable to Connect", "Err": "901"}
        for command in (light_control.turn_on, light_control.turn_off, light_control.switch_toggle):
            with self.subTest(command=command.__name__):
                with mock.patch.object(light_control, "_connect_and_status", return_value=_connected(error)):
                    with self.assertRaises(SystemExit) as raised:
                        command({"id": "x"})
                self.assertEqual(raised.exception.code, 1)

    def test_acknowledged_command_is_reported_done(self):
        for ack in (None, {"dps": {"20": True}}):
            with self.subTest(ack=ack):
                with mock.patch.object(light_control, "_connect_and_status", return_value=_connected(ack)):
                    light_control.turn_on({"id": "x"})


if __name__ == "__main__":
    unittest.main()
