#!/usr/bin/env python3
"""Smart Life / Tuya light controller via local network using tinytuya."""

import argparse
import json
import logging
import sys
from pathlib import Path
from typing import Any, Dict, List

import tinytuya

SCRIPT_DIR = Path(__file__).resolve().parent
DEVICES_FILE = SCRIPT_DIR / "devices.json"

logger = logging.getLogger(__name__)

# ── DPS mappings (Tuya Data Point Schema) ────────────────────────────────────
# DPS 20 = on/off for most Tuya light bulbs
# DPS 1  = on/off for most Tuya plugs/switches
DPS_SWITCH_BULB = "20"
DPS_SWITCH_PLUG = "1"


def _load_devices(path: Path) -> List[Dict[str, Any]]:
    """Load device list from JSON file.

    Accepts both the raw tinytuya scan format ``{"devices": [...]}``
    and a plain ``[...]`` array.
    """
    if not path.exists():
        logger.error(f"❌ Devices file not found: {path}")
        logger.info(
            "ℹ️  Run 'python -m tinytuya scan' to generate devices.json, "
            "then copy it to the smart_life/ folder.  See README for setup steps."
        )
        sys.exit(1)

    with open(path, "r", encoding="utf-8") as fh:
        data = json.load(fh)

    if isinstance(data, dict):
        data = data.get("devices", [])

    if not isinstance(data, list) or len(data) == 0:
        logger.error("❌ devices.json contains no devices")
        sys.exit(1)

    return data


def _find_device(devices: List[Dict[str, Any]], query: str) -> Dict[str, Any]:
    """Find a device by name (case-insensitive partial match) or by ID."""
    query_lower = query.lower()
    for dev in devices:
        if dev.get("id", "").lower() == query_lower:
            return dev
    for dev in devices:
        if query_lower in dev.get("name", "").lower():
            return dev

    logger.error(f"❌ No device matching '{query}' found in devices.json")
    logger.info("ℹ️  Available devices:")
    for dev in devices:
        logger.info(f"   • {dev.get('name', '?')}  (id={dev.get('id', '?')})")
    sys.exit(1)


def _get_version(device_info: Dict[str, Any]) -> float:
    """Extract protocol version, accepting both 'version' and 'ver' keys."""
    return float(device_info.get("version", device_info.get("ver", "3.3")))


def _connect(device_info: Dict[str, Any]) -> tinytuya.Device:
    """Create a tinytuya Device connection."""
    dev_id = device_info["id"]
    ip = device_info.get("ip", "Auto")
    local_key = device_info["key"]
    version = _get_version(device_info)

    dev = tinytuya.Device(dev_id, ip, local_key, version=version)
    dev.set_socketPersistent(False)
    return dev


def _detect_switch_dps(status: Dict[str, Any]) -> str:
    """Auto-detect whether the device uses DPS 20 (bulb) or DPS 1 (plug)."""
    dps = status.get("dps", {})
    if DPS_SWITCH_BULB in dps:
        return DPS_SWITCH_BULB
    if DPS_SWITCH_PLUG in dps:
        return DPS_SWITCH_PLUG
    logger.warning(f"⚠️  Could not auto-detect switch DPS. Raw DPS: {dps}")
    return DPS_SWITCH_PLUG


def turn_on(device_info: Dict[str, Any]) -> None:
    """Turn a light/device ON."""
    dev = _connect(device_info)
    name = device_info.get("name", device_info["id"])

    status = dev.status()
    if "Error" in str(status):
        logger.error(f"❌ Cannot reach '{name}': {status}")
        sys.exit(1)

    dps_key = _detect_switch_dps(status)
    dev.set_value(dps_key, True)
    logger.info(f"✅ '{name}' turned ON")


def turn_off(device_info: Dict[str, Any]) -> None:
    """Turn a light/device OFF."""
    dev = _connect(device_info)
    name = device_info.get("name", device_info["id"])

    status = dev.status()
    if "Error" in str(status):
        logger.error(f"❌ Cannot reach '{name}': {status}")
        sys.exit(1)

    dps_key = _detect_switch_dps(status)
    dev.set_value(dps_key, False)
    logger.info(f"✅ '{name}' turned OFF")


def get_status(device_info: Dict[str, Any]) -> None:
    """Print current device status."""
    dev = _connect(device_info)
    name = device_info.get("name", device_info["id"])

    status = dev.status()
    if "Error" in str(status):
        logger.error(f"❌ Cannot reach '{name}': {status}")
        sys.exit(1)

    dps = status.get("dps", {})
    dps_key = _detect_switch_dps(status)
    is_on = dps.get(dps_key, None)
    state_label = "ON" if is_on else "OFF" if is_on is not None else "UNKNOWN"

    logger.info(f"💡 '{name}' is {state_label}")
    logger.debug(f"📂 Full DPS: {json.dumps(dps, indent=2)}")


def switch_toggle(device_info: Dict[str, Any]) -> None:
    """Toggle device: if ON turn OFF, if OFF turn ON."""
    dev = _connect(device_info)
    name = device_info.get("name", device_info["id"])

    status = dev.status()
    if "Error" in str(status):
        logger.error(f"❌ Cannot reach '{name}': {status}")
        sys.exit(1)

    dps_key = _detect_switch_dps(status)
    dps = status.get("dps", {})
    is_on = dps.get(dps_key, False)
    new_state = not is_on
    dev.set_value(dps_key, new_state)
    state_label = "ON" if new_state else "OFF"
    logger.info(f"✅ '{name}' switched to {state_label}")


def list_devices(devices: List[Dict[str, Any]]) -> None:
    """Print all configured devices."""
    logger.info(f"📋 {len(devices)} device(s) in {DEVICES_FILE.name}:\n")
    for dev in devices:
        ip = dev.get("ip", "?")
        ver = dev.get("version", dev.get("ver", "?"))
        logger.info(f"   • {dev.get('name', '?'):30s}  ip={ip:15s}  ver={ver}")


def scan_network() -> None:
    """Run tinytuya network scan to discover devices."""
    logger.info("🔍 Scanning local network for Tuya devices …")
    tinytuya.scanner.scan()


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Control Smart Life / Tuya lights from the command line.",
        epilog=(
            "Examples:\n"
            "  python light_control.py on     --light \"luz despacho\"\n"
            "  python light_control.py off    --light \"luz despacho\"\n"
            "  python light_control.py switch --light \"luz despacho\"\n"
            "  python light_control.py status --light despacho\n"
            "  python light_control.py list\n"
            "  python light_control.py scan\n"
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "action",
        choices=["on", "off", "switch", "status", "list", "scan"],
        help="Action to perform",
    )
    parser.add_argument(
        "--light",
        "-l",
        type=str,
        default=None,
        help="Device name or ID (partial match, case-insensitive)",
    )
    parser.add_argument(
        "--devices-file",
        type=str,
        default=None,
        help=f"Path to devices.json (default: {DEVICES_FILE})",
    )
    parser.add_argument("--debug", action="store_true", help="Enable debug logging")
    return parser


def main() -> int:
    parser = _build_parser()
    args = parser.parse_args()

    log_level = logging.DEBUG if args.debug else logging.INFO
    logging.basicConfig(
        level=log_level,
        format="%(message)s",
    )

    global DEVICES_FILE
    if args.devices_file:
        DEVICES_FILE = Path(args.devices_file)

    if args.action == "scan":
        scan_network()
        return 0

    devices = _load_devices(DEVICES_FILE)

    if args.action == "list":
        list_devices(devices)
        return 0

    if not args.light:
        logger.error("❌ --light is required for on/off/switch/status actions")
        parser.print_help()
        return 1

    device_info = _find_device(devices, args.light)

    actions = {
        "on": turn_on,
        "off": turn_off,
        "switch": switch_toggle,
        "status": get_status,
    }
    actions[args.action](device_info)
    return 0


if __name__ == "__main__":
    sys.exit(main())
