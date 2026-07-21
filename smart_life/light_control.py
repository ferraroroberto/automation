#!/usr/bin/env python3
"""Smart Life / Tuya light controller: cloud-first, local fallback (tinytuya)."""

import argparse
import json
import logging
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import tinytuya

SCRIPT_DIR = Path(__file__).resolve().parent
DEVICES_FILE = SCRIPT_DIR / "devices.json"
SNAPSHOT_FILE = SCRIPT_DIR / "snapshot.json"
CLOUD_CONFIG_FILES = [SCRIPT_DIR / "cloud.json", SCRIPT_DIR / "tinytuya.json"]

logger = logging.getLogger(__name__)

# ── DPS mappings (Tuya Data Point Schema) ────────────────────────────────────
# DPS 20 = on/off for most Tuya light bulbs
# DPS 1  = on/off for most Tuya plugs/switches
DPS_SWITCH_BULB = "20"
DPS_SWITCH_PLUG = "1"


def _load_devices(path: Path) -> List[Dict[str, Any]]:
    """Load device list from JSON file.

    Accepts snapshot format ``{"timestamp": ..., "devices": [...]}``
    and legacy plain ``[...]`` array.
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


def _has_valid_ip(device_info: Dict[str, Any]) -> bool:
    """True if the device has a non-empty IP that looks like a local address (not 'Auto' or error)."""
    ip = device_info.get("ip") or ""
    if not ip or not isinstance(ip, str):
        return False
    ip = ip.strip()
    if ip in ("Auto", "") or "No IP" in ip or "Error" in ip:
        return False
    parts = ip.split(".")
    return len(parts) == 4 and all(p.isdigit() and 0 <= int(p) <= 255 for p in parts)


def _find_device(devices: List[Dict[str, Any]], query: str) -> Dict[str, Any]:
    """Find a device by name (case-insensitive partial match) or by ID.
    When multiple devices share the same name, prefer the one with a valid IP.
    """
    query_lower = query.lower()
    for dev in devices:
        if dev.get("id", "").lower() == query_lower:
            return dev
    name_matches = [d for d in devices if query_lower in d.get("name", "").lower()]
    if not name_matches:
        logger.error(f"❌ No device matching '{query}' found in devices.json")
        logger.info("ℹ️  Available devices:")
        for dev in devices:
            logger.info(f"   • {dev.get('name', '?'):30s}  (id={dev.get('id', '?')})")
        sys.exit(1)
    with_ip = [d for d in name_matches if _has_valid_ip(d)]
    return with_ip[0] if with_ip else name_matches[0]


def _get_version(device_info: Dict[str, Any]) -> float:
    """Extract protocol version, accepting both 'version' and 'ver' keys."""
    return float(device_info.get("version", device_info.get("ver", "3.3")))


def _get_cloud() -> Optional[tinytuya.Cloud]:
    """Load Tuya Cloud client from smart_life/cloud.json or tinytuya.json if present."""
    for path in CLOUD_CONFIG_FILES:
        if not path.exists():
            continue
        try:
            with open(path, "r", encoding="utf-8") as fh:
                config = json.load(fh)
            api_region = config.get("apiRegion") or config.get("api_region")
            api_key = config.get("apiKey") or config.get("api_key")
            api_secret = config.get("apiSecret") or config.get("api_secret")
            if api_key and api_secret and api_region:
                return tinytuya.Cloud(api_region, api_key, api_secret)
        except Exception as e:
            logger.debug("Could not load cloud config from %s: %s", path, e)
    return None


def _cloud_status_to_dps(cloud_response: Dict[str, Any]) -> Optional[Tuple[Dict[str, Any], str]]:
    """Convert Cloud getstatus() result to (status_dict with 'dps', switch_code).
    Cloud returns result: [{"code": "switch_1", "value": true}, ...].
    Returns ({"dps": {"1": True}}, "switch_1") or None on failure.
    """
    if not cloud_response.get("success") or "result" not in cloud_response:
        return None
    items = cloud_response.get("result") or []
    if not isinstance(items, list):
        return None
    dps = {}
    switch_code = None
    for item in items:
        if not isinstance(item, dict):
            continue
        code = item.get("code") or ""
        value = item.get("value")
        if code in ("switch_1", "switch", "switch_led"):
            switch_code = code
            dps["1" if code in ("switch_1", "switch") else "20"] = value
    if switch_code is None:
        return None
    return ({"dps": dps}, switch_code)


def _connect(
    device_info: Dict[str, Any],
    *,
    socket_timeout: float = 1.0,
    retry_limit: int = 1,
) -> tinytuya.Device:
    """Create a tinytuya Device connection (1s timeout, 1 retry for local attempt)."""
    dev_id = device_info["id"]
    ip = device_info.get("ip", "Auto")
    local_key = device_info["key"]
    version = _get_version(device_info)

    dev = tinytuya.Device(dev_id, ip, local_key, version=version)
    dev.set_socketPersistent(False)
    dev.set_socketTimeout(socket_timeout)
    dev.set_socketRetryLimit(retry_limit)
    dev.set_sendWait(1)
    return dev


def _format_status_error(status: Dict[str, Any]) -> str:
    """Turn Tuya status error dict into a short readable message."""
    err = status.get("Err", "")
    msg = status.get("Error", "")
    if err or msg:
        return f"Err={err!r}  Error={msg!r}"
    return str(status)


def _connect_and_status(
    device_info: Dict[str, Any],
) -> Tuple[Optional[tinytuya.Device], str, Dict[str, Any], Optional[Dict[str, Any]]]:
    """Connect to device and fetch status. Prefer Cloud (fast, reliable); if not configured or fails, try local (1s)."""
    name = device_info.get("name", device_info["id"])

    # Cloud first — fast and reliable when device is cloud-linked
    cloud = _get_cloud()
    if cloud:
        device_id = device_info["id"]
        resp = cloud.getstatus(device_id)
        parsed = _cloud_status_to_dps(resp)
        if parsed:
            status_dict, switch_code = parsed
            return None, name, status_dict, {"cloud": cloud, "device_id": device_id, "switch_code": switch_code}

    # No cloud or cloud failed — try local with 1s timeout so we don't wait long
    dev = _connect(device_info)
    status = dev.status()
    if "Error" not in str(status):
        return dev, name, status, None

    logger.error(f"❌ Cannot reach '{name}': {_format_status_error(status)}")
    logger.info("ℹ️  Run 'python light_control.py update' to refresh IPs; add smart_life/cloud.json for cloud control.")
    sys.exit(1)


def _detect_switch_dps(status: Dict[str, Any]) -> str:
    """Auto-detect whether the device uses DPS 20 (bulb) or DPS 1 (plug)."""
    dps = status.get("dps", {})
    if DPS_SWITCH_BULB in dps:
        return DPS_SWITCH_BULB
    if DPS_SWITCH_PLUG in dps:
        return DPS_SWITCH_PLUG
    logger.warning(f"⚠️  Could not auto-detect switch DPS. Raw DPS: {dps}")
    return DPS_SWITCH_PLUG


def _set_switch_via_cloud(cloud_ctx: Dict[str, Any], value: bool) -> bool:
    """Send switch command via Tuya Cloud. Returns True if success."""
    body = {"commands": [{"code": cloud_ctx["switch_code"], "value": value}]}
    resp = cloud_ctx["cloud"].sendcommand(cloud_ctx["device_id"], body)
    return bool(resp and resp.get("success"))


def turn_on(device_info: Dict[str, Any]) -> None:
    """Turn a light/device ON."""
    dev, name, status, cloud_ctx = _connect_and_status(device_info)
    if cloud_ctx:
        if not _set_switch_via_cloud(cloud_ctx, True):
            logger.error("❌ Cloud command failed for '%s'", name)
            sys.exit(1)
    else:
        dps_key = _detect_switch_dps(status)
        dev.set_value(dps_key, True)
    logger.info(f"✅ '{name}' turned ON")


def turn_off(device_info: Dict[str, Any]) -> None:
    """Turn a light/device OFF."""
    dev, name, status, cloud_ctx = _connect_and_status(device_info)
    if cloud_ctx:
        if not _set_switch_via_cloud(cloud_ctx, False):
            logger.error("❌ Cloud command failed for '%s'", name)
            sys.exit(1)
    else:
        dps_key = _detect_switch_dps(status)
        dev.set_value(dps_key, False)
    logger.info(f"✅ '{name}' turned OFF")


def get_status(device_info: Dict[str, Any]) -> None:
    """Print current device status."""
    _dev, name, status, _cloud_ctx = _connect_and_status(device_info)
    dps = status.get("dps", {})
    dps_key = _detect_switch_dps(status)
    is_on = dps.get(dps_key, None)
    state_label = "ON" if is_on else "OFF" if is_on is not None else "UNKNOWN"

    logger.info(f"💡 '{name}' is {state_label}")
    logger.debug(f"📂 Full DPS: {json.dumps(dps, indent=2)}")


def switch_toggle(device_info: Dict[str, Any]) -> None:
    """Toggle device: if ON turn OFF, if OFF turn ON."""
    dev, name, status, cloud_ctx = _connect_and_status(device_info)
    dps = status.get("dps", {})
    dps_key = _detect_switch_dps(status)
    is_on = dps.get(dps_key, False)
    new_state = not is_on
    if cloud_ctx:
        if not _set_switch_via_cloud(cloud_ctx, new_state):
            logger.error("❌ Cloud command failed for '%s'", name)
            sys.exit(1)
    else:
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
    """Run tinytuya network scan to discover devices (via CLI)."""
    logger.info("🔍 Scanning local network for Tuya devices …")
    subprocess.run(
        [sys.executable, "-m", "tinytuya", "scan"],
        check=False,
        creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0,
    )


def update_devices(devices_path: Path, snapshot_path: Path) -> None:
    """Run tinytuya snapshot (using devices.json) then overwrite devices.json with result.

    Snapshot discovers current IPs and DPS for each device in devices.json;
    writing the result back keeps devices.json in sync with snapshot format.
    Tinytuya expects the device file to be a JSON array of devices, so we write
    a temp file with just the devices list when our file is {timestamp, devices}.
    """
    with open(devices_path, "r", encoding="utf-8") as fh:
        data = json.load(fh)
    devices_list = data.get("devices", data) if isinstance(data, dict) else data
    if not isinstance(devices_list, list):
        logger.error("❌ devices.json must be a list of devices or {timestamp, devices}")
        sys.exit(1)

    logger.info("🔄 Running tinytuya snapshot to refresh IPs and status …")
    with tempfile.NamedTemporaryFile(
        mode="w",
        suffix=".json",
        delete=False,
        encoding="utf-8",
    ) as tmp:
        json.dump(devices_list, tmp, indent=2)
        tmp_path = tmp.name
    try:
        result = subprocess.run(
            [
                sys.executable,
                "-m",
                "tinytuya",
                "snapshot",
                "-y",
                "-device-file",
                tmp_path,
                "-snapshot-file",
                str(snapshot_path),
            ],
            cwd=str(devices_path.parent),
            check=False,
            creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0,
        )
    finally:
        Path(tmp_path).unlink(missing_ok=True)

    if result.returncode != 0:
        logger.error("❌ Snapshot failed; devices.json was not updated")
        sys.exit(1)
    if not snapshot_path.exists():
        logger.error("❌ Snapshot file was not created")
        sys.exit(1)
    with open(snapshot_path, "r", encoding="utf-8") as fh:
        snapshot = json.load(fh)
    devices = snapshot.get("devices", [])
    # Deduplicate by name: when multiple devices share a name, keep the one with a valid IP
    seen_names: Dict[str, Dict[str, Any]] = {}
    for dev in devices:
        name = (dev.get("name") or "").strip() or dev.get("id", "")
        if name not in seen_names:
            seen_names[name] = dev
        elif _has_valid_ip(dev) and not _has_valid_ip(seen_names[name]):
            seen_names[name] = dev
    snapshot["devices"] = list(seen_names.values())
    with open(devices_path, "w", encoding="utf-8") as fh:
        json.dump(snapshot, fh, indent=4)
    logger.info(f"✅ Updated {devices_path.name} with snapshot ({len(snapshot['devices'])} devices)")


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
            "  python light_control.py update\n"
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "action",
        choices=["on", "off", "switch", "status", "list", "scan", "update"],
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

    if args.action == "update":
        update_devices(DEVICES_FILE, SNAPSHOT_FILE)
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
