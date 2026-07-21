"""
GPU Recovery - Diagnose and recover an NVIDIA GPU that Windows has marked
"disabled" (Device Manager Code 22).

Why this exists
---------------
On this machine an NVIDIA RTX 5060 Ti has intermittently ended up *disabled*
(Device Manager problem Code 22) with no crash, TDR, or driver fault in the
event log. When that happens Windows falls back to the "Microsoft Basic
Render Driver", `nvidia-smi` stops responding, and every NVENC-based tool in
this `video/` folder silently loses GPU acceleration.

The non-obvious trap: the PowerShell `Enable-PnpDevice` cmdlet *reports
success but does not actually start the device* - the card stays at Code 22
through enable/disable cycles, parent-port cycles, and even a reboot. The
native `pnputil /enable-device` performs the real enable in one shot.

This tool diagnoses the GPU state and, if it finds the card disabled, applies
the working fix (`pnputil /enable-device`) and verifies recovery with
`nvidia-smi`. Enabling a device needs administrator rights; `gpu_recovery.bat`
self-elevates before launching this script.

Usage
-----
    python gpu_recovery.py            # diagnose, then fix if disabled
    python gpu_recovery.py --diagnose # report only, make no changes
"""

import argparse
import ctypes
import json
import logging
import os
import subprocess
import sys
from typing import Optional

# UTF-8 stdout so emoji/box characters survive redirected/captured runs on
# Windows (cp1252 fallback otherwise throws UnicodeEncodeError under capture).
try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except (AttributeError, ValueError):
    pass

logging.basicConfig(level=logging.INFO, format="%(message)s")
logger = logging.getLogger("gpu_recovery")

POWERSHELL = r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
PNPUTIL = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "System32", "pnputil.exe")
NVIDIA_VENDOR_PREFIX = "PCI\\VEN_10DE"

# Suppress the console window each spawn would otherwise flash when this
# tool runs unattended (e.g. relaunched via gpu_recovery.bat self-elevation).
_NO_WINDOW = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0

# Device Manager ConfigManagerErrorCode -> human meaning (the ones we care about).
CODE_MEANINGS = {
    0: "OK - working normally",
    10: "Code 10 - device cannot start (driver/init failure)",
    22: "Code 22 - device is DISABLED",
    28: "Code 28 - drivers not installed",
    31: "Code 31 - device not working properly (driver problem)",
    43: "Code 43 - Windows stopped it after a reported fault (fell off the bus)",
}


def is_admin() -> bool:
    """True if the current process holds administrator rights."""
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except OSError:
        return False


def _powershell(script: str) -> str:
    """Run a Windows PowerShell snippet and return its stdout (empty on error)."""
    try:
        result = subprocess.run(
            [POWERSHELL, "-NoProfile", "-NonInteractive", "-Command", script],
            capture_output=True,
            text=True,
            timeout=60,
            creationflags=_NO_WINDOW,
        )
        return result.stdout.strip()
    except (subprocess.SubprocessError, OSError) as exc:
        logger.warning("⚠️  PowerShell call failed: %s", exc)
        return ""


def find_nvidia_gpus() -> list[dict]:
    """Discover present NVIDIA display adapters and their problem state.

    Returns a list of dicts: {FriendlyName, InstanceId, Status, ConfigManagerErrorCode}.
    The instance id is discovered dynamically so the tool is portable across
    machines/cards - nothing is hardcoded.
    """
    script = (
        "Get-PnpDevice -PresentOnly -Class Display | "
        f"Where-Object {{ $_.InstanceId -like '{NVIDIA_VENDOR_PREFIX}*' }} | "
        "Select-Object FriendlyName, InstanceId, Status, ConfigManagerErrorCode | "
        "ConvertTo-Json -Compress"
    )
    raw = _powershell(script)
    if not raw:
        return []
    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        return []
    # ConvertTo-Json emits a bare object for a single item, a list for many.
    return data if isinstance(data, list) else [data]


def read_config_flags(instance_id: str) -> Optional[int]:
    """Read the device's registry ConfigFlags (0 = enabled, bit 0x1 = disabled)."""
    safe = instance_id.replace("'", "''")
    script = (
        "$k = Get-ItemProperty "
        f"'HKLM:\\SYSTEM\\CurrentControlSet\\Enum\\{safe}' "
        "-Name ConfigFlags -ErrorAction SilentlyContinue; "
        "if ($null -ne $k) { $k.ConfigFlags } else { 'NA' }"
    )
    raw = _powershell(script)
    try:
        return int(raw)
    except (TypeError, ValueError):
        return None


def nvidia_smi() -> Optional[str]:
    """Return a one-line nvidia-smi summary, or None if the GPU isn't live."""
    smi = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "System32", "nvidia-smi.exe")
    exe = smi if os.path.exists(smi) else "nvidia-smi"
    try:
        result = subprocess.run(
            [
                exe,
                "--query-gpu=name,driver_version,memory.total,temperature.gpu,utilization.gpu",
                "--format=csv,noheader",
            ],
            capture_output=True,
            text=True,
            timeout=30,
            creationflags=_NO_WINDOW,
        )
        if result.returncode == 0 and result.stdout.strip():
            return result.stdout.strip()
    except (subprocess.SubprocessError, OSError):
        pass
    return None


def recent_events(instance_id: str) -> list[str]:
    """Best-effort: recent System-log lines referencing this device (last 2 days).

    Helps answer the open question of *what* disabled the card - a manual
    Device Manager click or a utility (e.g. MSI Center) leaves a trail here.
    """
    # Match on the VEN/DEV core; the full instance id rarely appears verbatim.
    core = instance_id.split("\\")[1] if "\\" in instance_id else instance_id
    script = (
        "Get-WinEvent -FilterHashtable @{LogName='System'; "
        "StartTime=(Get-Date).AddDays(-2)} -ErrorAction SilentlyContinue | "
        "Where-Object { $_.ProviderName -match 'Kernel-PnP|nvlddmkm' -or "
        f"$_.Message -match '{core}|RTX 50|nvlddmkm' }} | "
        "Select-Object -First 12 TimeCreated, Id, ProviderName | "
        "ForEach-Object { '{0}  [{1}] {2}' -f $_.TimeCreated, $_.Id, $_.ProviderName }"
    )
    raw = _powershell(script)
    return [line for line in raw.splitlines() if line.strip()]


def enable_device(instance_id: str) -> bool:
    """Enable a disabled device via native pnputil (the call that actually works)."""
    if not is_admin():
        logger.error("❌ Enabling the device needs administrator rights. "
                     "Run gpu_recovery.bat (it self-elevates).")
        return False
    logger.info("🔧 Enabling via: pnputil /enable-device \"%s\"", instance_id)
    try:
        result = subprocess.run(
            [PNPUTIL, "/enable-device", instance_id],
            capture_output=True,
            text=True,
            timeout=60,
            creationflags=_NO_WINDOW,
        )
    except (subprocess.SubprocessError, OSError) as exc:
        logger.error("❌ pnputil failed: %s", exc)
        return False
    logger.info(result.stdout.strip() or result.stderr.strip())
    return result.returncode == 0


def describe_code(code: Optional[int]) -> str:
    if code is None:
        return "unknown"
    return CODE_MEANINGS.get(code, f"Code {code} (see Microsoft Device Manager error codes)")


def diagnose() -> list[dict]:
    """Print a diagnostic report and return the discovered GPU records."""
    logger.info("=" * 70)
    logger.info("NVIDIA GPU diagnostic")
    logger.info("=" * 70)

    gpus = find_nvidia_gpus()
    if not gpus:
        logger.warning("⚠️  No present NVIDIA display adapter found. The card may "
                       "be physically absent, or not enumerating on the PCIe bus.")
        return gpus

    for gpu in gpus:
        name = gpu.get("FriendlyName", "NVIDIA GPU")
        instance = gpu.get("InstanceId", "")
        code = gpu.get("ConfigManagerErrorCode")
        flags = read_config_flags(instance)
        logger.info("")
        logger.info("• %s", name)
        logger.info("    Instance     : %s", instance)
        logger.info("    State        : %s", describe_code(code))
        logger.info("    ConfigFlags  : %s%s", flags,
                    "  (0 = enabled in registry)" if flags == 0 else "")

    smi = nvidia_smi()
    logger.info("")
    if smi:
        logger.info("✅ nvidia-smi: %s", smi)
    else:
        logger.info("ℹ️  nvidia-smi: no response (expected while the GPU is disabled).")

    return gpus


def main() -> int:
    parser = argparse.ArgumentParser(description="Diagnose/recover a disabled NVIDIA GPU (Code 22).")
    parser.add_argument("--diagnose", action="store_true",
                        help="Report only; make no changes.")
    args = parser.parse_args()

    gpus = diagnose()
    if not gpus:
        return 1

    disabled = [g for g in gpus if g.get("ConfigManagerErrorCode") == 22]

    if not disabled:
        logger.info("")
        logger.info("✅ Nothing to fix - no NVIDIA GPU is in the Code 22 disabled state.")
        return 0

    if args.diagnose:
        logger.info("")
        logger.info("⚠️  %d GPU(s) disabled. Re-run without --diagnose to fix.", len(disabled))
        return 2

    logger.info("")
    logger.info("Found %d disabled GPU(s). Applying fix...", len(disabled))
    all_ok = True
    for gpu in disabled:
        instance = gpu.get("InstanceId", "")
        if not enable_device(instance):
            all_ok = False

    # Verify recovery.
    logger.info("")
    logger.info("Verifying...")
    post = find_nvidia_gpus()
    still_bad = [g for g in post if g.get("ConfigManagerErrorCode") == 22]
    smi = nvidia_smi()

    if not still_bad and smi:
        logger.info("✅ Recovered. nvidia-smi: %s", smi)
        return 0

    all_ok = False
    logger.warning("⚠️  GPU still not fully live after the enable.")
    if still_bad:
        logger.warning("    Still Code 22 - something may be actively re-disabling it.")
        hints = recent_events(still_bad[0].get("InstanceId", ""))
        if hints:
            logger.warning("    Recent System-log events touching the device:")
            for line in hints:
                logger.warning("      %s", line)
            logger.warning("    Suspect a utility (e.g. MSI Center) or a scheduled task.")
    return 0 if all_ok else 1


if __name__ == "__main__":
    sys.exit(main())
