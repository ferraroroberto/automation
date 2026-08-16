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

Reading the three console tools
-------------------------------
`powershell`, `nvidia-smi` and `pnputil` are native Windows console tools:
they write their own console code page, so their output is captured as raw
bytes and decoded explicitly through `system/_lib/console_output.py` - never
with `text=` or `encoding=`, which decode with the *parent's* locale instead
(`#106`, `#108`). Every query here therefore answers in a tri-state:
`GpuQuery.known` and `SmiReading.known` say whether the fact was established
at all, so "could not read the GPU" is never reported as "the GPU is gone".

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
from pathlib import Path
from typing import NamedTuple, Optional, Sequence

# system/_lib holds the shared native-console-tool decoder (see #106/#108).
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "system"))
from _lib.console_output import NO_WINDOW as _NO_WINDOW  # noqa: E402
from _lib.console_output import decode_console_bytes  # noqa: E402

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


class ConsoleRun(NamedTuple):
    """Outcome of one native-console-tool invocation.

    `ran` is the state that must never be folded into a negative answer: a
    query that could not run, or whose output could not be read, is
    *unknown* - not "the GPU is gone".
    """

    ran: bool
    returncode: int
    stdout: str
    stderr: str
    encoding: Optional[str]

    @property
    def succeeded(self) -> bool:
        """True when the tool ran and reported success."""
        return self.ran and self.returncode == 0

    @property
    def has_output(self) -> bool:
        """True when stdout carries something to parse."""
        return bool(self.stdout.strip())


def _run_console(command: Sequence[str], *, timeout: int, label: str) -> ConsoleRun:
    """Run a native Windows console tool and decode its output explicitly.

    Captures raw bytes - never `text=`, never `encoding=` - because those
    decode with the *parent's* locale rather than the child's console code
    page. See `system/_lib/console_output.py` for the measured failure mode
    (`#106`, `#108`); the short version is that under `PYTHONUTF8=1` a bad
    decode empties the whole output while still reporting exit 0, which is
    indistinguishable from "the tool answered nothing".

    Never raises. Every failure path logs, so a `ConsoleRun` that could not
    be read leaves a breadcrumb instead of a silent negative.
    """
    try:
        proc = subprocess.run(
            command,
            capture_output=True,
            timeout=timeout,
            creationflags=_NO_WINDOW,
        )
    except (subprocess.SubprocessError, OSError) as exc:
        logger.warning("⚠️  %s could not run: %s", label, exc)
        return ConsoleRun(ran=False, returncode=-1, stdout="", stderr="", encoding=None)

    stdout, encoding = decode_console_bytes(proc.stdout or b"")
    stderr, _ = decode_console_bytes(proc.stderr or b"")

    if proc.returncode != 0:
        logger.debug(
            "ℹ️  %s exited %d: %s", label, proc.returncode, (stderr or stdout).strip()
        )
    elif (proc.stdout or b"") and not stdout.strip():
        # Bytes came back but decoded to nothing - the answer is unknown,
        # never an empty result. This is exactly the shape #108 is about.
        logger.error(
            "❌ %s exited 0 but produced no readable output "
            "(%d raw bytes, decoded as %s) - result is unknown, not empty.",
            label,
            len(proc.stdout or b""),
            encoding,
        )

    return ConsoleRun(
        ran=True,
        returncode=proc.returncode,
        stdout=stdout,
        stderr=stderr,
        encoding=encoding,
    )


def _powershell(script: str, label: str) -> ConsoleRun:
    """Run a Windows PowerShell snippet and return its decoded result.

    Returns a `ConsoleRun` rather than a bare string so callers can tell
    "PowerShell answered nothing" apart from "PowerShell could not be read";
    an empty string used to mean both, silently.
    """
    run = _run_console(
        [POWERSHELL, "-NoProfile", "-NonInteractive", "-Command", script],
        timeout=60,
        label=f"PowerShell ({label})",
    )
    if not run.ran:
        return run
    if run.returncode != 0:
        logger.warning(
            "⚠️  PowerShell (%s) exited %d: %s",
            label,
            run.returncode,
            (run.stderr or run.stdout).strip() or "<no output>",
        )
    return run


class GpuQuery(NamedTuple):
    """Tri-state answer to "which NVIDIA display adapters are present?".

    `known` is False when the PowerShell query itself failed or could not be
    parsed, in which case an empty `gpus` means nothing - it must not be
    reported as "the card is physically absent".
    """

    gpus: list[dict]
    known: bool


def find_nvidia_gpus() -> GpuQuery:
    """Discover present NVIDIA display adapters and their problem state.

    Returns a `GpuQuery` whose `gpus` are dicts of
    {FriendlyName, InstanceId, Status, ConfigManagerErrorCode}. The instance
    id is discovered dynamically so the tool is portable across
    machines/cards - nothing is hardcoded.
    """
    script = (
        "Get-PnpDevice -PresentOnly -Class Display | "
        f"Where-Object {{ $_.InstanceId -like '{NVIDIA_VENDOR_PREFIX}*' }} | "
        "Select-Object FriendlyName, InstanceId, Status, ConfigManagerErrorCode | "
        "ConvertTo-Json -Compress"
    )
    run = _powershell(script, "enumerate display adapters")
    if not run.succeeded:
        return GpuQuery(gpus=[], known=False)
    raw = run.stdout.strip()
    if not raw:
        # ConvertTo-Json emits nothing when the pipeline is empty, so a
        # zero-exit empty stdout genuinely means "no NVIDIA adapter".
        return GpuQuery(gpus=[], known=True)
    try:
        data = json.loads(raw)
    except json.JSONDecodeError as exc:
        logger.error(
            "❌ Could not parse the adapter query output as JSON (%s, decoded as %s): %s",
            exc,
            run.encoding,
            raw[:200],
        )
        return GpuQuery(gpus=[], known=False)
    # ConvertTo-Json emits a bare object for a single item, a list for many.
    return GpuQuery(gpus=data if isinstance(data, list) else [data], known=True)


def read_config_flags(instance_id: str) -> Optional[int]:
    """Read the device's registry ConfigFlags (0 = enabled, bit 0x1 = disabled)."""
    safe = instance_id.replace("'", "''")
    script = (
        "$k = Get-ItemProperty "
        f"'HKLM:\\SYSTEM\\CurrentControlSet\\Enum\\{safe}' "
        "-Name ConfigFlags -ErrorAction SilentlyContinue; "
        "if ($null -ne $k) { $k.ConfigFlags } else { 'NA' }"
    )
    run = _powershell(script, "read ConfigFlags")
    try:
        return int(run.stdout.strip())
    except (TypeError, ValueError):
        return None


class SmiReading(NamedTuple):
    """Tri-state answer to "is the GPU live?".

    `known` is False when nvidia-smi could not be run or its output could not
    be read. That case must never be shown as "the GPU isn't live" - it is
    the single alarm condition this tool exists to detect, and reporting a
    healthy card as gone is worse than reporting nothing at all.
    """

    summary: Optional[str]
    known: bool


def _smi_command() -> list[str]:
    """Build the nvidia-smi argv (seam: the test swaps in a stand-in child)."""
    smi = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "System32", "nvidia-smi.exe")
    exe = smi if os.path.exists(smi) else "nvidia-smi"
    return [
        exe,
        "--query-gpu=name,driver_version,memory.total,temperature.gpu,utilization.gpu",
        "--format=csv,noheader",
    ]


def nvidia_smi() -> SmiReading:
    """Return a one-line nvidia-smi summary, distinguishing dead from unreadable."""
    run = _run_console(_smi_command(), timeout=30, label="nvidia-smi")
    if not run.ran:
        return SmiReading(summary=None, known=False)
    if run.returncode != 0:
        # nvidia-smi ran and refused: the GPU genuinely is not live.
        logger.debug("ℹ️  nvidia-smi exited %d - GPU not responding.", run.returncode)
        return SmiReading(summary=None, known=True)
    if not run.has_output:
        logger.error(
            "❌ nvidia-smi exited 0 with no readable output - the GPU state is "
            "UNKNOWN, not 'not live'."
        )
        return SmiReading(summary=None, known=False)
    return SmiReading(summary=run.stdout.strip(), known=True)


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
    run = _powershell(script, "recent device events")
    return [line for line in run.stdout.splitlines() if line.strip()]


def enable_device(instance_id: str) -> bool:
    """Enable a disabled device via native pnputil (the call that actually works)."""
    if not is_admin():
        logger.error("❌ Enabling the device needs administrator rights. "
                     "Run gpu_recovery.bat (it self-elevates).")
        return False
    logger.info("🔧 Enabling via: pnputil /enable-device \"%s\"", instance_id)
    run = _run_console(
        [PNPUTIL, "/enable-device", instance_id], timeout=60, label="pnputil /enable-device"
    )
    if not run.ran:
        logger.error("❌ pnputil could not be run - the device was not enabled.")
        return False
    evidence = run.stdout.strip() or run.stderr.strip()
    if evidence:
        logger.info(evidence)
    else:
        logger.warning(
            "⚠️  pnputil exited %d but said nothing readable - judging the enable "
            "by its exit code alone.", run.returncode
        )
    return run.returncode == 0


def describe_code(code: Optional[int]) -> str:
    if code is None:
        return "unknown"
    return CODE_MEANINGS.get(code, f"Code {code} (see Microsoft Device Manager error codes)")


def diagnose() -> GpuQuery:
    """Print a diagnostic report and return the discovered GPU records."""
    logger.info("=" * 70)
    logger.info("NVIDIA GPU diagnostic")
    logger.info("=" * 70)

    query = find_nvidia_gpus()
    gpus = query.gpus
    if not query.known:
        logger.error("❌ Could not enumerate display adapters - the GPU state is "
                     "UNKNOWN. This is not the same as 'no card found'; see the "
                     "error above for why the query could not be read.")
        return query
    if not gpus:
        logger.warning("⚠️  No present NVIDIA display adapter found. The card may "
                       "be physically absent, or not enumerating on the PCIe bus.")
        return query

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

    reading = nvidia_smi()
    logger.info("")
    if reading.summary:
        logger.info("✅ nvidia-smi: %s", reading.summary)
    elif reading.known:
        logger.info("ℹ️  nvidia-smi: no response (expected while the GPU is disabled).")
    else:
        logger.warning("⚠️  nvidia-smi: could not be read - GPU liveness is UNKNOWN, "
                       "not 'not live'.")

    return query


def main() -> int:
    parser = argparse.ArgumentParser(description="Diagnose/recover a disabled NVIDIA GPU (Code 22).")
    parser.add_argument("--diagnose", action="store_true",
                        help="Report only; make no changes.")
    args = parser.parse_args()

    query = diagnose()
    if not query.known or not query.gpus:
        return 1

    disabled = [g for g in query.gpus if g.get("ConfigManagerErrorCode") == 22]

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
    still_bad = [g for g in post.gpus if g.get("ConfigManagerErrorCode") == 22]
    reading = nvidia_smi()

    if post.known and not still_bad and reading.summary:
        logger.info("✅ Recovered. nvidia-smi: %s", reading.summary)
        return 0

    all_ok = False
    if not post.known or not reading.known:
        logger.warning("⚠️  Could not confirm the GPU's state after the enable - "
                       "recovery is UNKNOWN, not failed. Re-run --diagnose.")
    else:
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
