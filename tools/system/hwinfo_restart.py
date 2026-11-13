import logging
import os
import sys
import time

import psutil

log = logging.getLogger(__name__)

PROCESS_NAME = "HWiNFO64.exe"
INSTALL_PATHS = [
    r"C:\Program Files\HWiNFO64\HWiNFO64.exe",
    r"C:\Program Files (x86)\HWiNFO64\HWiNFO64.exe",
]
# The Elgato Stream Deck plugin's own hwinfo-plugin.exe and its parent hwinfo.exe.
PLUGIN_NAMES = ("hwinfo-plugin.exe", "hwinfo.exe")
PLUGIN_DIR = "com.exension.hwinfo.sdplugin"


def find_install_path() -> str:
    for path in INSTALL_PATHS:
        if os.path.isfile(path):
            return path
    raise FileNotFoundError(f"{PROCESS_NAME} not found in any of: {INSTALL_PATHS}")


def _matching_procs():
    return [
        proc
        for proc in psutil.process_iter(["name"])
        if (proc.info["name"] or "").lower() == PROCESS_NAME.lower()
    ]


def stop_running() -> None:
    # HWiNFO64 runs elevated for sensor/driver access, so from a non-elevated
    # process psutil can send TerminateProcess but cannot open a handle to
    # wait() or kill() it (AccessDenied) — poll process_iter instead.
    procs = _matching_procs()
    if not procs:
        log.info("%s was not running", PROCESS_NAME)
        return
    for proc in procs:
        log.info("Terminating %s (pid %d)", PROCESS_NAME, proc.pid)
        try:
            proc.terminate()
        except psutil.Error:
            log.warning("Could not signal %s (pid %d)", PROCESS_NAME, proc.pid, exc_info=True)

    deadline = time.monotonic() + 5
    while time.monotonic() < deadline and _matching_procs():
        time.sleep(0.5)

    for proc in _matching_procs():
        log.warning("%s (pid %d) did not exit gracefully, killing", PROCESS_NAME, proc.pid)
        try:
            proc.kill()
        except psutil.Error:
            log.warning("Could not kill %s (pid %d)", PROCESS_NAME, proc.pid, exc_info=True)


def _plugin_procs():
    # Matched by exact name AND an exe path inside the plugin folder, so neither HWiNFO64 nor any
    # other hwinfo.exe is touched; a process whose exe path can't be read (None) never matches.
    return [
        proc
        for proc in psutil.process_iter(["name", "exe"])
        if (proc.info["name"] or "").lower() in PLUGIN_NAMES
        and PLUGIN_DIR in (proc.info["exe"] or "").lower()
    ]


def stop_plugin() -> set:
    """Stop the Stream Deck HWiNFO plugin (it leaks memory); return the pids that were stopped."""
    procs = _plugin_procs()
    if not procs:
        log.info("Stream Deck HWiNFO plugin was not running")
        return set()
    for proc in procs:
        log.info("Terminating %s (pid %d)", proc.info["name"], proc.pid)
        try:
            proc.terminate()
        except psutil.Error:
            log.warning("Could not signal %s (pid %d)", proc.info["name"], proc.pid, exc_info=True)

    deadline = time.monotonic() + 5
    while time.monotonic() < deadline and _plugin_procs():
        time.sleep(0.5)

    for proc in _plugin_procs():
        log.warning("%s (pid %d) did not exit gracefully, killing", proc.info["name"], proc.pid)
        try:
            proc.kill()
        except psutil.Error:
            log.warning("Could not kill %s (pid %d)", proc.info["name"], proc.pid, exc_info=True)
    return {proc.pid for proc in procs}


def wait_plugin_relaunch(old_pids: set, timeout: float = 60, interval: float = 1) -> bool:
    """Stream Deck restarts the plugin on its own (about 15 s); True once a new one is running."""
    deadline = time.monotonic() + timeout
    while True:
        if any(
            proc.info["name"].lower() == PLUGIN_NAMES[0] and proc.pid not in old_pids
            for proc in _plugin_procs()
        ):
            log.info("plugin relaunched")
            return True
        if time.monotonic() >= deadline:
            log.warning("plugin relaunch not confirmed within %ds", timeout)
            return False
        time.sleep(interval)


def restart_plugin() -> None:
    stopped = stop_plugin()
    if stopped:
        wait_plugin_relaunch(stopped)


def restart() -> None:
    stop_running()
    install_path = find_install_path()
    log.info("Launching %s", install_path)
    # HWiNFO64.exe's requireAdministrator manifest only gets honored through
    # ShellExecute-style launches (os.startfile) — subprocess.Popen's plain
    # CreateProcess raises WinError 740 (elevation required) even from an
    # already-elevated caller (empirically verified: IsUserAnAdmin()==1 in
    # the calling process, subprocess.Popen still fails regardless of
    # close_fds, while os.startfile succeeds).
    os.startfile(install_path, cwd=os.path.dirname(install_path))
    restart_plugin()  # after HWiNFO64 is back, so the relaunched plugin finds its sensors


def main() -> None:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    try:
        restart()
    except Exception:
        log.exception("Failed to restart %s", PROCESS_NAME)
        sys.exit(1)
    log.info("%s restart complete", PROCESS_NAME)


if __name__ == "__main__":
    main()
