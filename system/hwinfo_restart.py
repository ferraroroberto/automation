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
