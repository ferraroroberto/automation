"""Whisper-server process manager — mirrors `claude-local-calls/src/llama_process.py`.

One singleton process per project, keyed by the sibling `whisper_server.yaml`.
The port is fixed, so the OS guarantees mutual exclusion: if another project
(or a manual run) already holds the port, `start()` returns cleanly with
`already running (external)` and ownership is reported as EXTERNAL — the
caller must not kill it on exit.
"""

from __future__ import annotations

# Standard library imports
import logging
import os
import signal
import socket
import subprocess
import sys
import threading
import time
from collections import deque
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Deque, Dict, List, Optional

# Third-party imports
import requests
import yaml

logger = logging.getLogger(__name__)

OWNERSHIP_NONE = "none"          # server not running
OWNERSHIP_OURS = "ours"          # we started it; we kill on exit
OWNERSHIP_EXTERNAL = "external"  # someone else started it; hands off


@dataclass(frozen=True)
class ServerConfig:
    host: str
    bind_host: str
    port: int
    binary_path: Path
    model_path: Path
    args: List[str]
    pid_file: Path
    log_ring_size: int
    startup_timeout_seconds: float
    poll_interval_seconds: float
    request_timeout_seconds: float
    project_root: Path

    @property
    def base_url(self) -> str:
        return f"http://{self.host}:{self.port}"


@dataclass
class ServerStatus:
    running: bool
    ownership: str  # OWNERSHIP_* constant
    pid: Optional[int] = None
    port: Optional[int] = None
    base_url: Optional[str] = None
    detail: str = ""


def _resolve_binary(project_root: Path, raw: str) -> Path:
    p = (project_root / raw).resolve()
    if p.exists():
        return p
    if sys.platform == "win32" and not raw.endswith(".exe"):
        alt = p.with_suffix(".exe")
        if alt.exists():
            return alt
    return p  # return the non-existent path; caller decides


def load_config(config_path: Optional[Path] = None) -> ServerConfig:
    """Load `whisper_server.yaml` from next to this file (or an override).

    The project root is the directory containing the `whisper_server/` folder.
    """
    if config_path is None:
        config_path = Path(__file__).resolve().parent / "whisper_server.yaml"
    else:
        config_path = Path(config_path).resolve()

    project_root = Path(__file__).resolve().parent.parent
    raw: Dict[str, Any] = yaml.safe_load(config_path.read_text(encoding="utf-8")) or {}

    server = raw.get("server") or {}
    binary = raw.get("binary") or {}
    model = raw.get("model") or {}
    health = raw.get("health") or {}

    return ServerConfig(
        host=str(server.get("host", "127.0.0.1")),
        bind_host=str(server.get("bind_host", "0.0.0.0")),
        port=int(server.get("port", 8090)),
        binary_path=_resolve_binary(project_root, str(binary.get("path", "vendor/whisper.cpp/whisper-server"))),
        model_path=(project_root / str(model.get("path", "vendor/whisper.cpp/models/ggml-small.bin"))).resolve(),
        args=list(raw.get("args", []) or []),
        pid_file=(project_root / str(raw.get("pid_file", ".whisper_server.pid"))).resolve(),
        log_ring_size=int(raw.get("log_ring_size", 1000)),
        startup_timeout_seconds=float(health.get("startup_timeout_seconds", 60)),
        poll_interval_seconds=float(health.get("poll_interval_seconds", 0.5)),
        request_timeout_seconds=float(health.get("request_timeout_seconds", 1.5)),
        project_root=project_root,
    )


class WhisperServerManager:
    """Start / stop / health-check a local whisper.cpp server."""

    def __init__(self, config: Optional[ServerConfig] = None) -> None:
        self.config: ServerConfig = config or load_config()
        self._proc: Optional[subprocess.Popen] = None
        self._log: Deque[str] = deque(maxlen=self.config.log_ring_size)
        self._lock = threading.Lock()
        self._reader: Optional[threading.Thread] = None

    # ------------------------------------------------------------------ status

    def is_reachable(self) -> bool:
        """HTTP health check — whisper.cpp server answers 200 on `/`."""
        url = self.config.base_url + "/"
        try:
            r = requests.get(url, timeout=self.config.request_timeout_seconds)
            return r.status_code == 200
        except requests.RequestException:
            return False

    def is_port_in_use(self) -> bool:
        """Low-level port probe — works even if HTTP is not yet listening."""
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.settimeout(0.2)
            return s.connect_ex((self.config.host, self.config.port)) == 0

    def status(self) -> ServerStatus:
        running_here = self._proc is not None and self._proc.poll() is None
        reachable = self.is_reachable() or self.is_port_in_use()

        if running_here and reachable:
            return ServerStatus(
                running=True,
                ownership=OWNERSHIP_OURS,
                pid=self._proc.pid,
                port=self.config.port,
                base_url=self.config.base_url,
                detail="running (started by this process)",
            )

        if reachable:
            pid = self._read_pid_file()
            return ServerStatus(
                running=True,
                ownership=OWNERSHIP_EXTERNAL,
                pid=pid,
                port=self.config.port,
                base_url=self.config.base_url,
                detail="running (external — started elsewhere)",
            )

        return ServerStatus(
            running=False,
            ownership=OWNERSHIP_NONE,
            port=self.config.port,
            base_url=self.config.base_url,
            detail="not running",
        )

    # ------------------------------------------------------------------- start

    def start(self, wait: bool = True) -> ServerStatus:
        """Start the server. Idempotent — returns current status if already up."""
        current = self.status()
        if current.running:
            logger.info(f"ℹ️  Whisper server already {current.detail}")
            return current

        self._validate_paths()

        cmd = self._build_command()
        logger.info(f"🚀 Starting whisper-server: {' '.join(str(c) for c in cmd)}")

        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"
        env["PYTHONUTF8"] = "1"
        if sys.platform == "win32":
            # Let the binary find sibling CUDA DLLs.
            env["PATH"] = str(self.config.binary_path.parent) + os.pathsep + env.get("PATH", "")

        try:
            popen_kwargs: Dict[str, Any] = dict(
                cwd=str(self.config.project_root),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                bufsize=1,
                env=env,
            )
            if sys.platform == "win32":
                popen_kwargs["creationflags"] = subprocess.CREATE_NEW_PROCESS_GROUP
            self._proc = subprocess.Popen(cmd, **popen_kwargs)
        except FileNotFoundError as e:
            raise RuntimeError(
                f"❌ whisper-server binary not found: {self.config.binary_path}"
            ) from e
        except Exception as e:
            raise RuntimeError(f"❌ failed to launch whisper-server: {e}") from e

        self._write_pid_file(self._proc.pid)
        self._reader = threading.Thread(
            target=self._drain_output,
            args=(self._proc,),
            daemon=True,
        )
        self._reader.start()

        if wait:
            self._wait_until_ready()

        return self.status()

    # -------------------------------------------------------------------- stop

    def stop(self) -> ServerStatus:
        """Stop the server we started. Never touches an EXTERNAL server."""
        status = self.status()
        if status.ownership == OWNERSHIP_EXTERNAL:
            logger.info("✋ Leaving external whisper-server running (not ours)")
            return status
        if not status.running or self._proc is None:
            logger.info("ℹ️  Whisper server was not running")
            self._clear_pid_file()
            return ServerStatus(
                running=False,
                ownership=OWNERSHIP_NONE,
                port=self.config.port,
                base_url=self.config.base_url,
                detail="not running",
            )

        p = self._proc
        logger.info(f"🛑 Stopping whisper-server (pid={p.pid})")
        try:
            if sys.platform == "win32":
                try:
                    p.send_signal(signal.CTRL_BREAK_EVENT)
                except Exception as exc:
                    logger.debug(f"CTRL_BREAK_EVENT failed: {exc}")
            p.terminate()
            try:
                p.wait(timeout=8)
            except subprocess.TimeoutExpired:
                logger.warning("⚠️  whisper-server didn't exit; killing")
                p.kill()
                p.wait(timeout=5)
        finally:
            self._proc = None
            self._clear_pid_file()

        return ServerStatus(
            running=False,
            ownership=OWNERSHIP_NONE,
            port=self.config.port,
            base_url=self.config.base_url,
            detail="stopped",
        )

    # -------------------------------------------------------------- diagnostics

    def log_lines(self) -> List[str]:
        with self._lock:
            return list(self._log)

    # ------------------------------------------------------------------ helpers

    def _validate_paths(self) -> None:
        if not self.config.binary_path.exists():
            raise RuntimeError(
                f"❌ whisper-server binary not found at {self.config.binary_path}. "
                f"Build or install whisper.cpp into "
                f"{self.config.binary_path.parent.relative_to(self.config.project_root)}."
            )
        if not self.config.model_path.exists():
            raise RuntimeError(
                f"❌ whisper model file not found at {self.config.model_path}. "
                f"Download it with whisper.cpp's `download-ggml-model` script."
            )

    def _build_command(self) -> List[str]:
        cmd: List[str] = [
            str(self.config.binary_path),
            "--host", self.config.bind_host,
            "--port", str(self.config.port),
            "--model", str(self.config.model_path),
        ]
        cmd.extend(self.config.args)
        return cmd

    def _drain_output(self, proc: subprocess.Popen) -> None:
        if proc.stdout is None:
            return
        for raw in proc.stdout:
            line = raw.rstrip("\n")
            with self._lock:
                self._log.append(line)

    def _wait_until_ready(self) -> None:
        deadline = time.time() + self.config.startup_timeout_seconds
        while time.time() < deadline:
            if self._proc is None or self._proc.poll() is not None:
                tail = "\n".join(self.log_lines()[-20:])
                raise RuntimeError(
                    f"❌ whisper-server exited before becoming ready.\nLast output:\n{tail}"
                )
            if self.is_reachable():
                logger.info(f"✅ Whisper server ready at {self.config.base_url}")
                return
            time.sleep(self.config.poll_interval_seconds)
        raise RuntimeError(
            f"❌ whisper-server did not become ready within "
            f"{self.config.startup_timeout_seconds}s"
        )

    def _write_pid_file(self, pid: int) -> None:
        try:
            self.config.pid_file.write_text(str(pid), encoding="utf-8")
        except OSError as e:
            logger.warning(f"⚠️  Could not write PID file {self.config.pid_file}: {e}")

    def _clear_pid_file(self) -> None:
        try:
            if self.config.pid_file.exists():
                self.config.pid_file.unlink()
        except OSError as e:
            logger.warning(f"⚠️  Could not remove PID file {self.config.pid_file}: {e}")

    def _read_pid_file(self) -> Optional[int]:
        try:
            if self.config.pid_file.exists():
                return int(self.config.pid_file.read_text(encoding="utf-8").strip())
        except (OSError, ValueError):
            return None
        return None
