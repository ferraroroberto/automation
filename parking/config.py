"""Config (config.json) + secrets (.env) loading."""

from __future__ import annotations

import json
import logging
import os
from dataclasses import dataclass
from datetime import time
from pathlib import Path
from typing import Dict, List, Tuple

logger = logging.getLogger("parking.config")

ROOT = Path(__file__).resolve().parent
ENV_PATH = ROOT / ".env"
CONFIG_PATH = ROOT / "config.json"
STATE_PATH = ROOT / "state" / "state.json"
LOG_DIR = ROOT / "logs"
PROFILE_DIR = Path(os.environ.get("PARKING_PROFILE_DIR", ROOT / ".browser-profile"))


@dataclass(frozen=True)
class Named:
    id: int
    name: str


@dataclass(frozen=True)
class Config:
    dry_run: bool
    weekdays: List[str]
    horizon_days: int
    type: str
    centers: List[Named]
    sizes: List[Named]
    treat_placeless_as_covered: bool
    poll_minutes: int
    active_start: time
    active_end: time
    timezone: str
    jitter_max_seconds: int
    request_pause_seconds: Tuple[float, float]
    token_warning_days: int
    repeated_error_threshold: int


def _parse_hhmm(value: str) -> time:
    hours, minutes = value.split(":")
    return time(int(hours), int(minutes))


def load_config(path: Path = CONFIG_PATH) -> Config:
    raw = json.loads(path.read_text(encoding="utf-8"))
    pause = raw.get("request_pause_seconds", [1.5, 3.0])
    hours = raw.get("active_hours", {})
    return Config(
        dry_run=bool(raw.get("dry_run", True)),
        weekdays=[str(w) for w in raw["weekdays"]],
        horizon_days=int(raw.get("horizon_days", 30)),
        type=str(raw.get("type", "standard")),
        centers=[Named(int(c["id"]), str(c["name"])) for c in raw["centers"]],
        sizes=[Named(int(s["id"]), str(s["name"])) for s in raw["sizes"]],
        treat_placeless_as_covered=bool(raw.get("treat_placeless_as_covered", True)),
        poll_minutes=int(raw.get("poll_minutes", 15)),
        active_start=_parse_hhmm(hours.get("start", "00:00")),
        active_end=_parse_hhmm(hours.get("end", "23:59")),
        timezone=str(raw.get("timezone", "Europe/Madrid")),
        jitter_max_seconds=int(raw.get("jitter_max_seconds", 0)),
        request_pause_seconds=(float(pause[0]), float(pause[1])),
        token_warning_days=int(raw.get("token_warning_days", 3)),
        repeated_error_threshold=int(raw.get("repeated_error_threshold", 3)),
    )


def read_env_file(path: Path = ENV_PATH) -> Dict[str, str]:
    """Parse KEY=VALUE lines (no interpolation; surrounding quotes are stripped)."""
    values: Dict[str, str] = {}
    if not path.exists():
        return values
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, value = line.partition("=")
        values[key.strip()] = value.strip().strip('"').strip("'")
    return values


def load_env(path: Path = ENV_PATH) -> Dict[str, str]:
    """`.env` values, overridden by real environment variables of the same name."""
    values = read_env_file(path)
    for key in list(values):
        if key in os.environ:
            values[key] = os.environ[key]
    return values


def set_env_value(key: str, value: str, path: Path = ENV_PATH) -> None:
    """Insert or replace one KEY=VALUE line in the env file; never logs the value."""
    lines = path.read_text(encoding="utf-8").splitlines() if path.exists() else []
    replaced = False
    for i, line in enumerate(lines):
        if line.split("=", 1)[0].strip() == key:
            lines[i] = f"{key}={value}"
            replaced = True
    if not replaced:
        lines.append(f"{key}={value}")
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    logger.info("ℹ️ %s written to %s", key, path.name)
