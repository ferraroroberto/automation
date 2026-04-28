"""
Configuration loader for the illustration-color-edit project.

Single source of truth: ``config.json`` in the project root.
Falls back to ``config.json.example`` if no real config exists yet (useful
on a fresh checkout — the app launches with sensible defaults instead of
crashing).
"""

from __future__ import annotations

import json
import logging
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Optional

log = logging.getLogger(__name__)


PROJECT_ROOT = Path(__file__).resolve().parent.parent


@dataclass
class MatchingConfig:
    nearest_enabled: bool = True
    metric: str = "lab"
    threshold: float = 10.0


@dataclass
class PrintSafetyConfig:
    min_gray_value: str = "#EEEEEE"
    warn_only: bool = True


@dataclass
class PathsConfig:
    input_dir: Path = field(default_factory=lambda: PROJECT_ROOT / "input")
    output_dir: Path = field(default_factory=lambda: PROJECT_ROOT / "output")
    metadata_dir: Path = field(default_factory=lambda: PROJECT_ROOT / "metadata")


@dataclass
class AppConfig:
    """Resolved application config. Use ``load_config()`` to construct."""

    global_color_map: dict[str, dict[str, str]] = field(default_factory=dict)
    matching: MatchingConfig = field(default_factory=MatchingConfig)
    print_safety: PrintSafetyConfig = field(default_factory=PrintSafetyConfig)
    paths: PathsConfig = field(default_factory=PathsConfig)
    log_level: str = "INFO"
    source_path: Optional[Path] = None

    def ensure_dirs(self) -> None:
        """Create the configured input/output/metadata directories if missing."""
        for p in (self.paths.input_dir, self.paths.output_dir, self.paths.metadata_dir):
            p.mkdir(parents=True, exist_ok=True)


def _resolve_path(raw: str, base: Path) -> Path:
    p = Path(raw)
    return p if p.is_absolute() else (base / p).resolve()


def load_config(path: Optional[Path] = None) -> AppConfig:
    """
    Load and validate config from ``path`` (default: project_root/config.json).

    Resolution order:
      1. explicit ``path`` argument
      2. ``PROJECT_ROOT/config.json``
      3. ``PROJECT_ROOT/config.json.example``  (fallback for fresh checkouts)
      4. built-in defaults                     (last resort)
    """
    candidates: list[Path] = []
    if path is not None:
        candidates.append(path)
    candidates.append(PROJECT_ROOT / "config.json")
    candidates.append(PROJECT_ROOT / "config.json.example")

    chosen: Optional[Path] = None
    for c in candidates:
        if c.is_file():
            chosen = c
            break

    if chosen is None:
        log.warning("No config.json found; using built-in defaults.")
        return AppConfig()

    log.info("Loading config from %s", chosen)
    raw: dict[str, Any] = json.loads(chosen.read_text(encoding="utf-8"))

    cfg = AppConfig(source_path=chosen)
    cfg.global_color_map = {
        k.upper(): v for k, v in raw.get("global_color_map", {}).items()
    }

    matching = raw.get("matching", {})
    cfg.matching = MatchingConfig(
        nearest_enabled=bool(matching.get("nearest_enabled", True)),
        metric=str(matching.get("metric", "lab")).lower(),
        threshold=float(matching.get("threshold", 10.0)),
    )

    safety = raw.get("print_safety", {})
    cfg.print_safety = PrintSafetyConfig(
        min_gray_value=str(safety.get("min_gray_value", "#EEEEEE")).upper(),
        warn_only=bool(safety.get("warn_only", True)),
    )

    paths = raw.get("paths", {})
    base = chosen.parent
    cfg.paths = PathsConfig(
        input_dir=_resolve_path(paths.get("input_dir", "./input"), base),
        output_dir=_resolve_path(paths.get("output_dir", "./output"), base),
        metadata_dir=_resolve_path(paths.get("metadata_dir", "./metadata"), base),
    )

    cfg.log_level = str(raw.get("logging", {}).get("level", "INFO")).upper()
    return cfg


def configure_logging(level: str = "INFO") -> None:
    """Configure root logger once. Idempotent."""
    logging.basicConfig(
        level=getattr(logging, level, logging.INFO),
        format="%(asctime)s %(levelname)-7s %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
