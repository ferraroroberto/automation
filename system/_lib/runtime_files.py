"""Writers for paths an operator names, plus the machine-local runtime folder.

Most tools here write where the caller points them: a CLI argument, a config
value, a folder picked in a tool's own UI. The runtime-write ratchet cannot
resolve such a path, so every call would need its own exempt marker. These
helpers take that path, refuse it at runtime when it resolves inside this
checkout, and then write. Output belongs in a folder the operator chose or in
``runtime_data_dir()``, never beside the code.

Scripts run as ``python <folder>/<script>.py``, so ``sys.path[0]`` is the
script's folder. Put the repo root on the path before importing::

    sys.path.append(str(Path(__file__).resolve().parents[1]))
    from system._lib import runtime_files  # noqa: E402
"""

from __future__ import annotations

import logging
import os
import shutil
from pathlib import Path
from typing import IO, Any, Union

PathLike = Union[str, "os.PathLike[str]"]

REPO_ROOT = Path(__file__).resolve().parents[2]

#: Overrides ``runtime_data_dir()`` (tests and relocated installs); must be outside the checkout.
RUNTIME_DIR_ENV = "AUTOMATION_RUNTIME_DIR"


def outside_repo(path: PathLike) -> Path:
    """Return ``path`` resolved, or raise ``ValueError`` when it sits inside this checkout."""
    resolved = Path(path).expanduser().resolve()
    if resolved == REPO_ROOT or REPO_ROOT in resolved.parents:
        raise ValueError(
            f"Refusing to write inside the repository: {resolved}. "
            "Point the destination at a folder outside the checkout."
        )
    return resolved


def runtime_data_dir() -> Path:
    """Machine-local folder for logs and state, outside the checkout (not created here)."""
    override = os.environ.get(RUNTIME_DIR_ENV)
    if override:
        return outside_repo(override)
    base = os.environ.get("LOCALAPPDATA") or str(Path.home() / ".local" / "share")
    return outside_repo(Path(base) / "automation")


def make_dirs(path: PathLike, *, exist_ok: bool = True) -> Path:
    """Create ``path`` and its parents; refuses a path inside the checkout."""
    target = outside_repo(path)
    # runtime-write-exempt: caller-named folder; outside_repo() above raises when it is inside the checkout
    os.makedirs(target, exist_ok=exist_ok)
    return target


def open_for_write(path: PathLike, mode: str = "w", **kwargs: Any) -> IO[Any]:
    """Open ``path`` for writing, creating no folders; refuses a path inside the checkout."""
    # runtime-write-exempt: caller-named file; outside_repo() raises when it is inside the checkout
    return open(outside_repo(path), mode, **kwargs)


def save_text(path: PathLike, text: str, encoding: str = "utf-8", **kwargs: Any) -> None:
    """Write ``text`` to ``path`` in one call."""
    with open_for_write(path, "w", encoding=encoding, **kwargs) as handle:
        handle.write(text)


def save_bytes(path: PathLike, data: bytes) -> None:
    """Write ``data`` to ``path`` in one call."""
    with open_for_write(path, "wb") as handle:
        handle.write(data)


def copy_file(src: PathLike, dst: PathLike) -> Path:
    """Copy ``src`` to ``dst`` keeping metadata; refuses a destination inside the checkout."""
    target = outside_repo(dst)
    # runtime-write-exempt: caller-named destination; outside_repo() above raises when it is inside the checkout
    shutil.copy2(src, target)
    return target


def file_handler(path: PathLike, mode: str = "a", encoding: str = "utf-8") -> logging.FileHandler:
    """A ``logging.FileHandler`` on ``path``, with its folder created; refuses a path inside the checkout."""
    target = outside_repo(path)
    make_dirs(target.parent)
    # runtime-write-exempt: caller-named log file; outside_repo() above raises when it is inside the checkout
    return logging.FileHandler(target, mode=mode, encoding=encoding)
