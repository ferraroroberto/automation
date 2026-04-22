"""Portable whisper-server management.

This subpackage is self-contained and meant to be dropped, unchanged, into
any project that needs to own the local whisper.cpp server (this repo and
`ferraroroberto/claude-local-calls`). The sibling `whisper_server.yaml`
must be identical across projects so the server always binds the same
host/port.
"""

from .manager import (
    OWNERSHIP_EXTERNAL,
    OWNERSHIP_NONE,
    OWNERSHIP_OURS,
    ServerConfig,
    ServerStatus,
    WhisperServerManager,
    load_config,
)

__all__ = [
    "OWNERSHIP_EXTERNAL",
    "OWNERSHIP_NONE",
    "OWNERSHIP_OURS",
    "ServerConfig",
    "ServerStatus",
    "WhisperServerManager",
    "load_config",
]
