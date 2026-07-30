"""
Shared Windows single-instance guard.

``mouse_mover.py`` and ``foldersearcher/foldersearcher.py`` both import
``acquire_named_mutex`` from here instead of each defining their own
byte-for-byte identical mutex-acquisition helper.

Uses a named mutex rather than a fixed loopback TCP port: Windows reserves
large high-port ranges (Hyper-V/WSL), so ``bind()`` can fail with
WinError 10013 even when no instance is running.
"""

import ctypes
from ctypes import wintypes

ERROR_ALREADY_EXISTS = 183


def acquire_named_mutex(name: str):
    """Take a process-wide named mutex.

    Args:
        name: The mutex name (process-wide, so keep it app-specific).

    Returns:
        Tuple of ``(handle, acquired)``. ``acquired`` is False if another
        instance already holds the mutex. The mutex is released
        automatically when the process exits, so no explicit cleanup is
        needed — but the caller must keep a reference to ``handle`` for the
        process lifetime (assign it to an instance attribute) so it is not
        garbage-collected early.
    """
    kernel32 = ctypes.windll.kernel32
    kernel32.CreateMutexW.restype = wintypes.HANDLE
    kernel32.CreateMutexW.argtypes = [wintypes.LPVOID, wintypes.BOOL, wintypes.LPCWSTR]
    handle = kernel32.CreateMutexW(None, True, name)
    acquired = kernel32.GetLastError() != ERROR_ALREADY_EXISTS
    return handle, acquired
