#!/usr/bin/env python3
"""
QuickDeck action executors.

One function per button `type` — shell, url, copy, python — plus the dispatch
that maps a button config onto them. The GUI is injected as callbacks through
`ActionContext` rather than imported, so the executors run (and can be driven
from a test or a REPL) without PySimpleGUI (audit issue #92).
"""

import os
import subprocess
import sys
import webbrowser
from dataclasses import dataclass
from io import StringIO
from typing import Any, Callable, Dict

import pyperclip

NO_WINDOW = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0


@dataclass
class ActionContext:
    """
    The four things an action needs from its host application.

    set_status: report progress/errors on the status line.
    show_error: modal error, called as (message, title).
    show_output: scrollable output, called as (body, title).
    reload:     rebuild the window from the config file on disk.
    """

    set_status: Callable[[str], None]
    show_error: Callable[[str, str], None]
    show_output: Callable[[str, str], None]
    reload: Callable[[], None]


def expand_path(path: str) -> str:
    """Expand ~ and environment variables in path."""
    if path.startswith('~'):
        path = os.path.expanduser(path)
    return os.path.expandvars(path)


def execute_shell_action(button_config: Dict[str, Any], ctx: ActionContext) -> None:
    """Execute shell command action."""
    try:
        cmd = button_config.get("cmd", "")
        args = button_config.get("args", [])
        cwd = button_config.get("cwd", ".")

        # Expand paths and environment variables
        cmd = expand_path(cmd)
        args = [expand_path(arg) for arg in args]
        cwd = expand_path(cwd)

        full_cmd = [cmd] + args if args else cmd

        ctx.set_status(f"🔄 Running: {cmd}")

        # Fire and forget - waiting for completion would block the UI thread
        subprocess.Popen(
            full_cmd,
            cwd=cwd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            creationflags=NO_WINDOW,
        )

        ctx.set_status(f"✅ Started: {cmd}")

    except Exception as e:
        error_msg = f"❌ Shell command failed: {e}"
        ctx.set_status(error_msg)
        ctx.show_error(error_msg, "Shell Command Error")


def execute_url_action(button_config: Dict[str, Any], ctx: ActionContext) -> None:
    """Execute URL action."""
    try:
        url = button_config.get("url", "")
        if not url:
            raise ValueError("No URL specified")

        ctx.set_status(f"🌐 Opening: {url}")
        webbrowser.open(url)
        ctx.set_status(f"✅ Opened: {url}")

    except Exception as e:
        error_msg = f"❌ URL action failed: {e}"
        ctx.set_status(error_msg)
        ctx.show_error(error_msg, "URL Error")


def execute_copy_action(button_config: Dict[str, Any], ctx: ActionContext) -> None:
    """Execute copy to clipboard action."""
    try:
        text = button_config.get("text", "")
        if not text:
            raise ValueError("No text specified")

        pyperclip.copy(text)
        ctx.set_status(f"📋 Copied: {text[:30]}{'...' if len(text) > 30 else ''}")

    except Exception as e:
        error_msg = f"❌ Copy action failed: {e}"
        ctx.set_status(error_msg)
        ctx.show_error(error_msg, "Copy Error")


# Names exposed to a "python" button's snippet. NOT a sandbox - see
# execute_python_action's docstring.
_SNIPPET_BUILTINS = {
    'print': print,
    'len': len,
    'str': str,
    'int': int,
    'float': float,
    'list': list,
    'dict': dict,
    'tuple': tuple,
    'set': set,
    'range': range,
    'enumerate': enumerate,
    'zip': zip,
    'min': min,
    'max': max,
    'sum': sum,
    'abs': abs,
    'round': round,
    'sorted': sorted,
    'reversed': reversed,
}


def execute_python_action(button_config: Dict[str, Any], ctx: ActionContext) -> None:
    """
    Execute a Python code snippet from the button config.

    This runs arbitrary Python, not sandboxed code: only `__builtins__` is
    restricted to a whitelist, which is trivially escapable (e.g. via
    `().__class__.__mro__[-1].__subclasses__()`) and does not restrict
    imports, file I/O, or process spawning. Only put trusted code in a
    "python" button's config (audit issue #67 — this used to be
    documented as "safely", which it isn't).
    """
    try:
        code = button_config.get("code", "")
        if not code:
            raise ValueError("No Python code specified")

        ctx.set_status("🐍 Executing Python code...")

        # Capture stdout and stderr
        old_stdout = sys.stdout
        old_stderr = sys.stderr
        stdout_capture = StringIO()
        stderr_capture = StringIO()

        try:
            sys.stdout = stdout_capture
            sys.stderr = stderr_capture
            exec(code, {'__builtins__': dict(_SNIPPET_BUILTINS)})
        finally:
            sys.stdout = old_stdout
            sys.stderr = old_stderr

        stdout_output = stdout_capture.getvalue()
        stderr_output = stderr_capture.getvalue()

        if stdout_output or stderr_output:
            result_text = ""
            if stdout_output:
                result_text += f"Output:\n{stdout_output}\n"
            if stderr_output:
                result_text += f"Errors:\n{stderr_output}\n"
            ctx.show_output(result_text, "Python Execution Result")

        ctx.set_status("✅ Python code executed successfully")

    except Exception as e:
        error_msg = f"❌ Python execution failed: {e}"
        ctx.set_status(error_msg)
        ctx.show_error(error_msg, "Python Execution Error")


_EXECUTORS: Dict[str, Callable[[Dict[str, Any], ActionContext], None]] = {
    "shell": execute_shell_action,
    "url": execute_url_action,
    "copy": execute_copy_action,
    "python": execute_python_action,
}


def execute_action(button_config: Dict[str, Any], ctx: ActionContext) -> None:
    """Dispatch a button config to the executor for its `type`."""
    try:
        action_type = button_config.get("type", "shell")

        if action_type == "reload":
            ctx.reload()
            return

        executor = _EXECUTORS.get(action_type)
        if executor is None:
            ctx.set_status(f"❌ Unknown action type: {action_type}")
            return

        executor(button_config, ctx)

    except Exception as e:
        error_msg = f"❌ Button action failed: {e}"
        ctx.set_status(error_msg)
        ctx.show_error(error_msg, "Button Action Error")
