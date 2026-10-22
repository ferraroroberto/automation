"""Regression tests for quickdeck shell actions (a throwaway python child, no real buttons)."""

import subprocess
import sys
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parents[3] / "system/quickdeck"))

from quickdeck_actions import ActionContext, execute_shell_action  # noqa: E402

# Writes far more than any pipe buffer holds, then exits.
NOISY_CHILD = "import sys; sys.stdout.write('x' * 2000000); sys.stderr.write('y' * 2000000)"


def _context(errors: list) -> ActionContext:
    return ActionContext(
        set_status=lambda message: None,
        show_error=lambda message, title: errors.append(message),
        show_output=lambda body, title: None,
        reload=lambda: None,
    )


class ShellActionOutputTests(unittest.TestCase):
    def test_a_chatty_command_is_not_blocked_on_unread_pipes(self):
        spawned = []
        real_popen = subprocess.Popen

        def spy(*args, **kwargs):
            process = real_popen(*args, **kwargs)
            spawned.append(process)
            return process

        errors = []
        config = {"cmd": sys.executable, "args": ["-c", NOISY_CHILD]}
        with mock.patch.object(subprocess, "Popen", side_effect=spy):
            execute_shell_action(config, _context(errors))

        self.assertEqual(errors, [])
        self.assertEqual(len(spawned), 1)
        process = spawned[0]
        try:
            # With stdout/stderr piped and never read, the child blocks on a full
            # buffer forever and wait() times out.
            self.assertEqual(process.wait(timeout=10), 0)
        finally:
            if process.poll() is None:
                process.kill()
                process.wait()
            for stream in (process.stdout, process.stderr):
                if stream is not None:
                    stream.close()


if __name__ == "__main__":
    unittest.main()
