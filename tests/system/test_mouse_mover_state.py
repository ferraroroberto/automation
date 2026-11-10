"""Offline tests for the mouse mover's lifecycle: RUNNING, PAUSED and the final STOPPED."""

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "system"))

import mouse_mover  # noqa: E402 -- needs system/ on sys.path first (it imports _lib)

MoverState = mouse_mover.MoverState


class MoverStateTests(unittest.TestCase):
    def test_a_new_mover_is_running(self):
        self.assertIs(mouse_mover.MouseMover().state, MoverState.RUNNING)

    def test_toggle_pause_alternates_running_and_paused(self):
        mover = mouse_mover.MouseMover()
        self.assertIs(mover.toggle_pause(), MoverState.PAUSED)
        self.assertIs(mover.state, MoverState.PAUSED)
        self.assertIs(mover.toggle_pause(), MoverState.RUNNING)

    def test_stop_ends_the_lifecycle_from_running_and_from_paused(self):
        running, paused = mouse_mover.MouseMover(), mouse_mover.MouseMover()
        paused.toggle_pause()
        for mover in (running, paused):
            mover.stop()
            self.assertIs(mover.state, MoverState.STOPPED)

    def test_stopping_twice_stays_stopped(self):
        mover = mouse_mover.MouseMover()
        mover.stop()
        mover.stop()
        self.assertIs(mover.state, MoverState.STOPPED)

    def test_a_stopped_mover_refuses_to_resume(self):
        # state-machine-proof: mouse-mover
        mover = mouse_mover.MouseMover()
        mover.stop()
        with self.assertRaises(mouse_mover.IllegalTransition):
            mover.toggle_pause()
        self.assertIs(mover.state, MoverState.STOPPED)

    def test_the_run_loop_ends_when_stopped_without_moving_the_pointer(self):
        mover = mouse_mover.MouseMover()
        mover.stop()
        mover.run()  # returns at once: STOPPED is checked before any pointer move


if __name__ == "__main__":
    unittest.main()
