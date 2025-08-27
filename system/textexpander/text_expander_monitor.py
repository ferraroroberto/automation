"""
Text Expander Keyboard Monitor Module

Handles global keyboard monitoring and abbreviation detection using pynput.
"""

import logging
import threading
import time
from typing import Optional, Callable, List
from collections import deque
import pyperclip

try:
    from pynput import keyboard
    PYNPUT_AVAILABLE = True
except ImportError:
    PYNPUT_AVAILABLE = False
    keyboard = None

from text_expander_core import TextExpanderCore

# Configure module logger
logger = logging.getLogger(__name__)

class KeyboardMonitor:
    """Monitors global keyboard input and detects text expander abbreviations."""

    def __init__(self, core: TextExpanderCore, on_expansion_triggered: Optional[Callable[[str], None]] = None) -> None:
        """Initialize the keyboard monitor.

        Args:
            core: The text expander core instance.
            on_expansion_triggered: Callback function called when expansion is triggered.
        """
        if not PYNPUT_AVAILABLE:
            raise ImportError("❌ pynput library is required for keyboard monitoring")

        self.core = core
        self.on_expansion_triggered = on_expansion_triggered

        # Keyboard monitoring state
        self.monitoring = False
        self.listener: Optional[keyboard.Listener] = None

        # Character buffer for detecting abbreviations
        self.buffer: deque = deque(maxlen=100)  # Store last 100 characters
        self.buffer_lock = threading.Lock()

        # Monitor thread
        self.monitor_thread: Optional[threading.Thread] = None

        logger.info("⌨️ Keyboard monitor initialized")

    def start_monitoring(self) -> bool:
        """Start global keyboard monitoring.

        Returns:
            True if monitoring started successfully, False otherwise.
        """
        try:
            if self.monitoring:
                logger.warning("⚠️ Keyboard monitoring is already active")
                return True

            logger.info("🎯 Starting keyboard monitoring")

            # Start listener in separate thread
            self.monitor_thread = threading.Thread(target=self._run_listener, daemon=True)
            self.monitor_thread.start()

            self.monitoring = True
            logger.info("✅ Keyboard monitoring started successfully")
            return True

        except Exception as e:
            error_msg = f"❌ Failed to start keyboard monitoring: {e}"
            logger.error(error_msg)
            return False

    def stop_monitoring(self) -> None:
        """Stop global keyboard monitoring."""
        try:
            if not self.monitoring:
                logger.debug("⌨️ Keyboard monitoring already stopped")
                return

            logger.info("🛑 Stopping keyboard monitoring")
            self.monitoring = False

            # Stop listener
            if self.listener:
                self.listener.stop()
                self.listener = None

            # Wait for monitor thread to finish
            if self.monitor_thread and self.monitor_thread.is_alive():
                self.monitor_thread.join(timeout=1.0)

            logger.info("✅ Keyboard monitoring stopped successfully")

        except Exception as e:
            logger.error(f"❌ Error stopping keyboard monitoring: {e}")

    def _run_listener(self) -> None:
        """Run the keyboard listener in a separate thread."""
        try:
            with keyboard.Listener(
                on_press=self._on_key_press,
                on_release=self._on_key_release,
                suppress=False  # Don't suppress keys, just monitor them
            ) as listener:
                self.listener = listener
                listener.join()

        except Exception as e:
            logger.error(f"❌ Keyboard listener error: {e}")
            self.monitoring = False

    def _on_key_press(self, key: keyboard.Key) -> None:
        """Handle key press events."""
        try:
            # Convert key to character
            char = self._key_to_char(key)
            if char:
                with self.buffer_lock:
                    self.buffer.append(char)

                # Check for abbreviation after adding character
                self._check_for_abbreviation()

        except Exception as e:
            logger.error(f"❌ Error handling key press: {e}")

    def _on_key_release(self, key: keyboard.Key) -> None:
        """Handle key release events."""
        # Could be used for additional logic if needed
        pass

    def _key_to_char(self, key: keyboard.Key) -> Optional[str]:
        """Convert a keyboard key to its character representation.

        Args:
            key: The keyboard key.

        Returns:
            Character representation of the key, or None if not a character key.
        """
        try:
            if hasattr(key, 'char') and key.char:
                return key.char
            elif key == keyboard.Key.space:
                return ' '
            elif key == keyboard.Key.enter:
                return '\n'
            elif key == keyboard.Key.tab:
                return '\t'
            else:
                return None
        except AttributeError:
            return None

    def _check_for_abbreviation(self) -> None:
        """Check the current buffer for abbreviations that need expansion."""
        try:
            with self.buffer_lock:
                if not self.buffer:
                    return

                # Convert buffer to string
                current_text = ''.join(self.buffer)

            # Get settings
            settings = self.core.get_settings()
            if not settings:
                return

            trigger_key = settings.trigger_key

            # Look for trigger key followed by abbreviation
            if trigger_key not in current_text:
                return

            # Find the last occurrence of trigger key
            last_trigger_pos = current_text.rfind(trigger_key)
            if last_trigger_pos == -1:
                return

            # Get text after the trigger key
            text_after_trigger = current_text[last_trigger_pos + len(trigger_key):]

            # Check if we have a complete abbreviation (followed by space, newline, or end of buffer)
            if not text_after_trigger:
                return

            # Look for word boundary (space, newline, tab, or common punctuation)
            word_boundary_chars = ' \n\t.,!?;:'
            abbreviation_end = -1

            for i, char in enumerate(text_after_trigger):
                if char in word_boundary_chars:
                    abbreviation_end = i
                    break

            # If no boundary found, the abbreviation might still be being typed
            if abbreviation_end == -1:
                return

            # Extract the abbreviation
            abbreviation = text_after_trigger[:abbreviation_end]
            if not abbreviation:
                return

            # Check if this abbreviation exists
            full_abbreviation = f"{trigger_key}{abbreviation}"
            expansion = self.core.get_expansion(abbreviation)

            if expansion:
                logger.info(f"🔥 Abbreviation detected: {full_abbreviation}")

                # Calculate how many characters to remove (trigger key + abbreviation + boundary char)
                chars_to_remove = len(trigger_key) + len(abbreviation)
                if abbreviation_end < len(text_after_trigger):
                    chars_to_remove += 1  # Include the boundary character

                # Trigger expansion in a separate thread to avoid blocking
                expansion_thread = threading.Thread(
                    target=self._perform_expansion,
                    args=(expansion, chars_to_remove),
                    daemon=True
                )
                expansion_thread.start()

                # Clear the buffer to prevent re-triggering
                with self.buffer_lock:
                    self.buffer.clear()

        except Exception as e:
            logger.error(f"❌ Error checking for abbreviation: {e}")

    def _perform_expansion(self, expansion: str, chars_to_remove: int) -> None:
        """Perform the text expansion by simulating backspace and paste, supporting {ENTER} placeholder for simulated Enter keypresses."""
        controller = None
        try:
            logger.debug(f"🔄 Performing expansion: removing {chars_to_remove} chars, pasting {len(expansion)} chars")
            time.sleep(0.05)
            controller = keyboard.Controller()
            # Remove abbreviation
            for _ in range(chars_to_remove):
                controller.press(keyboard.Key.backspace)
                controller.release(keyboard.Key.backspace)
                time.sleep(0.01)
            # Save clipboard
            original_clipboard = pyperclip.paste()
            # Split expansion by {ENTER} placeholder
            segments = expansion.split('{ENTER}')
            for i, segment in enumerate(segments):
                # Paste segment
                pyperclip.copy(segment)
                controller.press(keyboard.Key.ctrl)
                controller.press('v')
                controller.release('v')
                controller.release(keyboard.Key.ctrl)
                time.sleep(0.05)
                # Simulate Enter after each segment except last
                if i < len(segments) - 1:
                    controller.press(keyboard.Key.enter)
                    controller.release(keyboard.Key.enter)
                    time.sleep(0.05)
            # Restore clipboard
            time.sleep(0.1)
            pyperclip.copy(original_clipboard)
            logger.info("✅ Text expansion completed successfully")
            if self.on_expansion_triggered:
                try:
                    self.on_expansion_triggered(expansion)
                except Exception as e:
                    logger.error(f"❌ Error in expansion callback: {e}")
        except Exception as e:
            logger.error(f"❌ Failed to perform text expansion: {e}")
        finally:
            if controller:
                controller = None

    def is_monitoring(self) -> bool:
        """Check if keyboard monitoring is active.

        Returns:
            True if monitoring is active, False otherwise.
        """
        return self.monitoring

    def get_buffer_content(self) -> str:
        """Get current buffer content for debugging.

        Returns:
            Current buffer content as string.
        """
        with self.buffer_lock:
            return ''.join(self.buffer)
