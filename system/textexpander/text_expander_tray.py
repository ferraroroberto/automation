"""
Text Expander System Tray Module

Provides system tray icon functionality for background operation.
"""

import logging
import threading
import time
from typing import Optional, Callable
from pathlib import Path

try:
    import pystray
    from PIL import Image, ImageDraw
    PYSTRAY_AVAILABLE = True
except ImportError:
    PYSTRAY_AVAILABLE = False
    pystray = None
    Image = None
    ImageDraw = None

from text_expander_core import TextExpanderCore

# Configure module logger
logger = logging.getLogger(__name__)

class TextExpanderTray:
    """System tray icon for text expander background operation."""

    def __init__(self, core: TextExpanderCore) -> None:
        """Initialize the system tray.

        Args:
            core: The text expander core instance.
        """
        if not PYSTRAY_AVAILABLE:
            raise ImportError("❌ pystray library is required for system tray functionality")

        self.core = core
        self.tray_icon: Optional[pystray.Icon] = None
        self.monitoring_active = False

        # Callbacks
        self.on_toggle_monitoring: Optional[Callable[[bool], None]] = None
        self.on_show_gui: Optional[Callable[[], None]] = None
        self.on_exit: Optional[Callable[[], None]] = None

        logger.info("🔔 System tray initialized")

    def create_icon(self) -> pystray.Icon:
        """Create the system tray icon.

        Returns:
            The pystray Icon instance.
        """
        try:
            # Create a simple "E" icon
            icon_size = 64
            icon = Image.new('RGBA', (icon_size, icon_size), (0, 0, 0, 255))  # Black background
            draw = ImageDraw.Draw(icon)

            # Draw large white "E" - make it fill most of the icon
            # Use multiple rectangles to create a clear "E" shape
            margin = 8
            bar_width = 8
            bar_height = icon_size - 2 * margin

            # Vertical bar of E
            draw.rectangle([margin, margin, margin + bar_width, margin + bar_height],
                         fill=(255, 255, 255, 255))

            # Top horizontal bar
            draw.rectangle([margin + bar_width, margin, icon_size - margin, margin + bar_width],
                         fill=(255, 255, 255, 255))

            # Middle horizontal bar
            mid_y = icon_size // 2
            draw.rectangle([margin + bar_width, mid_y - bar_width//2,
                          icon_size - margin//2, mid_y + bar_width//2],
                         fill=(255, 255, 255, 255))

            # Bottom horizontal bar
            draw.rectangle([margin + bar_width, icon_size - margin - bar_width,
                          icon_size - margin, icon_size - margin],
                         fill=(255, 255, 255, 255))

            # Create tray icon
            menu = self._create_menu()
            self.tray_icon = pystray.Icon(
                "Text Expander",
                icon,
                "Text Expander (E)",
                menu=menu
            )

            return self.tray_icon

        except Exception as e:
            error_msg = f"❌ Failed to create tray icon: {e}"
            logger.error(error_msg)
            raise RuntimeError(error_msg) from e

    def _create_menu(self) -> pystray.Menu:
        """Create the tray icon context menu.

        Returns:
            The pystray menu instance.
        """
        return pystray.Menu(
            pystray.MenuItem(
                "Toggle Monitoring",
                self._on_toggle_monitoring_click,
                checked=lambda item: self.monitoring_active
            ),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem(
                "Manage Abbreviations",
                self._on_manage_click
            ),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem(
                "Exit",
                self._on_exit_click
            )
        )

    def show(self) -> None:
        """Show the system tray icon."""
        try:
            if not self.tray_icon:
                self.create_icon()

            logger.info("🔔 Showing system tray icon")
            self.tray_icon.run()

        except Exception as e:
            error_msg = f"❌ Failed to show system tray: {e}"
            logger.error(error_msg)

    def hide(self) -> None:
        """Hide the system tray icon."""
        try:
            if self.tray_icon:
                self.tray_icon.stop()
                self.tray_icon = None
                logger.info("🔔 System tray icon hidden")

        except Exception as e:
            logger.error(f"❌ Error hiding system tray: {e}")

    def update_status(self, monitoring_active: bool) -> None:
        """Update the tray icon status.

        Args:
            monitoring_active: Whether keyboard monitoring is active.
        """
        try:
            self.monitoring_active = monitoring_active

            # Update tooltip to show current status
            status_text = "Active" if monitoring_active else "Inactive"
            if self.tray_icon:
                self.tray_icon.title = f"Text Expander - {status_text}"

            logger.debug(f"📊 Tray status updated: {status_text}")

        except Exception as e:
            logger.error(f"❌ Error updating tray status: {e}")

    def show_notification(self, title: str, message: str) -> None:
        """Show a Windows notification.

        Args:
            title: Notification title.
            message: Notification message.
        """
        try:
            if self.tray_icon:
                self.tray_icon.notify(title, message)
                logger.debug(f"🔔 Notification shown: {title}")

        except Exception as e:
            logger.error(f"❌ Failed to show notification: {e}")

    def _on_toggle_monitoring_click(self, icon: pystray.Icon, item: pystray.MenuItem) -> None:
        """Handle toggle monitoring menu click."""
        try:
            new_state = not self.monitoring_active
            logger.info(f"🔄 Toggle monitoring clicked: {new_state}")

            if self.on_toggle_monitoring:
                self.on_toggle_monitoring(new_state)

        except Exception as e:
            logger.error(f"❌ Error handling toggle monitoring click: {e}")

    def _on_manage_click(self, icon: pystray.Icon, item: pystray.MenuItem) -> None:
        """Handle manage abbreviations menu click."""
        try:
            logger.info("🖱️ Manage abbreviations clicked")

            if self.on_show_gui:
                self.on_show_gui()

        except Exception as e:
            logger.error(f"❌ Error handling manage click: {e}")

    def _on_exit_click(self, icon: pystray.Icon, item: pystray.MenuItem) -> None:
        """Handle exit menu click."""
        try:
            logger.info("🚪 Exit clicked")

            if self.on_exit:
                self.on_exit()

        except Exception as e:
            logger.error(f"❌ Error handling exit click: {e}")

    def set_callbacks(
        self,
        on_toggle_monitoring: Optional[Callable[[bool], None]] = None,
        on_show_gui: Optional[Callable[[], None]] = None,
        on_exit: Optional[Callable[[], None]] = None
    ) -> None:
        """Set callback functions for tray events.

        Args:
            on_toggle_monitoring: Callback for monitoring toggle.
            on_show_gui: Callback for showing GUI.
            on_exit: Callback for application exit.
        """
        self.on_toggle_monitoring = on_toggle_monitoring
        self.on_show_gui = on_show_gui
        self.on_exit = on_exit

        logger.debug("🔗 Tray callbacks updated")
