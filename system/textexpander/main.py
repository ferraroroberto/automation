"""
Text Expander Main Application

Orchestrates the text expander application components including GUI, keyboard monitoring,
and system tray functionality.
"""

import logging
import sys
import signal
import threading
from typing import Optional
from pathlib import Path

# Configure module logger
logger = logging.getLogger(__name__)

# Import our modules
from .text_expander_core import TextExpanderCore, ConfigError
from .text_expander_gui import TextExpanderGUI
from .text_expander_monitor import KeyboardMonitor
from .text_expander_tray import TextExpanderTray

class TextExpanderApp:
    """Main text expander application orchestrator."""

    def __init__(self, config_path: Optional[Path] = None) -> None:
        """Initialize the text expander application.

        Args:
            config_path: Optional path to configuration file.
        """
        self.config_path = config_path
        self.core: Optional[TextExpanderCore] = None
        self.gui: Optional[TextExpanderGUI] = None
        self.monitor: Optional[KeyboardMonitor] = None
        self.tray: Optional[TextExpanderTray] = None

        # Application state
        self.running = False
        self.monitoring_active = False

        # Shutdown event for clean exit
        self.shutdown_event = threading.Event()

        logger.info("🚀 Text Expander Application initializing")

    def initialize(self) -> bool:
        """Initialize all application components.

        Returns:
            True if initialization successful, False otherwise.
        """
        try:
            logger.info("⚙️ Initializing application components")

            # Initialize core
            self.core = TextExpanderCore(self.config_path)
            logger.info("✅ Core component initialized")

            # Initialize GUI
            self.gui = TextExpanderGUI(self.core)
            logger.info("✅ GUI component initialized")

            # Initialize keyboard monitor
            self.monitor = KeyboardMonitor(
                self.core,
                on_expansion_triggered=self._on_expansion_triggered
            )
            logger.info("✅ Keyboard monitor initialized")

            # Initialize system tray
            self.tray = TextExpanderTray(self.core)
            self.tray.set_callbacks(
                on_toggle_monitoring=self._on_toggle_monitoring,
                on_show_gui=self._on_show_gui,
                on_exit=self._on_exit
            )
            logger.info("✅ System tray initialized")

            logger.info("🎉 All components initialized successfully")
            return True

        except Exception as e:
            error_msg = f"❌ Failed to initialize application: {e}"
            logger.error(error_msg)
            self._cleanup()
            return False

    def start(self) -> None:
        """Start the text expander application."""
        try:
            if self.running:
                logger.warning("⚠️ Application is already running")
                return

            logger.info("▶️ Starting Text Expander Application")

            # Set up signal handlers for clean shutdown
            signal.signal(signal.SIGINT, self._signal_handler)
            signal.signal(signal.SIGTERM, self._signal_handler)

            self.running = True

            # Start system tray in background thread
            tray_thread = threading.Thread(target=self._run_tray, daemon=True)
            tray_thread.start()

            # Check if auto-start monitoring is enabled
            settings = self.core.get_settings() if self.core else None
            if settings and settings.auto_start:
                logger.info("🔄 Auto-start enabled, starting keyboard monitoring")
                self._start_monitoring()
            else:
                logger.info("⏸️ Auto-start disabled, monitoring inactive")

            logger.info("✅ Text Expander Application started successfully")
            logger.info("💡 Right-click the tray icon to access options")

            # Keep main thread alive
            while self.running and not self.shutdown_event.is_set():
                self.shutdown_event.wait(1.0)

        except KeyboardInterrupt:
            logger.info("⏹️ Application interrupted by user")
        except Exception as e:
            logger.error(f"❌ Error running application: {e}")
        finally:
            self._shutdown()

    def _run_tray(self) -> None:
        """Run the system tray in a separate thread."""
        try:
            if self.tray:
                logger.debug("🔔 Starting system tray")
                self.tray.show()
        except Exception as e:
            logger.error(f"❌ System tray error: {e}")

    def _start_monitoring(self) -> None:
        """Start keyboard monitoring."""
        try:
            if not self.monitor:
                logger.error("❌ Keyboard monitor not initialized")
                return

            if self.monitor.start_monitoring():
                self.monitoring_active = True
                self._update_tray_status()

                if self.tray:
                    self.tray.show_notification(
                        "Text Expander Active",
                        "Keyboard monitoring is now active. Type your abbreviations!"
                    )

                logger.info("✅ Keyboard monitoring started")
            else:
                logger.error("❌ Failed to start keyboard monitoring")

        except Exception as e:
            logger.error(f"❌ Error starting monitoring: {e}")

    def _stop_monitoring(self) -> None:
        """Stop keyboard monitoring."""
        try:
            if not self.monitor:
                logger.warning("⚠️ Keyboard monitor not initialized")
                return

            self.monitor.stop_monitoring()
            self.monitoring_active = False
            self._update_tray_status()

            logger.info("✅ Keyboard monitoring stopped")

        except Exception as e:
            logger.error(f"❌ Error stopping monitoring: {e}")

    def _on_toggle_monitoring(self, enable: bool) -> None:
        """Handle monitoring toggle from system tray.

        Args:
            enable: Whether to enable or disable monitoring.
        """
        try:
            logger.info(f"🔄 Toggle monitoring: {enable}")

            if enable:
                self._start_monitoring()
            else:
                self._stop_monitoring()

        except Exception as e:
            logger.error(f"❌ Error toggling monitoring: {e}")

    def _on_show_gui(self) -> None:
        """Handle GUI show request from system tray."""
        try:
            logger.info("🖥️ Showing GUI from tray")

            if self.gui:
                # Show GUI in separate thread to avoid blocking
                gui_thread = threading.Thread(target=self.gui.show, daemon=True)
                gui_thread.start()
            else:
                logger.error("❌ GUI not initialized")

        except Exception as e:
            logger.error(f"❌ Error showing GUI: {e}")

    def _on_exit(self) -> None:
        """Handle exit request from system tray."""
        try:
            logger.info("🚪 Exit requested from tray")
            self._shutdown()
        except Exception as e:
            logger.error(f"❌ Error handling exit: {e}")

    def _on_expansion_triggered(self, expansion: str) -> None:
        """Handle successful text expansion.

        Args:
            expansion: The text that was expanded.
        """
        try:
            logger.debug(f"🔥 Expansion triggered: {len(expansion)} characters")

            # Show notification if enabled
            settings = self.core.get_settings() if self.core else None
            if settings and settings.show_notifications and self.tray:
                # Truncate long expansions for notification
                preview = expansion[:50] + "..." if len(expansion) > 50 else expansion
                self.tray.show_notification(
                    "Text Expanded",
                    f"Expanded to: {preview}"
                )

        except Exception as e:
            logger.error(f"❌ Error handling expansion: {e}")

    def _update_tray_status(self) -> None:
        """Update system tray status indicator."""
        try:
            if self.tray:
                self.tray.update_status(self.monitoring_active)
        except Exception as e:
            logger.error(f"❌ Error updating tray status: {e}")

    def _signal_handler(self, signum: int, frame) -> None:
        """Handle system signals for clean shutdown.

        Args:
            signum: Signal number.
            frame: Current stack frame.
        """
        logger.info(f"📡 Received signal {signum}, initiating shutdown")
        self._shutdown()

    def _shutdown(self) -> None:
        """Shutdown the application cleanly."""
        try:
            if not self.running:
                return

            logger.info("🔄 Shutting down Text Expander Application")
            self.running = False

            # Stop monitoring
            self._stop_monitoring()

            # Hide tray icon
            if self.tray:
                self.tray.hide()

            # Close GUI if open
            if self.gui:
                self.gui.close()

            # Set shutdown event
            self.shutdown_event.set()

            logger.info("✅ Application shutdown complete")

        except Exception as e:
            logger.error(f"❌ Error during shutdown: {e}")
        finally:
            self._cleanup()

    def _cleanup(self) -> None:
        """Clean up resources."""
        try:
            self.core = None
            self.gui = None
            self.monitor = None
            self.tray = None
            logger.debug("🧹 Resources cleaned up")

        except Exception as e:
            logger.error(f"❌ Error during cleanup: {e}")

def main() -> None:
    """Main entry point for the text expander application."""
    try:
        # Configure logging
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            handlers=[
                logging.StreamHandler(sys.stdout)
            ]
        )

        logger.info("🚀 Text Expander Application starting")

        # Create and run application
        app = TextExpanderApp()

        if app.initialize():
            app.start()
        else:
            logger.error("❌ Application initialization failed")
            sys.exit(1)

    except KeyboardInterrupt:
        logger.info("⏹️ Application interrupted by user")
    except Exception as e:
        logger.error(f"❌ Fatal error: {e}")
        sys.exit(1)
    finally:
        logger.info("👋 Text Expander Application finished")

if __name__ == "__main__":
    main()
