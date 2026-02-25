"""
Xiaomi Camera Monitor — custom desktop app for viewing RTSP camera streams.

Usage:
    python monitor.py [--config CONFIG_PATH]
"""

import argparse
import logging
import sys
from pathlib import Path

from PyQt6.QtWidgets import QApplication

from main_window import MainWindow

LOG_FORMAT = "%(asctime)s [%(levelname)s] %(name)s — %(message)s"
DEFAULT_CONFIG = Path(__file__).parent / "config.json"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Xiaomi Camera Monitor — RTSP multi-camera viewer"
    )
    parser.add_argument(
        "--config",
        type=Path,
        default=DEFAULT_CONFIG,
        help="Path to JSON configuration file (default: config.json)",
    )
    parser.add_argument(
        "--debug",
        action="store_true",
        help="Enable debug logging",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    level = logging.DEBUG if args.debug else logging.INFO
    logging.basicConfig(level=level, format=LOG_FORMAT)
    logger = logging.getLogger(__name__)

    logger.info("🚀 Starting Xiaomi Camera Monitor")
    logger.info(f"📂 Config: {args.config}")

    app = QApplication(sys.argv)
    app.setApplicationName("Xiaomi Camera Monitor")
    app.setStyle("Fusion")

    app.setStyleSheet(
        """
        QToolTip {
            background-color: #2a2a4a;
            color: #e0e0e0;
            border: 1px solid #444;
            padding: 4px;
            border-radius: 4px;
        }
        """
    )

    window = MainWindow(config_path=args.config)
    window.show()

    logger.info("✅ Monitor window ready")
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
