"""Main application window with camera grid and management toolbar."""

import json
import logging
from pathlib import Path
from typing import Any, Dict, List, Optional

from PyQt6.QtCore import Qt, QSize
from PyQt6.QtGui import QAction, QFont, QColor
from PyQt6.QtWidgets import (
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QGridLayout,
    QLabel,
    QPushButton,
    QToolBar,
    QStatusBar,
    QDialog,
    QFormLayout,
    QLineEdit,
    QSpinBox,
    QCheckBox,
    QDialogButtonBox,
    QMessageBox,
    QScrollArea,
    QSizePolicy,
    QComboBox,
    QFrame,
    QApplication,
)

from camera_widget import CameraWidget

logger = logging.getLogger(__name__)

DEFAULT_CONFIG = {
    "cameras": [],
    "grid_columns": 2,
    "snapshot_dir": "snapshots",
    "recording_dir": "recordings",
    "reconnect_interval_sec": 5,
    "max_reconnect_attempts": 0,
    "stream_timeout_sec": 10,
    "window_width": 1280,
    "window_height": 720,
}


class AddCameraDialog(QDialog):
    """Dialog for adding or editing a camera entry."""

    def __init__(self, camera: Optional[Dict[str, Any]] = None, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Add Camera" if camera is None else "Edit Camera")
        self.setMinimumWidth(420)
        self.setStyleSheet(
            """
            QDialog { background-color: #1e1e2e; color: #e0e0e0; }
            QLabel { color: #e0e0e0; }
            QLineEdit {
                background-color: #2a2a4a; color: #e0e0e0;
                border: 1px solid #444; border-radius: 4px; padding: 6px;
            }
            QLineEdit:focus { border-color: #00bcd4; }
            """
        )

        layout = QFormLayout(self)

        self._name_edit = QLineEdit()
        self._name_edit.setPlaceholderText("e.g. Living Room")
        layout.addRow("Name:", self._name_edit)

        self._url_edit = QLineEdit()
        self._url_edit.setPlaceholderText("rtsp://192.168.1.100:554/stream1")
        layout.addRow("RTSP URL:", self._url_edit)

        self._enabled_check = QCheckBox("Enabled")
        self._enabled_check.setChecked(True)
        self._enabled_check.setStyleSheet("color: #e0e0e0;")
        layout.addRow("", self._enabled_check)

        if camera:
            self._name_edit.setText(camera.get("name", ""))
            self._url_edit.setText(camera.get("rtsp_url", ""))
            self._enabled_check.setChecked(camera.get("enabled", True))

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.setStyleSheet(
            """
            QPushButton {
                background-color: #00bcd4; color: #fff; padding: 6px 16px;
                border-radius: 4px; border: none; font-weight: bold;
            }
            QPushButton:hover { background-color: #00acc1; }
            """
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addRow(buttons)

    def get_camera_data(self) -> Dict[str, Any]:
        return {
            "name": self._name_edit.text().strip(),
            "rtsp_url": self._url_edit.text().strip(),
            "enabled": self._enabled_check.isChecked(),
        }


class FullscreenWindow(QMainWindow):
    """Borderless fullscreen window showing a single camera stream."""

    def __init__(self, camera_widget: CameraWidget, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setWindowTitle(f"Fullscreen - {camera_widget.camera_name}")
        self._original_parent_layout = camera_widget.parentWidget().layout() if camera_widget.parentWidget() else None
        self._camera_widget = camera_widget
        self.setStyleSheet("background-color: #000;")

        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)

        hint = QLabel("Press ESC or double-click to exit fullscreen")
        hint.setStyleSheet("color: #666; font-size: 11px; padding: 4px;")
        hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(hint)
        layout.addWidget(camera_widget, 1)

        self.setCentralWidget(container)
        self.showFullScreen()

    def keyPressEvent(self, event) -> None:
        if event.key() == Qt.Key.Key_Escape:
            self.close()
        super().keyPressEvent(event)

    def closeEvent(self, event) -> None:
        self.centralWidget().layout().removeWidget(self._camera_widget)
        super().closeEvent(event)


class MainWindow(QMainWindow):
    """Main application window hosting the camera grid."""

    def __init__(self, config_path: Path) -> None:
        super().__init__()
        self._config_path = config_path
        self._config = self._load_config()
        self._camera_widgets: List[CameraWidget] = []
        self._fullscreen_window: Optional[FullscreenWindow] = None
        self._grid_columns = self._config.get("grid_columns", 2)

        self._setup_window()
        self._setup_toolbar()
        self._setup_grid()
        self._setup_statusbar()
        self._populate_cameras()

    def _load_config(self) -> Dict[str, Any]:
        if not self._config_path.exists():
            logger.warning(f"⚠️ Config not found at {self._config_path}, using defaults")
            return DEFAULT_CONFIG.copy()

        try:
            with open(self._config_path, "r", encoding="utf-8") as f:
                config = json.load(f)
            logger.info(f"✅ Configuration loaded from {self._config_path}")
            return {**DEFAULT_CONFIG, **config}
        except (json.JSONDecodeError, OSError) as e:
            logger.error(f"❌ Failed to load config: {e}")
            return DEFAULT_CONFIG.copy()

    def _save_config(self) -> None:
        try:
            with open(self._config_path, "w", encoding="utf-8") as f:
                json.dump(self._config, f, indent=4)
            logger.info(f"✅ Configuration saved to {self._config_path}")
        except OSError as e:
            logger.error(f"❌ Failed to save config: {e}")

    def _setup_window(self) -> None:
        self.setWindowTitle("Xiaomi Camera Monitor")
        w = self._config.get("window_width", 1280)
        h = self._config.get("window_height", 720)
        self.resize(w, h)
        self.setMinimumSize(800, 600)
        self.setStyleSheet(
            """
            QMainWindow { background-color: #121225; }
            QToolBar {
                background-color: #1a1a2e;
                border-bottom: 1px solid #333;
                spacing: 6px;
                padding: 4px;
            }
            QStatusBar {
                background-color: #1a1a2e;
                color: #888;
                border-top: 1px solid #333;
            }
            """
        )

    def _setup_toolbar(self) -> None:
        toolbar = QToolBar("Main Toolbar")
        toolbar.setMovable(False)
        toolbar.setIconSize(QSize(20, 20))
        self.addToolBar(toolbar)

        title = QLabel("  📹 Xiaomi Cam Monitor  ")
        title.setStyleSheet("color: #00bcd4; font-size: 16px; font-weight: bold;")
        toolbar.addWidget(title)
        toolbar.addSeparator()

        btn_style = """
            QPushButton {
                background-color: #2a2a4a; color: #e0e0e0;
                border: 1px solid #444; border-radius: 4px;
                padding: 6px 12px; font-size: 12px;
            }
            QPushButton:hover { background-color: #3a3a5a; border-color: #00bcd4; }
        """

        self._btn_add = QPushButton("➕ Add Camera")
        self._btn_add.setStyleSheet(btn_style)
        self._btn_add.setCursor(Qt.CursorShape.PointingHandCursor)
        self._btn_add.clicked.connect(self._on_add_camera)
        toolbar.addWidget(self._btn_add)

        self._btn_start_all = QPushButton("▶ Start All")
        self._btn_start_all.setStyleSheet(btn_style)
        self._btn_start_all.setCursor(Qt.CursorShape.PointingHandCursor)
        self._btn_start_all.clicked.connect(self._on_start_all)
        toolbar.addWidget(self._btn_start_all)

        self._btn_stop_all = QPushButton("⏹ Stop All")
        self._btn_stop_all.setStyleSheet(btn_style)
        self._btn_stop_all.setCursor(Qt.CursorShape.PointingHandCursor)
        self._btn_stop_all.clicked.connect(self._on_stop_all)
        toolbar.addWidget(self._btn_stop_all)

        toolbar.addSeparator()

        col_label = QLabel(" Grid: ")
        col_label.setStyleSheet("color: #aaa; font-size: 12px;")
        toolbar.addWidget(col_label)

        self._col_spin = QSpinBox()
        self._col_spin.setRange(1, 6)
        self._col_spin.setValue(self._grid_columns)
        self._col_spin.setStyleSheet(
            """
            QSpinBox {
                background-color: #2a2a4a; color: #e0e0e0;
                border: 1px solid #444; border-radius: 4px; padding: 4px;
            }
            """
        )
        self._col_spin.valueChanged.connect(self._on_columns_changed)
        toolbar.addWidget(self._col_spin)

        spacer = QWidget()
        spacer.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        toolbar.addWidget(spacer)

        self._btn_snapshot_all = QPushButton("📷 Snapshot All")
        self._btn_snapshot_all.setStyleSheet(btn_style)
        self._btn_snapshot_all.setCursor(Qt.CursorShape.PointingHandCursor)
        self._btn_snapshot_all.clicked.connect(self._on_snapshot_all)
        toolbar.addWidget(self._btn_snapshot_all)

    def _setup_grid(self) -> None:
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet(
            "QScrollArea { border: none; background-color: #121225; }"
        )

        self._grid_container = QWidget()
        self._grid_layout = QGridLayout(self._grid_container)
        self._grid_layout.setSpacing(8)
        self._grid_layout.setContentsMargins(8, 8, 8, 8)
        scroll.setWidget(self._grid_container)
        self.setCentralWidget(scroll)

    def _setup_statusbar(self) -> None:
        self._statusbar = QStatusBar()
        self.setStatusBar(self._statusbar)
        self._status_label = QLabel("Ready")
        self._statusbar.addPermanentWidget(self._status_label)

    def _populate_cameras(self) -> None:
        cameras = self._config.get("cameras", [])
        snapshot_dir = Path(self._config.get("snapshot_dir", "snapshots"))
        recording_dir = Path(self._config.get("recording_dir", "recordings"))
        reconnect = self._config.get("reconnect_interval_sec", 5)
        max_reconnect = self._config.get("max_reconnect_attempts", 0)
        timeout = self._config.get("stream_timeout_sec", 10)

        for cam in cameras:
            if not cam.get("enabled", True):
                continue
            widget = CameraWidget(
                name=cam["name"],
                rtsp_url=cam["rtsp_url"],
                snapshot_dir=snapshot_dir,
                recording_dir=recording_dir,
                reconnect_interval=reconnect,
                max_reconnect=max_reconnect,
                stream_timeout=timeout,
            )
            widget.fullscreen_requested.connect(self._on_fullscreen)
            self._camera_widgets.append(widget)

        self._relayout_grid()
        self._update_status()

    def _relayout_grid(self) -> None:
        while self._grid_layout.count():
            item = self._grid_layout.takeAt(0)
            if item.widget():
                item.widget().setParent(None)

        for i, widget in enumerate(self._camera_widgets):
            row = i // self._grid_columns
            col = i % self._grid_columns
            self._grid_layout.addWidget(widget, row, col)

        if not self._camera_widgets:
            empty = QLabel("No cameras configured.\nClick '➕ Add Camera' to get started.")
            empty.setAlignment(Qt.AlignmentFlag.AlignCenter)
            empty.setStyleSheet("color: #555; font-size: 16px; padding: 40px;")
            self._grid_layout.addWidget(empty, 0, 0)

    def _update_status(self) -> None:
        total = len(self._camera_widgets)
        online = sum(1 for w in self._camera_widgets if w.status == "online")
        self._status_label.setText(f"Cameras: {online}/{total} online")

    def _on_add_camera(self) -> None:
        dialog = AddCameraDialog(parent=self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            data = dialog.get_camera_data()
            if not data["name"] or not data["rtsp_url"]:
                QMessageBox.warning(self, "Invalid", "Name and RTSP URL are required.")
                return

            self._config.setdefault("cameras", []).append(data)
            self._save_config()

            if data.get("enabled", True):
                widget = CameraWidget(
                    name=data["name"],
                    rtsp_url=data["rtsp_url"],
                    snapshot_dir=Path(self._config.get("snapshot_dir", "snapshots")),
                    recording_dir=Path(self._config.get("recording_dir", "recordings")),
                    reconnect_interval=self._config.get("reconnect_interval_sec", 5),
                    max_reconnect=self._config.get("max_reconnect_attempts", 0),
                    stream_timeout=self._config.get("stream_timeout_sec", 10),
                )
                widget.fullscreen_requested.connect(self._on_fullscreen)
                self._camera_widgets.append(widget)
                self._relayout_grid()
                widget.start()

            self._update_status()
            logger.info(f"✅ Camera '{data['name']}' added")

    def _on_start_all(self) -> None:
        logger.info("▶ Starting all cameras")
        for widget in self._camera_widgets:
            widget.start()
        self._update_status()

    def _on_stop_all(self) -> None:
        logger.info("⏹ Stopping all cameras")
        for widget in self._camera_widgets:
            widget.stop()
        self._update_status()

    def _on_columns_changed(self, value: int) -> None:
        self._grid_columns = value
        self._config["grid_columns"] = value
        self._save_config()
        self._relayout_grid()

    def _on_snapshot_all(self) -> None:
        logger.info("📷 Taking snapshots from all cameras")
        for widget in self._camera_widgets:
            if widget.status == "online" and widget._worker:
                widget._worker.take_snapshot()

    def _on_fullscreen(self, widget: CameraWidget) -> None:
        if self._fullscreen_window:
            self._fullscreen_window.close()

        parent_idx = self._camera_widgets.index(widget)
        row = parent_idx // self._grid_columns
        col = parent_idx % self._grid_columns

        self._grid_layout.removeWidget(widget)
        self._fullscreen_window = FullscreenWindow(widget, self)
        self._fullscreen_window.destroyed.connect(
            lambda: self._on_fullscreen_closed(widget, row, col)
        )

    def _on_fullscreen_closed(self, widget: CameraWidget, row: int, col: int) -> None:
        self._grid_layout.addWidget(widget, row, col)
        self._fullscreen_window = None

    def closeEvent(self, event) -> None:
        logger.info("ℹ️ Shutting down camera monitor")
        if self._fullscreen_window:
            self._fullscreen_window.close()
        for widget in self._camera_widgets:
            widget.stop()
        self._config["window_width"] = self.width()
        self._config["window_height"] = self.height()
        self._save_config()
        super().closeEvent(event)
