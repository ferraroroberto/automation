"""Individual camera view widget with overlay controls."""

import logging
from pathlib import Path
from typing import Optional

import cv2
import numpy as np
from PyQt6.QtCore import Qt, pyqtSignal, QSize, QTimer
from PyQt6.QtGui import QImage, QPixmap, QPainter, QColor, QFont, QIcon
from PyQt6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QSizePolicy,
    QFrame,
)

from camera_worker import CameraWorker

logger = logging.getLogger(__name__)

STATUS_COLORS = {
    "online": "#4CAF50",
    "connecting": "#FF9800",
    "reconnecting": "#FF9800",
    "offline": "#F44336",
}


class CameraWidget(QFrame):
    """Widget displaying a single camera feed with status and controls."""

    fullscreen_requested = pyqtSignal(object)

    def __init__(
        self,
        name: str,
        rtsp_url: str,
        snapshot_dir: Path,
        recording_dir: Path,
        reconnect_interval: int = 5,
        max_reconnect: int = 0,
        stream_timeout: int = 10,
        parent: Optional[QWidget] = None,
    ) -> None:
        super().__init__(parent)
        self._name = name
        self._status = "offline"
        self._worker: Optional[CameraWorker] = None
        self._rtsp_url = rtsp_url
        self._snapshot_dir = snapshot_dir
        self._recording_dir = recording_dir
        self._reconnect_interval = reconnect_interval
        self._max_reconnect = max_reconnect
        self._stream_timeout = stream_timeout

        self._setup_ui()
        self._setup_worker()

    def _setup_ui(self) -> None:
        self.setFrameStyle(QFrame.Shape.Box | QFrame.Shadow.Sunken)
        self.setLineWidth(1)
        self.setStyleSheet(
            "CameraWidget { background-color: #1a1a2e; border: 1px solid #333; border-radius: 8px; }"
        )
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.setMinimumSize(320, 240)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(4, 4, 4, 4)
        layout.setSpacing(2)

        header = QHBoxLayout()
        header.setSpacing(6)

        self._status_dot = QLabel("●")
        self._status_dot.setStyleSheet("color: #F44336; font-size: 12px;")
        self._status_dot.setFixedWidth(16)
        header.addWidget(self._status_dot)

        self._name_label = QLabel(self._name)
        self._name_label.setStyleSheet(
            "color: #e0e0e0; font-weight: bold; font-size: 13px;"
        )
        header.addWidget(self._name_label)

        header.addStretch()

        self._status_label = QLabel("Offline")
        self._status_label.setStyleSheet("color: #888; font-size: 11px;")
        header.addWidget(self._status_label)

        layout.addLayout(header)

        self._video_label = QLabel()
        self._video_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._video_label.setStyleSheet("background-color: #0d0d1a; border-radius: 4px;")
        self._video_label.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding
        )
        self._video_label.setMinimumSize(300, 200)
        self._set_placeholder()
        layout.addWidget(self._video_label, 1)

        controls = QHBoxLayout()
        controls.setSpacing(4)

        self._btn_snapshot = self._make_button("📷", "Take snapshot")
        self._btn_snapshot.clicked.connect(self._on_snapshot)
        controls.addWidget(self._btn_snapshot)

        self._btn_record = self._make_button("⏺", "Start recording")
        self._btn_record.clicked.connect(self._on_toggle_record)
        controls.addWidget(self._btn_record)

        self._btn_fullscreen = self._make_button("⛶", "Fullscreen")
        self._btn_fullscreen.clicked.connect(self._on_fullscreen)
        controls.addWidget(self._btn_fullscreen)

        self._btn_reconnect = self._make_button("🔄", "Reconnect")
        self._btn_reconnect.clicked.connect(self.reconnect)
        controls.addWidget(self._btn_reconnect)

        controls.addStretch()
        layout.addLayout(controls)

    def _make_button(self, text: str, tooltip: str) -> QPushButton:
        btn = QPushButton(text)
        btn.setToolTip(tooltip)
        btn.setFixedSize(36, 28)
        btn.setCursor(Qt.CursorShape.PointingHandCursor)
        btn.setStyleSheet(
            """
            QPushButton {
                background-color: #2a2a4a;
                color: #e0e0e0;
                border: 1px solid #444;
                border-radius: 4px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #3a3a5a;
                border-color: #00bcd4;
            }
            QPushButton:pressed {
                background-color: #1a1a3a;
            }
            """
        )
        return btn

    def _set_placeholder(self) -> None:
        self._video_label.setText("No Signal")
        self._video_label.setStyleSheet(
            "background-color: #0d0d1a; color: #555; font-size: 16px; border-radius: 4px;"
        )

    def _setup_worker(self) -> None:
        self._worker = CameraWorker(
            name=self._name,
            rtsp_url=self._rtsp_url,
            reconnect_interval=self._reconnect_interval,
            max_reconnect=self._max_reconnect,
            stream_timeout=self._stream_timeout,
        )
        self._worker.set_dirs(self._snapshot_dir, self._recording_dir)
        self._worker.frame_ready.connect(self._on_frame)
        self._worker.status_changed.connect(self._on_status_changed)
        self._worker.error_occurred.connect(self._on_error)

    def start(self) -> None:
        if self._worker and not self._worker.isRunning():
            self._worker.start()

    def stop(self) -> None:
        if self._worker and self._worker.isRunning():
            self._worker.stop()

    def reconnect(self) -> None:
        logger.info(f"🔄 [{self._name}] Manual reconnect requested")
        self.stop()
        self._setup_worker()
        self.start()

    def _on_frame(self, frame: np.ndarray) -> None:
        rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        h, w, ch = rgb.shape
        bytes_per_line = ch * w
        qt_image = QImage(rgb.data, w, h, bytes_per_line, QImage.Format.Format_RGB888)

        label_size = self._video_label.size()
        pixmap = QPixmap.fromImage(qt_image).scaled(
            label_size,
            Qt.AspectRatioMode.KeepAspectRatio,
            Qt.TransformationMode.SmoothTransformation,
        )
        self._video_label.setPixmap(pixmap)

    def _on_status_changed(self, status: str) -> None:
        self._status = status
        color = STATUS_COLORS.get(status, "#888")
        self._status_dot.setStyleSheet(f"color: {color}; font-size: 12px;")
        self._status_label.setText(status.capitalize())
        self._status_label.setStyleSheet(f"color: {color}; font-size: 11px;")

        if status == "offline":
            self._set_placeholder()

    def _on_error(self, message: str) -> None:
        logger.error(f"❌ [{self._name}] {message}")

    def _on_snapshot(self) -> None:
        if self._worker:
            path = self._worker.take_snapshot()
            if path:
                self._flash_button(self._btn_snapshot, "#4CAF50")

    def _on_toggle_record(self) -> None:
        if not self._worker:
            return

        if self._worker.is_recording:
            self._worker.stop_recording()
            self._btn_record.setText("⏺")
            self._btn_record.setToolTip("Start recording")
            self._btn_record.setStyleSheet(self._btn_record.styleSheet())
        else:
            path = self._worker.start_recording()
            if path:
                self._btn_record.setText("⏹")
                self._btn_record.setToolTip("Stop recording")
                self._btn_record.setStyleSheet(
                    """
                    QPushButton {
                        background-color: #b71c1c;
                        color: #fff;
                        border: 1px solid #d32f2f;
                        border-radius: 4px;
                        font-size: 14px;
                    }
                    QPushButton:hover { background-color: #c62828; }
                    QPushButton:pressed { background-color: #8e0000; }
                    """
                )

    def _on_fullscreen(self) -> None:
        self.fullscreen_requested.emit(self)

    def _flash_button(self, btn: QPushButton, color: str) -> None:
        original = btn.styleSheet()
        btn.setStyleSheet(
            f"""
            QPushButton {{
                background-color: {color};
                color: #fff;
                border: 1px solid {color};
                border-radius: 4px;
                font-size: 14px;
            }}
            """
        )
        QTimer.singleShot(300, lambda: btn.setStyleSheet(original))

    @property
    def camera_name(self) -> str:
        return self._name

    @property
    def status(self) -> str:
        return self._status

    def mouseDoubleClickEvent(self, event) -> None:
        self.fullscreen_requested.emit(self)
        super().mouseDoubleClickEvent(event)
