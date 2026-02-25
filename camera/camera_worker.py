"""Threaded camera stream worker for RTSP capture and frame delivery."""

import logging
import os
import time
from datetime import datetime
from pathlib import Path
from typing import Optional

import cv2
import numpy as np
from PyQt6.QtCore import QThread, pyqtSignal, QMutex, QMutexLocker

logger = logging.getLogger(__name__)

RECONNECT_BASE_SEC = 5
STREAM_TIMEOUT_SEC = 10
FRAME_READ_TIMEOUT_MS = 5000


class CameraWorker(QThread):
    """Background thread that captures frames from an RTSP source."""

    frame_ready = pyqtSignal(np.ndarray)
    status_changed = pyqtSignal(str)
    error_occurred = pyqtSignal(str)

    def __init__(
        self,
        name: str,
        rtsp_url: str,
        reconnect_interval: int = RECONNECT_BASE_SEC,
        max_reconnect: int = 0,
        stream_timeout: int = STREAM_TIMEOUT_SEC,
    ) -> None:
        super().__init__()
        self._name = name
        self._rtsp_url = rtsp_url
        self._reconnect_interval = reconnect_interval
        self._max_reconnect = max_reconnect
        self._stream_timeout = stream_timeout

        self._running = False
        self._recording = False
        self._video_writer: Optional[cv2.VideoWriter] = None
        self._mutex = QMutex()
        self._last_frame: Optional[np.ndarray] = None
        self._recording_dir = Path("recordings")
        self._snapshot_dir = Path("snapshots")

    @property
    def name(self) -> str:
        return self._name

    @property
    def rtsp_url(self) -> str:
        return self._rtsp_url

    @rtsp_url.setter
    def rtsp_url(self, url: str) -> None:
        self._rtsp_url = url

    def set_dirs(self, snapshot_dir: Path, recording_dir: Path) -> None:
        self._snapshot_dir = snapshot_dir
        self._recording_dir = recording_dir

    def run(self) -> None:
        self._running = True
        attempt = 0

        while self._running:
            cap = self._connect()
            if cap is None:
                attempt += 1
                if self._max_reconnect > 0 and attempt >= self._max_reconnect:
                    logger.error(f"❌ [{self._name}] Max reconnect attempts reached")
                    self.status_changed.emit("offline")
                    break
                wait = min(self._reconnect_interval * (2 ** min(attempt - 1, 4)), 60)
                logger.warning(
                    f"⚠️ [{self._name}] Reconnecting in {wait}s (attempt {attempt})"
                )
                self.status_changed.emit("reconnecting")
                self._sleep(wait)
                continue

            attempt = 0
            self.status_changed.emit("online")
            logger.info(f"✅ [{self._name}] Stream connected")

            while self._running:
                ret, frame = cap.read()
                if not ret:
                    logger.warning(f"⚠️ [{self._name}] Frame read failed, reconnecting")
                    self.status_changed.emit("reconnecting")
                    break

                with QMutexLocker(self._mutex):
                    self._last_frame = frame.copy()

                if self._recording and self._video_writer is not None:
                    self._video_writer.write(frame)

                self.frame_ready.emit(frame)

            cap.release()
            if not self._running:
                break
            self._sleep(self._reconnect_interval)

        self._stop_recording_internal()
        self.status_changed.emit("offline")
        logger.info(f"ℹ️ [{self._name}] Worker stopped")

    def _connect(self) -> Optional[cv2.VideoCapture]:
        logger.info(f"📡 [{self._name}] Connecting to {self._rtsp_url}")
        self.status_changed.emit("connecting")

        os.environ["OPENCV_FFMPEG_CAPTURE_OPTIONS"] = "rtsp_transport;tcp"
        cap = cv2.VideoCapture(self._rtsp_url, cv2.CAP_FFMPEG)
        cap.set(cv2.CAP_PROP_OPEN_TIMEOUT_MSEC, self._stream_timeout * 1000)
        cap.set(cv2.CAP_PROP_READ_TIMEOUT_MSEC, FRAME_READ_TIMEOUT_MS)
        cap.set(cv2.CAP_PROP_BUFFERSIZE, 1)

        if not cap.isOpened():
            logger.error(f"❌ [{self._name}] Failed to open stream")
            self.error_occurred.emit(f"Cannot connect to {self._name}")
            cap.release()
            return None

        return cap

    def _sleep(self, seconds: float) -> None:
        end = time.monotonic() + seconds
        while self._running and time.monotonic() < end:
            time.sleep(0.25)

    def stop(self) -> None:
        self._running = False
        self.wait(5000)

    def take_snapshot(self) -> Optional[Path]:
        with QMutexLocker(self._mutex):
            if self._last_frame is None:
                logger.warning(f"⚠️ [{self._name}] No frame available for snapshot")
                return None
            frame = self._last_frame.copy()

        self._snapshot_dir.mkdir(parents=True, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_name = self._name.replace(" ", "_").lower()
        path = self._snapshot_dir / f"{safe_name}_{ts}.jpg"
        cv2.imwrite(str(path), frame)
        logger.info(f"📸 [{self._name}] Snapshot saved: {path}")
        return path

    def start_recording(self) -> Optional[Path]:
        with QMutexLocker(self._mutex):
            if self._last_frame is None:
                logger.warning(f"⚠️ [{self._name}] No frame available to start recording")
                return None
            h, w = self._last_frame.shape[:2]

        self._recording_dir.mkdir(parents=True, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_name = self._name.replace(" ", "_").lower()
        path = self._recording_dir / f"{safe_name}_{ts}.mp4"

        fourcc = cv2.VideoWriter_fourcc(*"mp4v")
        self._video_writer = cv2.VideoWriter(str(path), fourcc, 20.0, (w, h))
        self._recording = True
        logger.info(f"🔴 [{self._name}] Recording started: {path}")
        return path

    def stop_recording(self) -> None:
        self._stop_recording_internal()

    def _stop_recording_internal(self) -> None:
        if self._recording:
            self._recording = False
            if self._video_writer is not None:
                self._video_writer.release()
                self._video_writer = None
            logger.info(f"⏹️ [{self._name}] Recording stopped")

    @property
    def is_recording(self) -> bool:
        return self._recording
