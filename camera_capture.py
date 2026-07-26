"""Thread-safe, non-blocking camera reader used by every AI view."""
import threading
import time

try:
    import cv2
except ImportError:  # camera is optional on development machines
    cv2 = None


class CameraCapture:
    def __init__(self, index=0, capture_factory=None, retry_delay=0.1):
        self.index = int(index)
        self.capture_factory = capture_factory or (lambda i: cv2.VideoCapture(i))
        self.retry_delay = retry_delay
        self._lock = threading.Lock()
        self._stop = threading.Event()
        self._thread = None
        self._capture = None
        self._frame = None
        self.error = ""

    @property
    def running(self):
        return bool(self._thread and self._thread.is_alive())

    def start(self):
        if self.running:
            return
        if cv2 is None and self.capture_factory is None:
            self.error = "未安裝 OpenCV"
            return
        self._stop.clear()
        self._thread = threading.Thread(target=self._read_loop, daemon=True, name="ai-camera")
        self._thread.start()

    def _read_loop(self):
        try:
            cap = self.capture_factory(self.index)
            with self._lock:
                self._capture = cap
            if cap is None or not cap.isOpened():
                self.error = "相機未連接或開啟失敗"
                return
            while not self._stop.is_set():
                ok, frame = cap.read()
                if ok and frame is not None:
                    with self._lock:
                        self._frame = frame.copy()
                    self.error = ""
                else:
                    self.error = "相機讀取失敗"
                    self._stop.wait(self.retry_delay)
        except Exception as exc:
            self.error = "相機錯誤：{}".format(exc)
        finally:
            with self._lock:
                cap, self._capture = self._capture, None
            if cap is not None:
                cap.release()

    def latest_frame(self):
        with self._lock:
            return None if self._frame is None else self._frame.copy()

    def reconnect(self, index=None):
        self.stop()
        if index is not None:
            self.index = int(index)
        with self._lock:
            self._frame = None
        self.start()

    def stop(self):
        self._stop.set()
        thread = self._thread
        if thread and thread is not threading.current_thread():
            thread.join(timeout=2)
        self._thread = None

