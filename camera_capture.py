"""Thread-safe, non-blocking camera reader used by every AI view."""
import threading
import time

try:
    import cv2
except ImportError:  # camera is optional on development machines
    cv2 = None


class CameraCapture:
    BACKENDS = {"auto": None, "dshow": "CAP_DSHOW", "msmf": "CAP_MSMF", "v4l2": "CAP_V4L2"}

    def __init__(self, index=0, capture_factory=None, retry_delay=0.1, backend="auto"):
        self.index = int(index)
        self.backend = str(backend or "auto").lower()
        self.capture_factory = capture_factory
        self.retry_delay = retry_delay
        self._lock = threading.Lock()
        self._stop = threading.Event()
        self._thread = None
        self._capture = None
        self._frame = None
        self.error = ""
        self.state = "stopped"

    def _open(self, index):
        if self.capture_factory is not None:
            return self.capture_factory(index)
        if cv2 is None:
            raise RuntimeError("無法載入 OpenCV；請安裝 opencv-python 並確認 DLL 可正常載入")
        constant = self.BACKENDS.get(self.backend)
        api = getattr(cv2, constant, None) if constant else None
        return cv2.VideoCapture(index) if api is None else cv2.VideoCapture(index, api)

    @classmethod
    def probe(cls, maximum=10, capture_factory=None, backend="auto"):
        """Return indexes which can actually be opened, releasing every probe."""
        found = []
        reader = cls(capture_factory=capture_factory, backend=backend)
        if cv2 is None and capture_factory is None:
            return found
        for index in range(maximum):
            cap = None
            try:
                cap = reader._open(index)
                if cap is not None and cap.isOpened():
                    found.append(index)
            except Exception:
                pass
            finally:
                if cap is not None:
                    cap.release()
        return found

    @property
    def running(self):
        return bool(self._thread and self._thread.is_alive())

    def start(self):
        if self.running:
            return
        if cv2 is None and self.capture_factory is None:
            self.state = "dependency_error"
            self.error = "無法載入 OpenCV；請安裝 opencv-python 並確認 DLL 可正常載入"
            return
        self.state = "starting"
        self._stop.clear()
        self._thread = threading.Thread(target=self._read_loop, daemon=True, name="ai-camera")
        self._thread.start()

    def _read_loop(self):
        try:
            cap = self._open(self.index)
            with self._lock:
                self._capture = cap
            if cap is None or not cap.isOpened():
                self.state = "open_error"
                self.error = "無法開啟相機 {}（backend: {}）；請確認裝置索引及是否被 OBS 占用".format(self.index, self.backend)
                return
            self.state = "connected"
            while not self._stop.is_set():
                ok, frame = cap.read()
                if ok and frame is not None:
                    with self._lock:
                        self._frame = frame.copy()
                    self.error = ""
                else:
                    self.state = "read_error"
                    self.error = "相機讀取失敗"
                    self._stop.wait(self.retry_delay)
        except Exception as exc:
            self.state = "dependency_error" if cv2 is None and self.capture_factory is None else "open_error"
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

    def configure(self, index=None, backend=None):
        if backend is not None:
            self.backend = str(backend).lower()
        self.reconnect(index)

    def stop(self):
        self._stop.set()
        thread = self._thread
        if thread and thread is not threading.current_thread():
            thread.join(timeout=2)
        self._thread = None
        if self.state not in ("dependency_error", "open_error", "read_error"):
            self.state = "stopped"
