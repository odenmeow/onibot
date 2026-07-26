"""Camera discovery, stable device selection and threaded capture.

``CameraCapture`` is the camera entry point used by ``ai_config_dialog.py``;
``front.py`` never opens OpenCV devices directly.  Discovery returns friendly
names and stable device IDs while numeric indexes remain runtime details.
"""
import json
import os
import subprocess
import threading
import time

try:
    import cv2
except ImportError:  # camera support is optional on development machines
    cv2 = None


TYPE_LABELS = {"usb": "USB", "integrated": "內建", "virtual": "虛擬", "unknown": "相機"}
BACKEND_LABELS = {"dshow": "DirectShow", "msmf": "Media Foundation", "v4l2": "V4L2", "auto": "自動"}


def classify_device(name, device_id=""):
    text = (str(name) + " " + str(device_id)).lower()
    if any(word in text for word in ("obs", "virtual", "manycam", "snap camera")):
        return "virtual"
    if any(word in text for word in ("integrated", "built-in", "facetime", "內建")):
        return "integrated"
    if "usb" in text or "vid_" in text:
        return "usb"
    return "unknown"


class CameraCapture:
    BACKENDS = {"auto": None, "dshow": "CAP_DSHOW", "msmf": "CAP_MSMF", "v4l2": "CAP_V4L2"}
    COMMON_RESOLUTIONS = ((640, 480), (1280, 720), (1920, 1080))
    COMMON_FPS = (15, 30, 60)

    def __init__(self, index=0, capture_factory=None, retry_delay=.1, backend="auto", **settings):
        self.index, self.backend = int(index), str(backend or "auto").lower()
        self.device_id = str(settings.get("device_id") or "")
        self.device_name = str(settings.get("camera_name") or "")
        self.width = settings.get("width"); self.height = settings.get("height")
        self.fps = settings.get("fps"); self.fourcc = settings.get("fourcc")
        self.capture_factory, self.retry_delay = capture_factory, retry_delay
        self.actual = {}; self.error = ""; self.state = "stopped"
        self._lock = threading.Lock(); self._stop = threading.Event()
        self._thread = self._capture = self._frame = None

    def _open(self, index):
        if self.capture_factory is not None: return self.capture_factory(index)
        if cv2 is None: raise RuntimeError("無法載入 OpenCV；請安裝 opencv-python 並確認 DLL 可正常載入")
        constant = self.BACKENDS.get(self.backend)
        api = getattr(cv2, constant, None) if constant else None
        return cv2.VideoCapture(index) if api is None else cv2.VideoCapture(index, api)

    @staticmethod
    def _windows_names():
        """Read DirectShow/PnP friendly names without an extra dependency."""
        if os.name != "nt": return []
        script = ("Get-PnpDevice -PresentOnly | Where-Object {$_.Class -in 'Camera','Image'} | "
                  "Select-Object FriendlyName,InstanceId | ConvertTo-Json -Compress")
        try:
            raw = subprocess.check_output(["powershell", "-NoProfile", "-Command", script], timeout=8,
                                          creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
            data = json.loads(raw.decode("utf-8-sig") or "[]")
            return data if isinstance(data, list) else [data]
        except Exception:
            return []

    @classmethod
    def discover(cls, maximum=10, capture_factory=None, backend=None):
        backend = (backend or ("dshow" if os.name == "nt" else "v4l2")).lower()
        indexes = cls.probe(maximum, capture_factory, backend)
        pnp = cls._windows_names()
        devices = []
        for position, index in enumerate(indexes):
            info = pnp[position] if position < len(pnp) else {}
            name = str(info.get("FriendlyName") or ("Camera {}".format(index)))
            device_id = str(info.get("InstanceId") or "{}:{}".format(backend, name))
            kind = classify_device(name, device_id)
            devices.append({"name": name, "display_name": "[{}] {}".format(TYPE_LABELS[kind], name),
                            "device_id": device_id, "index": index, "device_type": kind, "backend": backend})
        return devices

    @classmethod
    def probe(cls, maximum=10, capture_factory=None, backend="auto"):
        found, reader = [], cls(capture_factory=capture_factory, backend=backend)
        if cv2 is None and capture_factory is None: return found
        for index in range(maximum):
            cap = None
            try:
                cap = reader._open(index)
                if cap is not None and cap.isOpened(): found.append(index)
            except Exception: pass
            finally:
                if cap is not None: cap.release()
        return found

    @classmethod
    def probe_capabilities(cls, device, capture_factory=None):
        """Verify common settings with a real frame; return unique actual modes."""
        results = []
        if cv2 is None and capture_factory is None: return results
        for width, height in cls.COMMON_RESOLUTIONS:
            for fps in cls.COMMON_FPS:
                reader = cls(device["index"], capture_factory, backend=device.get("backend", "auto"),
                             width=width, height=height, fps=fps)
                cap = None
                try:
                    cap = reader._open(reader.index)
                    if not cap or not cap.isOpened(): continue
                    reader._apply(cap, validate=False)
                    ok, frame = cap.read()
                    if not ok or frame is None: continue
                    actual = reader._actual(cap)
                    mode = (actual["width"], actual["height"], round(actual["fps"] or fps, 2))
                    if mode not in [(x["width"], x["height"], x["fps"]) for x in results]:
                        results.append(dict(width=mode[0], height=mode[1], fps=mode[2]))
                except Exception: pass
                finally:
                    if cap is not None: cap.release()
        return results

    @staticmethod
    def select_device(devices, device_id="", name="", index=None):
        for key, value in (("device_id", device_id), ("name", name)):
            if value:
                found = next((d for d in devices if d.get(key) == value), None)
                if found: return found
        found = next((d for d in devices if d.get("index") == index), None)
        # Never silently use a virtual camera as fallback.
        return found if found and (device_id or name or found.get("device_type") != "virtual") else None

    def _actual(self, cap):
        if cv2 is None: return {}
        return {"width": int(cap.get(cv2.CAP_PROP_FRAME_WIDTH)), "height": int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT)),
                "fps": float(cap.get(cv2.CAP_PROP_FPS)), "fourcc": self.fourcc or ""}

    def _apply(self, cap, validate=True):
        if cv2 is not None:
            if self.fourcc:
                code = str(self.fourcc)[:4]
                cap.set(cv2.CAP_PROP_FOURCC, cv2.VideoWriter_fourcc(*code))
            if self.width: cap.set(cv2.CAP_PROP_FRAME_WIDTH, int(self.width))
            if self.height: cap.set(cv2.CAP_PROP_FRAME_HEIGHT, int(self.height))
            if self.fps: cap.set(cv2.CAP_PROP_FPS, float(self.fps))
        if validate:
            ok, frame = cap.read()
            if not ok or frame is None: raise RuntimeError("可開啟裝置但無法讀取 frame；相機可能被 OBS 或其他程式占用")
            with self._lock: self._frame = frame.copy()
        self.actual = self._actual(cap)

    @property
    def running(self): return bool(self._thread and self._thread.is_alive())

    def start(self):
        if self.running: return
        if cv2 is None and self.capture_factory is None:
            self.state, self.error = "dependency_error", "無法載入 OpenCV；請安裝 opencv-python 並確認 DLL 可正常載入"; return
        self.state = "starting"; self._stop.clear()
        self._thread = threading.Thread(target=self._read_loop, daemon=True, name="ai-camera"); self._thread.start()

    def _read_loop(self):
        cap = None
        try:
            cap = self._open(self.index)
            with self._lock: self._capture = cap
            if cap is None or not cap.isOpened():
                raise RuntimeError("{} 開啟失敗（backend: {}，{}，index {}）；請確認裝置是否被占用".format(BACKEND_LABELS.get(self.backend, self.backend), self.backend, self.device_name or "相機", self.index))
            try: self._apply(cap)
            except RuntimeError:
                # MJPG is an optimisation, not a reason to permanently lose video.
                if str(self.fourcc).upper() == "MJPG": self.fourcc = None; self._apply(cap)
                else: raise
            self.state, self.error = "connected", ""
            while not self._stop.is_set():
                ok, frame = cap.read()
                if ok and frame is not None:
                    with self._lock: self._frame = frame.copy()
                    self.state, self.error = "connected", ""
                else:
                    self.state, self.error = "read_error", "可開啟裝置但無法讀取 frame"
                    self._stop.wait(self.retry_delay)
        except Exception as exc:
            self.state = "dependency_error" if cv2 is None and self.capture_factory is None else "open_error"
            self.error = str(exc)
        finally:
            with self._lock: self._capture = None
            if cap is not None: cap.release()

    def latest_frame(self):
        with self._lock: return None if self._frame is None else self._frame.copy()

    def configure(self, device_id=None, index=None, backend=None, width=None, height=None, fps=None, fourcc=None):
        # Compatibility with the former configure(index, backend) positional API.
        if isinstance(device_id, int): device_id, index = None, device_id
        self.stop()
        if device_id is not None: self.device_id = str(device_id)
        if index is not None: self.index = int(index)
        if backend is not None: self.backend = str(backend).lower()
        for key, value in (("width", width), ("height", height), ("fps", fps), ("fourcc", fourcc)):
            if value is not None: setattr(self, key, value)
        with self._lock: self._frame = None
        self.start()

    def reconnect(self, index=None): self.configure(index=index)

    def stop(self):
        self._stop.set(); thread = self._thread
        if thread and thread is not threading.current_thread(): thread.join(timeout=2)
        self._thread = None
        if self.state not in ("dependency_error", "open_error", "read_error"): self.state = "stopped"
