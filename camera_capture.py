"""Camera discovery, stable device selection and threaded capture.

``CameraCapture`` is the camera entry point used by ``ai_config_dialog.py``;
``front.py`` never opens OpenCV devices directly.  Discovery returns friendly
names and stable device IDs while numeric indexes remain runtime details.
"""
import json
import os
import re
import shutil
import subprocess
import threading
import time

try:
    import cv2
except ImportError:  # camera support is optional on development machines
    cv2 = None


TYPE_LABELS = {"usb": "USB", "integrated": "內建", "virtual": "虛擬", "unknown": "相機"}
BACKEND_LABELS = {"dshow": "DirectShow", "msmf": "Media Foundation", "v4l2": "V4L2", "auto": "自動"}
_UNSET = object()


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

    @staticmethod
    def ffmpeg_path(configured=""):
        path = os.path.expanduser(str(configured or "").strip())
        return path if path and os.path.isfile(path) else shutil.which("ffmpeg")

    @staticmethod
    def parse_dshow_options(output, backend="dshow"):
        """Parse every DirectShow mode emitted by FFmpeg (normally stderr)."""
        modes = []
        # Examples contain either pixel_format=... or vcodec=..., followed by
        # min/max s=WxH fps=N; preserve ranges instead of inventing validation.
        pattern = re.compile(
            r"(?:pixel_format=(?P<pixel>\S+)|vcodec=(?P<codec>\S+)).*?"
            r"min s=(?P<minw>\d+)x(?P<minh>\d+) fps=(?P<minfps>[\d.]+)"
            r"(?: max s=(?P<maxw>\d+)x(?P<maxh>\d+) fps=(?P<maxfps>[\d.]+))?",
            re.IGNORECASE)
        for match in pattern.finditer(str(output or "")):
            item = match.groupdict(); codec = item["codec"] or ""
            fourcc = codec.upper()[:4] if codec else ""
            modes.append({"width": int(item["minw"]), "height": int(item["minh"]),
                "min_width": int(item["minw"]), "min_height": int(item["minh"]),
                "max_width": int(item["maxw"] or item["minw"]),
                "max_height": int(item["maxh"] or item["minh"]),
                "fps": float(item["minfps"]), "min_fps": float(item["minfps"]),
                "max_fps": float(item["maxfps"] or item["minfps"]),
                "pixel_format": item["pixel"] or "", "video_codec": codec,
                "fourcc": fourcc, "backend": backend, "capability_source": "ffmpeg_dshow",
                "validation_status": "裝置回報"})
        return modes

    @classmethod
    def enumerate_dshow_capabilities(cls, device_name, ffmpeg_path="", device_number=None,
                                     timeout=20):
        executable = cls.ffmpeg_path(ffmpeg_path)
        if not executable: raise FileNotFoundError("找不到 FFmpeg；目前只能使用部分探測")
        source = 'video="{}"'.format(str(device_name).replace('"', '\\"'))
        command = [executable, "-hide_banner", "-list_options", "true", "-f", "dshow"]
        if device_number is not None: command += ["-video_device_number", str(int(device_number))]
        command += ["-i", source]
        process = subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                 timeout=timeout, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        # list_options commonly exits non-zero because no output is requested.
        text = process.stderr.decode("utf-8", "replace") + "\n" + process.stdout.decode("utf-8", "replace")
        return cls.parse_dshow_options(text)

    def __init__(self, index=0, capture_factory=None, retry_delay=.1, backend="auto", **settings):
        self.index, self.backend = int(index), str(backend or "auto").lower()
        self.device_id = str(settings.get("device_id") or "")
        self.device_name = str(settings.get("camera_name") or "")
        self.width = settings.get("width"); self.height = settings.get("height")
        self.fps = settings.get("fps"); self.fourcc = settings.get("fourcc")
        self.capture_factory, self.retry_delay = capture_factory, retry_delay
        self.actual = {}; self.error = ""; self.state = "stopped"
        self.last_operation = ""; self.fallback = {}
        self._lock = threading.Lock(); self._stop = threading.Event()
        self._thread = self._capture = self._frame = None

    def _open(self, index, backend=None):
        if self.capture_factory is not None: return self.capture_factory(index)
        if cv2 is None: raise RuntimeError("無法載入 OpenCV；請安裝 opencv-python 並確認 DLL 可正常載入")
        constant = self.BACKENDS.get(backend or self.backend)
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
            # PnP and DirectShow do not promise the same ordering.  Never bind a
            # PnP identity to an OpenCV index merely because both are Nth.
            info = pnp[position] if len(pnp) == 1 and len(indexes) == 1 else {}
            name = str(info.get("FriendlyName") or ("Camera {}（名稱對應未驗證）".format(index)))
            device_id = str(info.get("InstanceId") or "{}:index:{}".format(backend, index))
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
                if cap is not None:
                    try: cap.release()
                    except Exception: pass
        return found

    @classmethod
    def probe_capabilities(cls, device, capture_factory=None, cancelled=None):
        """Verify common settings with a real frame; return unique actual modes."""
        results = []
        if cv2 is None and capture_factory is None: return results
        for width, height in cls.COMMON_RESOLUTIONS:
            for fps in cls.COMMON_FPS:
                for fourcc in (("MJPG", None) if (width, height) == (1920, 1080) else (None,)):
                    if cancelled and cancelled(): return results
                    reader = cls(device["index"], capture_factory, backend=device.get("backend", "auto"),
                                 width=width, height=height, fps=fps, fourcc=fourcc)
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
                        if cap is not None:
                            try: cap.release()
                            except Exception: pass
                        if cancelled:
                            deadline = time.monotonic() + .15
                            while time.monotonic() < deadline:
                                if cancelled(): return results
                                time.sleep(.02)
                        else: time.sleep(.15)
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
        values = {}
        for name, prop in (("width", cv2.CAP_PROP_FRAME_WIDTH), ("height", cv2.CAP_PROP_FRAME_HEIGHT),
                           ("fps", cv2.CAP_PROP_FPS), ("fourcc", cv2.CAP_PROP_FOURCC)):
            try: values[name] = cap.get(prop)
            except Exception as exc: self._record_error("cap.get({})".format(name), exc); values[name] = 0
        code = int(values["fourcc"] or 0)
        values["fourcc"] = "".join(chr((code >> (8 * i)) & 0xff) for i in range(4)).strip("\x00 ")
        values["width"], values["height"] = int(values["width"] or 0), int(values["height"] or 0)
        values["fps"] = float(values["fps"] or 0) or None
        return values

    def _context(self):
        return "backend={} index={} width={} height={} fps={} fourcc={}".format(
            self.backend, self.index, self.width, self.height, self.fps, self.fourcc)

    def _record_error(self, operation, exc):
        self.last_operation = operation
        self.error = "相機操作失敗（{}；{}）：{}".format(operation, self._context(), exc)

    def _apply(self, cap, validate=True):
        if cv2 is not None:
            if self.fourcc:
                code = str(self.fourcc)[:4]
                try: value = cv2.VideoWriter_fourcc(*code)
                except Exception as exc: self._record_error("VideoWriter_fourcc", exc); raise
                try: cap.set(cv2.CAP_PROP_FOURCC, value)
                except Exception as exc: self._record_error("cap.set(fourcc)", exc); raise
            for name, prop, value in (("width", cv2.CAP_PROP_FRAME_WIDTH, self.width), ("height", cv2.CAP_PROP_FRAME_HEIGHT, self.height), ("fps", cv2.CAP_PROP_FPS, self.fps)):
                if value is not None:
                    try: cap.set(prop, float(value))
                    except Exception as exc: self._record_error("cap.set({})".format(name), exc); raise
        if validate:
            try: ok, frame = cap.read()
            except Exception as exc: self._record_error("cap.read(validate)", exc); raise
            if not ok or frame is None: raise RuntimeError("可開啟裝置但無法讀取 frame；相機可能被 OBS 或其他程式占用")
            with self._lock: self._frame = frame.copy()
        self.actual = self._actual(cap)

    @property
    def running(self): return bool(self._thread and self._thread.is_alive())

    def start(self):
        if self.running or (self._thread is not None and self._thread.is_alive()): return False
        if cv2 is None and self.capture_factory is None:
            self.state, self.error = "dependency_error", "無法載入 OpenCV；請安裝 opencv-python 並確認 DLL 可正常載入"; return False
        self.state = "starting"; self._stop.clear()
        self._thread = threading.Thread(target=self._read_loop, daemon=True, name="ai-camera"); self._thread.start()
        return True

    def _read_loop(self):
        cap = None
        try:
            cap = self._open_with_fallback()
            with self._lock: self._capture = cap
            if cap is None or not cap.isOpened():
                raise RuntimeError("{} 開啟失敗（backend: {}，{}，index {}）；請確認裝置是否被占用".format(BACKEND_LABELS.get(self.backend, self.backend), self.backend, self.device_name or "相機", self.index))
            self.state, self.error = "connected", ""
            while not self._stop.is_set():
                try: ok, frame = cap.read()
                except Exception as exc:
                    self._record_error("cap.read(loop)", exc); self.state = "read_error"; break
                if ok and frame is not None:
                    with self._lock: self._frame = frame.copy()
                    self.state, self.error = "connected", ""
                else:
                    self.state, self.error = "read_error", "可開啟裝置但無法讀取 frame"
                    self._stop.wait(self.retry_delay)
        except Exception as exc:
            self.state = "dependency_error" if cv2 is None and self.capture_factory is None else "configure_error"
            if not self.error: self._record_error("VideoCapture/open/configure", exc)
        finally:
            with self._lock: self._capture = None
            if cap is not None:
                try: cap.release()
                except Exception as exc: self._record_error("cap.release", exc)

    def _open_with_fallback(self):
        # A failed attempt is fully released before opening the same index again.
        attempts = [(self.backend, self.fourcc, self.width, self.height, self.fps)]
        if os.name == "nt" and self.backend == "dshow" and self.capture_factory is None:
            attempts += [("dshow", None, self.width, self.height, self.fps), ("dshow", None, None, None, None),
                         ("msmf", None, None, None, None), ("auto", None, None, None, None)]
        original = (self.backend, self.fourcc, self.width, self.height, self.fps)
        last = None
        for backend, fourcc, width, height, fps in attempts:
            cap = None
            try:
                self.backend, self.fourcc, self.width, self.height, self.fps = backend, fourcc, width, height, fps
                cap = self._open(self.index, backend)
                if cap is None or not cap.isOpened(): raise RuntimeError("裝置無法開啟")
                self._apply(cap)
                self.fallback = {"backend": backend, "width": width, "height": height, "fps": fps, "fourcc": fourcc}
                self.backend, self.fourcc, self.width, self.height, self.fps = original
                return cap
            except Exception as exc:
                last = exc
                if cap is not None:
                    try: cap.release()
                    except Exception: pass
                time.sleep(.15)
        self.backend, self.fourcc, self.width, self.height, self.fps = original
        raise last or RuntimeError("相機開啟失敗")

    def latest_frame(self):
        with self._lock: return None if self._frame is None else self._frame.copy()

    def configure(self, device_id=_UNSET, index=_UNSET, backend=_UNSET, width=_UNSET, height=_UNSET, fps=_UNSET, fourcc=_UNSET):
        # Compatibility with the former configure(index, backend) positional API.
        if isinstance(device_id, int): device_id, index = _UNSET, device_id
        if not self.stop(): return False
        if device_id is not _UNSET: self.device_id = str(device_id or "")
        if index is not _UNSET: self.index = int(index)
        if backend is not _UNSET: self.backend = str(backend or "auto").lower()
        for key, value in (("width", width), ("height", height), ("fps", fps), ("fourcc", fourcc)):
            if value is not _UNSET: setattr(self, key, value)
        with self._lock: self._frame = None
        return self.start()

    def reconnect(self, index=None): self.configure(index=index)

    def stop(self):
        self._stop.set(); thread = self._thread
        if thread and thread is not threading.current_thread(): thread.join(timeout=2)
        if thread and thread.is_alive():
            self.state, self.error = "stop_timeout", "相機仍在釋放中，請稍候再重新套用設定"
            return False
        self._thread = None
        if self.state not in ("dependency_error", "open_error", "read_error"): self.state = "stopped"
        return True
