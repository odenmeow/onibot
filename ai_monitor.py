"""Single-flight automatic recognition and alarm services."""
import os
import queue
import threading
import time


def is_alarm_response(response_text, accept_lowercase=False):
    value = str(response_text or "").strip()
    return value == "O" or (accept_lowercase and value == "o")


def combine_prompt_profiles(profiles):
    """Combine enabled profiles in stored order (never checkbox click order)."""
    selected = [p for p in profiles if p.get("enabled")]
    if not selected:
        raise ValueError("請先選擇至少一個提問配置")
    return "\n\n".join(str(p.get("prompt", "")).strip() for p in selected if str(p.get("prompt", "")).strip())


class AlarmPlayer:
    def __init__(self, sound_path="", player=None):
        self.sound_path = sound_path
        self.player = player or self._platform_play
        self.enabled = True
        self.error = ""
        self._lock = threading.Lock()
        self._playing = False

    @staticmethod
    def _platform_play(path, stop=False):
        if os.name != "nt":
            raise RuntimeError("此平台未配置聲音播放器")
        import winsound
        if stop:
            winsound.PlaySound(None, 0)
        else:
            winsound.PlaySound(path, winsound.SND_FILENAME | winsound.SND_ASYNC)

    def play(self):
        with self._lock:
            if not self.enabled or self._playing:
                return False
            self._playing = True
        def worker():
            failed = False
            try:
                if not self.sound_path or not os.path.isfile(self.sound_path):
                    raise RuntimeError("鬧鐘聲音檔不存在")
                self.player(self.sound_path)
            except Exception as exc:
                self.error = str(exc)
                failed = True
            finally:
                if failed:
                    with self._lock:
                        self._playing = False
        threading.Thread(target=worker, daemon=True).start()
        return True

    def stop(self):
        try:
            self.player(self.sound_path, stop=True)
        except TypeError:
            pass
        except Exception as exc:
            self.error = str(exc)
        with self._lock:
            self._playing = False


class AIMonitor:
    """A generation-token monitor; callbacks are delivered through a queue."""
    def __init__(self, camera, client, profiles, interval=5, alarm=None):
        self.camera, self.client = camera, client
        self.profiles = profiles
        self.interval = max(0.1, float(interval))
        self.alarm = alarm
        self.results = queue.Queue()
        self._lock = threading.Lock()
        self._stop = threading.Event()
        self._thread = None
        self._generation = 0

    @property
    def enabled(self):
        with self._lock:
            return bool(self._thread and self._thread.is_alive() and not self._stop.is_set())

    def start(self):
        prompt = combine_prompt_profiles(self.profiles)
        if self.camera.latest_frame() is None:
            raise ValueError("相機尚無可用畫面")
        with self._lock:
            if self._thread and self._thread.is_alive() and not self._stop.is_set():
                return
            self._generation += 1
            generation = self._generation
            self._stop.clear()
            self._thread = threading.Thread(target=self._loop, args=(generation, prompt), daemon=True)
            self._thread.start()

    def _loop(self, generation, prompt):
        while not self._stop.is_set():
            started = time.time()
            frame = self.camera.latest_frame()
            try:
                if frame is None:
                    raise RuntimeError(self.camera.error or "相機尚無畫面")
                try:
                    import cv2
                    ok, encoded = cv2.imencode(".jpg", frame)
                    if not ok:
                        raise RuntimeError("圖片編碼失敗")
                    image = encoded.tobytes()
                except ImportError as exc:
                    raise RuntimeError("未安裝 OpenCV") from exc
                answer = self.client.chat(prompt, image=image)
                if generation == self._generation and not self._stop.is_set():
                    self.results.put((generation, "answer", {"captured_at": started, "sent_at": time.time(), "text": answer}))
                    if self.alarm and is_alarm_response(answer):
                        self.alarm.play()
            except Exception as exc:
                if generation == self._generation and not self._stop.is_set():
                    self.results.put((generation, "error", str(exc)))
            self._stop.wait(max(0, self.interval - (time.time() - started)))

    def stop(self):
        with self._lock:
            self._generation += 1
            self._stop.set()
            thread = self._thread
        if thread and thread is not threading.current_thread():
            thread.join(timeout=0.2)
        with self._lock:
            self._thread = None
