"""Single-flight automatic recognition and alarm services."""
import os
import queue
import socket
import threading
import time
import uuid

from qwen_client import QwenError
from ai_image_processing import prepare_and_encode_ai_image


def is_alarm_response(response_text, accept_lowercase=False):
    value = str(response_text or "").strip()
    return value == "O" or (accept_lowercase and value == "o")


def combine_prompt_profiles(profiles):
    """Combine enabled profiles in stored order (never checkbox click order)."""
    selected = [p for p in profiles if p.get("enabled")]
    if not selected:
        raise ValueError("請先選擇至少一個提問配置")
    return "\n\n".join(str(p.get("prompt", "")).strip() for p in selected if str(p.get("prompt", "")).strip())


def is_timeout_error(exc):
    """Recognise both transport timeouts and Qwen's stable timeout error."""
    current = exc
    while current:
        if isinstance(current, (socket.timeout, TimeoutError)):
            return True
        if isinstance(current, QwenError) and str(current) == "API timeout":
            return True
        current = current.__cause__
    return False


class AlarmPlayer:
    """Play built-in or custom alarms without requiring a sound file."""
    MODES = ("system_alarm", "system_notice", "custom", "mute")

    def __init__(self, sound_path="", player=None, sound_mode="system_alarm"):
        self.sound_path, self.player = sound_path, player
        self.sound_mode = sound_mode if sound_mode in self.MODES else "system_alarm"
        self.error = ""
        self._lock, self._stop = threading.Lock(), threading.Event()
        self._playing = False

    def _beep(self, alarm_type):
        if os.name != "nt":
            # A terminal bell is a harmless best-effort fallback on non-Windows.
            print("\a", end="", flush=True)
            return
        import winsound
        if self.sound_mode == "system_notice" or alarm_type == "error":
            winsound.MessageBeep(winsound.MB_ICONASTERISK)
        else:
            winsound.MessageBeep(winsound.MB_ICONEXCLAMATION)

    def _custom(self, alarm_type):
        if not self.sound_path or not os.path.isfile(self.sound_path):
            self.error = "自訂聲音檔不存在，已改用系統警報聲"
            self._beep(alarm_type)
            return
        if self.player:
            self.player(self.sound_path)
            return
        if os.name != "nt":
            raise RuntimeError("此平台未配置自訂聲音播放器")
        import winsound
        flags = winsound.SND_FILENAME | winsound.SND_ASYNC
        if alarm_type == "detected": flags |= winsound.SND_LOOP
        winsound.PlaySound(self.sound_path, flags)

    def play(self, alarm_type="detected"):
        if alarm_type not in ("detected", "timeout", "error") or self.sound_mode == "mute": return False
        with self._lock:
            if self._playing: return False
            self._playing = True; self._stop.clear(); self.error = ""
        def worker():
            try:
                if self.sound_mode == "custom":
                    self._custom(alarm_type)
                    if alarm_type == "detected" and os.name == "nt" and not self.player:
                        self._stop.wait()
                else:
                    count = 1 if alarm_type == "error" else (3 if alarm_type == "timeout" else None)
                    played = 0
                    while not self._stop.is_set() and (count is None or played < count):
                        self._beep(alarm_type); played += 1
                        self._stop.wait(.25 if alarm_type != "detected" else .6)
            except Exception as exc:
                self.error = str(exc)
            finally:
                with self._lock: self._playing = False
        threading.Thread(target=worker, daemon=True, name="ai-alarm").start()
        return True

    def stop(self):
        self._stop.set()
        try:
            if self.player: self.player(self.sound_path, stop=True)
            elif os.name == "nt":
                import winsound
                winsound.PlaySound(None, 0)
        except TypeError: pass
        except Exception as exc: self.error = str(exc)
        with self._lock: self._playing = False


class AIMonitor:
    """A single-flight monitor that waits *after* each completed attempt."""
    def __init__(self, camera, client, profiles, after_answer_delay=0, alarm=None,
                 alarm_on_detected=True, alarm_on_timeout=True, alarm_on_error=False,
                 stop_on_timeout=False, history_dir=None, alarm_suppression_sec=10,
                 system_prompt="", image_settings=None):
        self.camera, self.client, self.profiles = camera, client, profiles
        self.after_answer_delay = max(0, float(after_answer_delay))
        self.alarm, self.alarm_on_detected = alarm, alarm_on_detected
        self.alarm_on_timeout, self.alarm_on_error = alarm_on_timeout, alarm_on_error
        self.stop_on_timeout, self.history_dir = stop_on_timeout, history_dir
        self.system_prompt = str(system_prompt or "")
        self.image_settings = dict(image_settings or {})
        self.alarm_suppression_sec, self._last_error_alarm = alarm_suppression_sec, {}
        self.results, self._lock, self._stop = queue.Queue(), threading.Lock(), threading.Event()
        self._thread, self._generation, self._last_frame_sequence = None, 0, -1

    @property
    def enabled(self):
        with self._lock: return bool(self._thread and self._thread.is_alive() and not self._stop.is_set())

    def start(self):
        prompt = combine_prompt_profiles(self.profiles)
        if self.camera.latest_frame() is None: raise ValueError("相機尚無可用畫面")
        with self._lock:
            if self._thread and self._thread.is_alive() and not self._stop.is_set(): return
            self._generation += 1; generation = self._generation; self._stop.clear()
            self._thread = threading.Thread(target=self._loop, args=(generation, prompt), daemon=True)
            self._thread.start()

    def _sound_error(self, alarm_type, error_key):
        if not self.alarm: return
        now = time.monotonic()
        if now - self._last_error_alarm.get(error_key, -1e9) >= self.alarm_suppression_sec:
            self._last_error_alarm[error_key] = now; self.alarm.play(alarm_type)

    def _save_capture(self, encoded, started):
        if not self.history_dir: return "", ""
        os.makedirs(self.history_dir, exist_ok=True)
        milliseconds = int(started * 1000) % 1000
        filename = "auto_{}_{:03d}.jpg".format(
            time.strftime("%Y%m%d_%H%M%S", time.localtime(started)), milliseconds)
        path = os.path.join(self.history_dir, filename)
        with open(path, "wb") as stream: stream.write(encoded)
        return filename, path

    def _loop(self, generation, prompt):
        tags = ", ".join(str(p.get("name", "")) for p in self.profiles if p.get("enabled"))
        while not self._stop.is_set():
            started, encoded, filename, path = time.time(), None, "", ""
            frame_sequence, frame_captured_at = None, None
            fresh_reader = getattr(self.camera, "latest_frame_after", None)
            packet = fresh_reader(self._last_frame_sequence, timeout=1.0) if callable(fresh_reader) else None
            if isinstance(packet, tuple) and len(packet) == 3:
                frame, frame_sequence, frame_captured_at = packet
                if frame_sequence is not None: self._last_frame_sequence = frame_sequence
            else:
                frame = self.camera.latest_frame()
            try:
                if frame is None: raise RuntimeError(self.camera.error or "相機尚無畫面")
                try:
                    try:
                        if not self.image_settings:
                            raise TypeError("legacy encoder")
                        with self._lock: image_settings = dict(self.image_settings)
                        encoded, _metadata = prepare_and_encode_ai_image(
                            frame, image_settings, "camera")
                    except TypeError:
                        # Compatibility for synthetic/custom capture objects.
                        import cv2
                        ok, data = cv2.imencode(".jpg", frame)
                        if not ok: raise RuntimeError("圖片編碼失敗")
                        encoded = data.tobytes()
                    filename, path = self._save_capture(encoded, started)
                except ImportError as exc: raise RuntimeError("未安裝 Pillow 或 NumPy") from exc
                answer = self.client.chat(prompt, image=encoded,
                                          system_prompt=self.system_prompt,
                                          cancel_event=self._stop)
                ended = time.time()
                value = {"history_id": uuid.uuid4().hex, "captured_at": started, "sent_at": started, "ended_at": ended,
                         "elapsed_sec": ended - started, "text": answer, "answer": answer,
                         "error": "", "filename": filename, "path": path, "tag": tags,
                         "profile_tags": [str(p.get("id", "")) for p in self.profiles if p.get("enabled")],
                         "mode": "auto", "prompt": prompt, "frame_sequence": frame_sequence,
                         "frame_captured_at": frame_captured_at}
                if generation == self._generation and not self._stop.is_set():
                    self.results.put((generation, "answer", value))
                    if self.alarm_on_detected and is_alarm_response(answer) and self.alarm: self.alarm.play("detected")
            except Exception as exc:
                ended, timeout = time.time(), is_timeout_error(exc)
                error = "AI 回答逾時" if timeout else str(exc)
                value = {"history_id": uuid.uuid4().hex, "captured_at": started, "sent_at": started, "ended_at": ended,
                         "elapsed_sec": ended - started, "text": "", "answer": "", "error": error,
                         "filename": filename, "path": path, "tag": tags, "prompt": prompt,
                         "profile_tags": [str(p.get("id", "")) for p in self.profiles if p.get("enabled")],
                         "mode": "auto", "frame_sequence": frame_sequence,
                         "frame_captured_at": frame_captured_at,
                         "timeout_sec": self.client.timeout if timeout else None}
                if generation == self._generation and not self._stop.is_set():
                    self.results.put((generation, "timeout" if timeout else "error", value))
                    if timeout and self.alarm_on_timeout: self._sound_error("timeout", "timeout")
                    elif not timeout and self.alarm_on_error: self._sound_error("error", type(exc).__name__ + ":" + str(exc))
                if timeout and self.stop_on_timeout: self._stop.set(); break
            # Delay begins only after the answer/error handling has completed.
            self._stop.wait(self.after_answer_delay)

    def stop(self):
        with self._lock: self._generation += 1; self._stop.set(); thread = self._thread
        if thread and thread is not threading.current_thread(): thread.join(timeout=.2)
        with self._lock: self._thread = None

    def update_image_settings(self, settings):
        """Apply a newly saved camera crop to the very next monitor request."""
        with self._lock: self.image_settings = dict(settings or {})
