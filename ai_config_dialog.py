"""Tk AI configuration surface; camera/network work stays in worker modules."""
import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from ai_monitor import AIMonitor, AlarmPlayer, combine_prompt_profiles
from camera_capture import CameraCapture
from image_library import ImageLibrary
from qwen_client import QwenClient


class AIConfigDialog:
    def __init__(self, parent, config, on_save, on_state=None):
        self.config, self.on_save, self.on_state = config, on_save, on_state
        self.window = tk.Toplevel(parent)
        self.window.title("AI 配置")
        self.window.geometry("900x720")
        ai = config.setdefault("ai", {})
        self.camera = CameraCapture(ai.get("camera_index", 0))
        self.library = ImageLibrary(os.path.join(os.path.dirname(__file__), "saved_ai_images"))
        self.alarm = AlarmPlayer(ai.get("sound_path", ""))
        self.monitor = None
        self.displayed_frame = None
        self._after_ids = set()
        self._build(ai)
        self.camera.start()
        self._schedule(100, self._poll_preview)
        self.window.protocol("WM_DELETE_WINDOW", self.close)

    def _build(self, ai):
        preview = ttk.LabelFrame(self.window, text="相機預覽")
        preview.pack(fill="x", padx=8, pady=5)
        self.preview_label = ttk.Label(preview, text="尚無畫面", anchor="center")
        self.preview_label.pack(fill="both", expand=True, pady=4)
        self.mode = tk.StringVar(value=ai.get("preview_mode", "auto"))
        ttk.Radiobutton(preview, text="自動顯示", variable=self.mode, value="auto").pack(side="left")
        ttk.Radiobutton(preview, text="手動顯示", variable=self.mode, value="manual").pack(side="left")
        ttk.Button(preview, text="最新畫面", command=self.show_latest).pack(side="left")
        self.camera_index = tk.IntVar(value=ai.get("camera_index", 0))
        ttk.Combobox(preview, textvariable=self.camera_index, values=(0, 1, 2), width=3, state="readonly").pack(side="left")
        ttk.Button(preview, text="重新連接", command=lambda: self.camera.reconnect(self.camera_index.get())).pack(side="left")
        ttk.Button(preview, text="保存目前圖片", command=self.save_image).pack(side="left")

        qbox = ttk.LabelFrame(self.window, text="本機 Qwen / Ollama")
        qbox.pack(fill="x", padx=8, pady=5)
        self.base_url = self._entry(qbox, "API URL", ai.get("base_url", "http://127.0.0.1:11434"))
        self.model = self._entry(qbox, "模型", ai.get("model", ""))
        self.timeout = self._entry(qbox, "Timeout", ai.get("timeout", 30))
        self.system = self._entry(qbox, "System prompt", ai.get("system_prompt", ""))
        ttk.Button(qbox, text="測試連線", command=self.test_connection).pack(side="left", padx=4)

        profiles = ttk.LabelFrame(self.window, text="提問配置（依列表固定順序組合）")
        profiles.pack(fill="x", padx=8, pady=5)
        self.profile_vars = []
        for profile in ai.get("prompt_profiles", []):
            var = tk.BooleanVar(value=profile.get("enabled", False)); self.profile_vars.append((profile, var))
            ttk.Checkbutton(profiles, text=profile.get("name", "未命名"), variable=var).pack(anchor="w")
        self.user_text = tk.Text(self.window, height=4); self.user_text.pack(fill="x", padx=8)
        ttk.Button(self.window, text="選取圖片", command=self.pick_image).pack(side="left", padx=8)
        ttk.Button(self.window, text="送出測試", command=self.send_test).pack(side="left")
        ttk.Button(self.window, text="AI啟用 / 停用", command=self.toggle_monitor).pack(side="left")
        ttk.Button(self.window, text="停止鬧鐘", command=self.alarm.stop).pack(side="left")
        self.response = tk.Text(self.window, height=10, state="disabled"); self.response.pack(fill="both", expand=True, padx=8, pady=5)
        self.selected_image = None

    @staticmethod
    def _entry(parent, label, value):
        ttk.Label(parent, text=label).pack(side="left", padx=(4, 1))
        entry = ttk.Entry(parent, width=20); entry.insert(0, str(value)); entry.pack(side="left")
        return entry

    def _client(self):
        return QwenClient(self.base_url.get(), self.model.get(), float(self.timeout.get()))

    def _schedule(self, delay, callback):
        holder = {}
        def run():
            self._after_ids.discard(holder.get("id")); callback()
        holder["id"] = self.window.after(delay, run); self._after_ids.add(holder["id"])

    def _poll_preview(self):
        if self.mode.get() == "auto": self.show_latest()
        if self.window.winfo_exists(): self._schedule(100, self._poll_preview)

    def show_latest(self):
        frame = self.camera.latest_frame()
        if frame is None:
            self.preview_label.config(text=self.camera.error or "尚無畫面", image=""); return
        self.displayed_frame = frame
        try:
            from PIL import Image, ImageTk
            import cv2
            rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            image = ImageTk.PhotoImage(Image.fromarray(rgb).resize((480, 270)))
            self.preview_label.config(image=image, text=""); self.preview_label.image = image
        except Exception as exc: self.preview_label.config(text="預覽失敗：{}".format(exc))

    def save_image(self):
        if self.displayed_frame is None: messagebox.showwarning("圖片", "目前沒有顯示中的畫面", parent=self.window); return
        try: self.library.save(self.displayed_frame)
        except Exception as exc: messagebox.showerror("保存失敗", str(exc), parent=self.window)

    def pick_image(self): self.selected_image = filedialog.askopenfilename(parent=self.window, filetypes=[("圖片", "*.jpg *.jpeg *.png")]) or None

    def _worker(self, work):
        def run():
            try: result = (True, work())
            except Exception as exc: result = (False, str(exc))
            self._schedule(0, lambda: self._append(("成功：" if result[0] else "錯誤：") + str(result[1])))
        threading.Thread(target=run, daemon=True).start()

    def test_connection(self): self._worker(lambda: "可用模型：" + ", ".join(self._client().test_connection()))
    def send_test(self):
        profiles = self._sync_profiles()
        try: configured = combine_prompt_profiles(profiles)
        except ValueError: configured = ""
        user = self.user_text.get("1.0", "end").strip()
        prompt = "\n\n".join(x for x in (configured, user) if x)
        if not prompt: messagebox.showwarning("提問", "請輸入文字或選擇提問配置", parent=self.window); return
        self._append("User：" + prompt)
        self._worker(lambda: self._client().chat(prompt, self.selected_image, self.system.get()))

    def _append(self, text):
        self.response.config(state="normal"); self.response.insert("end", text + "\n\n"); self.response.config(state="disabled"); self.response.see("end")

    def _sync_profiles(self):
        for profile, var in self.profile_vars: profile["enabled"] = var.get()
        return [x[0] for x in self.profile_vars]

    def toggle_monitor(self):
        if self.monitor and self.monitor.enabled:
            self.monitor.stop(); self.on_state and self.on_state(False); return
        try:
            client = self._client(); client.test_connection()
            self.monitor = AIMonitor(self.camera, client, self._sync_profiles(), self.config["ai"].get("interval", 5), self.alarm)
            self.monitor.start(); self.on_state and self.on_state(True)
        except Exception as exc: messagebox.showerror("無法啟用 AI", str(exc), parent=self.window)

    def close(self):
        if self.monitor: self.monitor.stop()
        self.camera.stop(); self.alarm.stop()
        for after_id in list(self._after_ids):
            try: self.window.after_cancel(after_id)
            except Exception: pass
        ai = self.config["ai"]
        ai.update({"camera_index": self.camera_index.get(), "preview_mode": self.mode.get(), "base_url": self.base_url.get(), "model": self.model.get(), "timeout": float(self.timeout.get()), "system_prompt": self.system.get(), "prompt_profiles": self._sync_profiles()})
        ai["enabled"] = False; self.on_save(self.config); self.on_state and self.on_state(False)
        self.window.destroy()

