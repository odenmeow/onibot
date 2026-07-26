"""Labelled Tk AI configuration surface with safe background operations."""
import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from ai_monitor import AIMonitor, AlarmPlayer, combine_prompt_profiles
from camera_capture import CameraCapture
from image_library import ImageLibrary
from qwen_client import QwenClient, choose_model


class AIConfigDialog:
    def __init__(self, parent, config, on_save, on_state=None):
        self.config, self.on_save, self.on_state = config, on_save, on_state
        self.window = tk.Toplevel(parent); self.window.title("AI 配置"); self.window.geometry("1050x780")
        ai = config.setdefault("ai", {})
        self.camera = CameraCapture(ai.get("camera_index", 0), backend=ai.get("camera_backend", "auto"))
        self.library = ImageLibrary(os.path.join(os.path.dirname(__file__), "saved_ai_images"))
        self.alarm = AlarmPlayer(ai.get("sound_path", "")); self.monitor = None
        self.displayed_frame = self.selected_image = self._selected_item = None
        self._after_ids, self._busy, self._generation, self._closed = set(), False, 0, False
        self._build(ai); self.camera.start(); self._schedule(150, self._poll_preview)
        self.window.protocol("WM_DELETE_WINDOW", self.close)

    def _build(self, ai):
        self.window.columnconfigure(0, weight=1); self.window.columnconfigure(1, weight=1); self.window.rowconfigure(2, weight=1)
        preview = ttk.LabelFrame(self.window, text="相機預覽與連線狀態"); preview.grid(row=0, column=0, columnspan=2, sticky="nsew", padx=8, pady=5)
        self.preview_label = ttk.Label(preview, text="尚無畫面", anchor="center"); self.preview_label.grid(row=0, column=0, columnspan=10, sticky="nsew")
        self.camera_status = ttk.Label(preview, text="正在連接相機……"); self.camera_status.grid(row=1, column=0, columnspan=10, sticky="w")
        self.mode = tk.StringVar(value=ai.get("preview_mode", "auto"))
        ttk.Radiobutton(preview, text="自動顯示", variable=self.mode, value="auto").grid(row=2, column=0)
        ttk.Radiobutton(preview, text="手動顯示", variable=self.mode, value="manual").grid(row=2, column=1)
        ttk.Button(preview, text="顯示最新畫面", command=self.show_latest).grid(row=2, column=2)
        ttk.Label(preview, text="裝置索引").grid(row=2, column=3)
        self.camera_index = tk.IntVar(value=ai.get("camera_index", 0)); self.camera_box = ttk.Combobox(preview, textvariable=self.camera_index, values=tuple(range(10)), width=5)
        self.camera_box.grid(row=2, column=4)
        ttk.Label(preview, text="連線方式").grid(row=2, column=5)
        self.backend = tk.StringVar(value=ai.get("camera_backend", "auto")); ttk.Combobox(preview, textvariable=self.backend, values=("auto", "dshow", "msmf", "v4l2"), state="readonly", width=7).grid(row=2, column=6)
        ttk.Button(preview, text="掃描相機", command=self.scan_cameras).grid(row=2, column=7)
        ttk.Button(preview, text="重新連接", command=self.reconnect_camera).grid(row=2, column=8)
        ttk.Button(preview, text="保存目前圖片", command=self.save_image).grid(row=2, column=9)

        qbox = ttk.LabelFrame(self.window, text="Ollama 連線與模型設定"); qbox.grid(row=1, column=0, columnspan=2, sticky="ew", padx=8, pady=5)
        self.base_url = self._entry(qbox, 0, "API 位址", ai.get("base_url", "http://127.0.0.1:11434"))
        ttk.Label(qbox, text="模型").grid(row=0, column=2); self.model = ttk.Combobox(qbox, width=23); self.model.set(ai.get("model", "")); self.model.grid(row=0, column=3)
        self.timeout = self._entry(qbox, 4, "逾時秒數", ai.get("timeout", 30), 8)
        self.system = self._entry(qbox, 6, "系統提示詞", ai.get("system_prompt", ""), 24)
        ttk.Button(qbox, text="檢查 Ollama／自動選模型", command=self.test_connection).grid(row=0, column=8, padx=4)

        left = ttk.Frame(self.window); left.grid(row=2, column=0, sticky="nsew", padx=(8,4)); left.rowconfigure(1, weight=1); left.columnconfigure(0, weight=1)
        images = ttk.LabelFrame(left, text="圖片庫（雙擊切換我的最愛）"); images.grid(row=0, column=0, sticky="ew")
        self.image_info = ttk.Label(images, text="本次提問尚未附加圖片"); self.image_info.pack(fill="x")
        self.attachment_preview = ttk.Label(images, text="無附件", anchor="center"); self.attachment_preview.pack(fill="x")
        for text_, command in (("選取檔案", self.pick_image), ("使用最新畫面", self.use_latest), ("移除附件", self.clear_image), ("重新整理圖片庫", self.refresh_library)):
            ttk.Button(images, text=text_, command=command).pack(side="left")
        self.library_tree = ttk.Treeview(left, columns=("favorite", "time", "source", "note"), show="headings", height=7)
        for key, label in (("favorite", "最愛"), ("time", "保存時間"), ("source", "來源"), ("note", "備註")): self.library_tree.heading(key, text=label)
        self.library_tree.grid(row=1, column=0, sticky="nsew"); self.library_tree.bind("<Double-1>", self.toggle_favorite); self.library_tree.bind("<<TreeviewSelect>>", self.select_library_image)
        profiles = ttk.LabelFrame(left, text="提問配置（依顯示順序組合）"); profiles.grid(row=2, column=0, sticky="ew", pady=5)
        self.profile_vars = []
        for profile in ai.get("prompt_profiles", []):
            var = tk.BooleanVar(value=profile.get("enabled", False)); self.profile_vars.append((profile, var)); ttk.Checkbutton(profiles, text=profile.get("name", "未命名"), variable=var).pack(anchor="w")
        if not self.profile_vars: ttk.Label(profiles, text="尚未建立提問配置；仍可在右側直接輸入問題。").pack(anchor="w")

        right = ttk.Frame(self.window); right.grid(row=2, column=1, sticky="nsew", padx=(4,8)); right.rowconfigure(3, weight=1); right.columnconfigure(0, weight=1)
        ttk.Label(right, text="使用者提問（Ctrl+Enter 送出）").grid(row=0, column=0, sticky="w")
        self.user_text = tk.Text(right, height=5); self.user_text.grid(row=1, column=0, sticky="ew"); self.user_text.bind("<Control-Return>", lambda _e: self.send_test())
        controls = ttk.Frame(right); controls.grid(row=2, column=0, sticky="ew")
        self.send_button = ttk.Button(controls, text="送出提問", command=self.send_test); self.send_button.pack(side="left")
        ttk.Button(controls, text="AI Monitor 啟用／停用", command=self.toggle_monitor).pack(side="left"); ttk.Button(controls, text="停止鬧鐘", command=self.alarm.stop).pack(side="left")
        logbox = ttk.LabelFrame(right, text="對話、狀態與錯誤紀錄"); logbox.grid(row=3, column=0, sticky="nsew")
        self.response = tk.Text(logbox, state="disabled"); self.response.pack(fill="both", expand=True)
        self.refresh_library()

    @staticmethod
    def _entry(parent, column, label, value, width=20):
        ttk.Label(parent, text=label).grid(row=0, column=column); entry = ttk.Entry(parent, width=width); entry.insert(0, str(value)); entry.grid(row=0, column=column+1); return entry

    def _client(self): return QwenClient(self.base_url.get(), self.model.get(), float(self.timeout.get()))
    def _schedule(self, delay, callback):
        if self._closed: return
        holder = {}
        def run(): self._after_ids.discard(holder.get("id")); callback()
        holder["id"] = self.window.after(delay, run); self._after_ids.add(holder["id"])
    def _poll_preview(self):
        if self.camera.error:
            status = self.camera.error
        elif self.camera.state == "connected":
            status = "相機已連接：索引 {}／{}".format(self.camera.index, self.camera.backend)
        else:
            status = "相機狀態：{}".format(self.camera.state)
        self.camera_status.config(text=status)
        if self.mode.get() == "auto": self.show_latest()
        self._schedule(150, self._poll_preview)
    def show_latest(self):
        frame = self.camera.latest_frame()
        if frame is None: return
        self.displayed_frame = frame
        try:
            from PIL import Image, ImageTk
            import cv2
            photo = ImageTk.PhotoImage(Image.fromarray(cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)).resize((480, 270)))
            self.preview_label.config(image=photo, text=""); self.preview_label.image = photo
        except Exception as exc: self.preview_label.config(text="預覽失敗：{}".format(exc), image="")
    def reconnect_camera(self): self.camera.configure(self.camera_index.get(), self.backend.get()); self._append("系統", "正在重新連接相機 {}／{}".format(self.camera_index.get(), self.backend.get()))
    def scan_cameras(self):
        self._worker("掃描相機", lambda: CameraCapture.probe(10, backend=self.backend.get()), self._scanned)
    def _scanned(self, indexes):
        self.camera_box["values"] = indexes or tuple(range(10)); self._append("系統", "找到相機索引：{}".format(", ".join(map(str, indexes)) if indexes else "無"))
    def save_image(self):
        if self.displayed_frame is None: messagebox.showwarning("圖片", "目前沒有顯示中的畫面", parent=self.window); return
        try: self.library.save(self.displayed_frame); self.refresh_library()
        except Exception as exc: messagebox.showerror("保存失敗", str(exc), parent=self.window)
    def pick_image(self):
        path = filedialog.askopenfilename(parent=self.window, filetypes=[("圖片", "*.jpg *.jpeg *.png")]);
        if path: self._select_image(path, "檔案")
    def use_latest(self):
        if self.displayed_frame is None: messagebox.showwarning("圖片", "目前沒有相機畫面", parent=self.window); return
        try: item = self.library.save(self.displayed_frame, source="camera-test"); self.refresh_library(); self._select_image(item["path"], "最新相機畫面")
        except Exception as exc: messagebox.showerror("圖片", str(exc), parent=self.window)
    def _select_image(self, path, source):
        try:
            from PIL import Image, ImageTk
            with Image.open(path) as im:
                size = "{}×{}".format(*im.size)
                thumb = im.convert("RGB"); thumb.thumbnail((320, 160)); photo = ImageTk.PhotoImage(thumb)
        except Exception as exc: messagebox.showerror("圖片無法開啟", str(exc), parent=self.window); return
        self.selected_image = path; self.image_info.config(text="已附加：{}（{}，{}）".format(os.path.basename(path), size, source)); self.attachment_preview.config(image=photo, text=""); self.attachment_preview.image = photo
    def clear_image(self): self.selected_image = None; self.image_info.config(text="本次提問尚未附加圖片"); self.attachment_preview.config(image="", text="無附件"); self.attachment_preview.image = None
    def refresh_library(self):
        if not hasattr(self, "library_tree"): return
        for row in self.library_tree.get_children(): self.library_tree.delete(row)
        for item in reversed(self.library.list()): self.library_tree.insert("", "end", iid=item["id"], values=("★" if item.get("favorite") else "", time_text(item.get("saved_at")), item.get("source", ""), item.get("note", "")))
    def select_library_image(self, _event=None):
        selected = self.library_tree.selection()
        if selected:
            item = next((x for x in self.library.list() if x["id"] == selected[0]), None)
            if item: self._selected_item = item; self._select_image(item["path"], "圖片庫")
    def toggle_favorite(self, _event=None):
        selected = self.library_tree.selection()
        if selected:
            item = next((x for x in self.library.list() if x["id"] == selected[0]), None)
            if item: self.library.set_favorite(item["id"], not item.get("favorite")); self.refresh_library()

    def _worker(self, operation, work, success=None):
        if self._busy: self._append("系統", "已有工作進行中，請稍候"); return
        self._busy = True; self._generation += 1; generation = self._generation; self.send_button.config(state="disabled"); self._append("系統", operation + "中……")
        def run():
            try: result = (True, work())
            except Exception as exc: result = (False, str(exc))
            def done():
                if self._closed or generation != self._generation: return
                self._busy = False; self.send_button.config(state="normal")
                if result[0]:
                    if success: success(result[1])
                    else: self._append("系統", operation + "完成")
                else: self._append("錯誤", result[1])
            self._schedule(0, done)
        threading.Thread(target=run, daemon=True).start()
    def test_connection(self):
        self._worker("檢查 Ollama", self._client().test_connection, self._models_loaded)
    def _models_loaded(self, names):
        self.model["values"] = names; chosen = choose_model(names, self.model.get()); self.model.set(chosen)
        self._append("系統", "Ollama 連線成功；可用模型：{}；已選擇：{}".format(", ".join(names) or "無", chosen or "無"))
    def send_test(self):
        profiles = self._sync_profiles()
        try: configured = combine_prompt_profiles(profiles)
        except ValueError: configured = ""
        user = self.user_text.get("1.0", "end").strip(); prompt = "\n\n".join(x for x in (configured, user) if x)
        if not prompt: messagebox.showwarning("提問", "請輸入文字或選擇提問配置", parent=self.window); return
        if not self.model.get().strip(): messagebox.showwarning("模型", "請先檢查 Ollama 並選擇模型", parent=self.window); return
        self._append("使用者", prompt + ("\n[附件：{}]".format(os.path.basename(self.selected_image)) if self.selected_image else ""))
        self._worker("等待 AI 回答", lambda: self._client().chat(prompt, self.selected_image, self.system.get()), lambda answer: self._append("AI", answer))
    def _append(self, role, text):
        self.response.config(state="normal"); self.response.insert("end", "{}：{}\n\n".format(role, text)); self.response.config(state="disabled"); self.response.see("end")
    def _sync_profiles(self):
        for profile, var in self.profile_vars: profile["enabled"] = var.get()
        return [x[0] for x in self.profile_vars]
    def toggle_monitor(self):
        if self.monitor and self.monitor.enabled: self.monitor.stop(); self.on_state and self.on_state(False); self._append("系統", "AI Monitor 已停用"); return
        try:
            if not self.model.get().strip(): raise ValueError("請先選擇模型")
            self.monitor = AIMonitor(self.camera, self._client(), self._sync_profiles(), self.config["ai"].get("interval", 5), self.alarm); self.monitor.start(); self.on_state and self.on_state(True); self._append("系統", "AI Monitor 已啟用")
        except Exception as exc: messagebox.showerror("無法啟用 AI", str(exc), parent=self.window)
    def close(self):
        self._closed = True; self._generation += 1
        if self.monitor: self.monitor.stop()
        self.camera.stop(); self.alarm.stop()
        for after_id in list(self._after_ids):
            try: self.window.after_cancel(after_id)
            except Exception: pass
        ai = self.config["ai"]; ai.update({"camera_index": self.camera_index.get(), "camera_backend": self.backend.get(), "preview_mode": self.mode.get(), "base_url": self.base_url.get(), "model": self.model.get(), "timeout": float(self.timeout.get()), "system_prompt": self.system.get(), "prompt_profiles": self._sync_profiles(), "enabled": False})
        self.on_save(self.config); self.on_state and self.on_state(False); self.window.destroy()


def time_text(timestamp):
    import time
    try: return time.strftime("%Y-%m-%d %H:%M", time.localtime(float(timestamp)))
    except Exception: return ""
