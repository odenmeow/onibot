"""AI configuration window and all AI-facing text/settings management.

``AIConfigDialog`` is created only by ``front.py``.  It owns camera/monitor
threads, prompt profiles, draft autosave and image viewing, then persists via
the save callback supplied by the main application.
"""
import os
import threading
import time
import tkinter as tk
import uuid
from tkinter import filedialog, messagebox, ttk

from ai_monitor import AIMonitor, AlarmPlayer, combine_prompt_profiles
from camera_capture import BACKEND_LABELS, TYPE_LABELS, CameraCapture
from image_library import ImageLibrary
from ai_image_processing import map_canvas_point_to_image, prepare_and_encode_ai_image
from qwen_client import QwenClient, choose_model
from ollama_runtime_options import build_ollama_options, QWEN_OX_PRESET


QUESTION_HISTORY_LIMIT = 1000


class AIConfigDialog:
    def __init__(self, parent, config, on_save, on_state=None):
        self.config, self.on_save, self.on_state = config, on_save, on_state
        self.window = tk.Toplevel(parent); self.window.title("AI 配置")
        ai = config.setdefault("ai", {})
        self.ai_layout = ai.setdefault("ai_window_layout", {})
        self.window.geometry(self.ai_layout.get("geometry", "1180x850"))
        self.camera = CameraCapture(ai.get("camera_index", 0), backend=ai.get("camera_backend", "dshow"),
            device_id=ai.get("camera_device_id", ""), camera_name=ai.get("camera_name", ""),
            width=ai.get("camera_width"), height=ai.get("camera_height"), fps=ai.get("camera_fps"), fourcc=ai.get("camera_fourcc"))
        self.library = ImageLibrary(os.path.join(os.path.dirname(__file__), "saved_ai_images"))
        self.alarm = AlarmPlayer(ai.get("sound_path", ""), sound_mode=ai.get("sound_mode", "system_alarm"))
        self.monitor = None; self.devices = []; self.capabilities = []
        self.displayed_frame = self.selected_image = self._selected_item = None
        self.zoom_window = self._viewer_image = self._viewer_photo = None
        self._viewer_scale = self._viewer_offset = self._viewer_drag = None
        self._viewer_after = self.detached_window = self.detached_preview = None
        self._detached_photo = self._detached_scale = self._detached_offset = self._detached_drag = None
        self._history_by_id = {}; self.camera_view_state = self.ai_layout.get("camera_state", "docked")
        self._after_ids, self._busy, self._generation, self._closed = set(), False, 0, False
        self._draft_after = self._system_after = None
        self._sections = {}
        self._probe_cancel = threading.Event()
        # Attachment crops are deliberately per-selection.  A crop saved for a
        # previous manual image must never silently affect the next attachment.
        self.attachment_crop_roi = None
        self.profiles = [dict(x) for x in ai.get("prompt_profiles", []) if isinstance(x, dict)]
        self._build(ai); self._schedule(100, self.scan_cameras); self._schedule(150, self._poll_preview)
        self._schedule(80, self._restore_ai_layout)
        self.window.protocol("WM_DELETE_WINDOW", self.close)

    def _build(self, ai):
        self.window.columnconfigure((0, 1), weight=1); self.window.rowconfigure(2, weight=1)
        preview = self._section(self.window, "相機裝置與預覽", "選擇相機、顯示方式及要求的解析度、FPS 與 FourCC；實際規格取決於驅動與裝置支援。", row=0, column=0, columnspan=2, sticky="nsew", padx=8, pady=5)
        preview.columnconfigure(0, weight=1)
        self.preview_label = ttk.Label(preview, text="尚無畫面", anchor="center"); self.preview_label.grid(row=0, column=0, columnspan=12, sticky="nsew")
        self.preview_label.bind("<Double-Button-1>", self._on_camera_double_click)
        self.camera_status = ttk.Label(preview, text="正在背景掃描相機……", justify="left"); self.camera_status.grid(row=1, column=0, columnspan=12, sticky="w")
        self.mode = tk.StringVar(value=ai.get("preview_mode", "auto"))
        ttk.Radiobutton(preview, text="自動顯示", variable=self.mode, value="auto").grid(row=2, column=0)
        ttk.Radiobutton(preview, text="手動顯示", variable=self.mode, value="manual").grid(row=2, column=1)
        self.device_choice = tk.StringVar(); self.camera_box = ttk.Combobox(preview, textvariable=self.device_choice, state="readonly", width=37)
        ttk.Label(preview, text="相機裝置").grid(row=2, column=2); self.camera_box.grid(row=2, column=3); self.camera_box.bind("<<ComboboxSelected>>", self._device_changed)
        self.resolution = tk.StringVar(value=self._resolution_text(ai.get("camera_width"), ai.get("camera_height")))
        ttk.Label(preview, text="解析度").grid(row=2, column=4); self.res_box = ttk.Combobox(preview, textvariable=self.resolution, state="readonly", width=12); self.res_box.grid(row=2, column=5)
        self.fps = tk.StringVar(value=str(ai.get("camera_fps") or "自動")); ttk.Label(preview, text="FPS").grid(row=2, column=6); self.fps_box = ttk.Combobox(preview, textvariable=self.fps, state="readonly", width=7); self.fps_box.grid(row=2, column=7)
        self.res_box["values"] = ("自動", "640 × 480", "1280 × 720", "1920 × 1080")
        self.fps_box["values"] = ("自動", "15", "30", "60")
        self.fourcc = tk.StringVar(value=ai.get("camera_fourcc", "MJPG") or "自動"); ttk.Label(preview, text="FourCC").grid(row=2, column=8); ttk.Combobox(preview, textvariable=self.fourcc, values=("自動", "MJPG", "YUY2"), width=7).grid(row=2, column=9)
        ttk.Button(preview, text="重新掃描", command=self.scan_cameras).grid(row=2, column=10)
        ttk.Button(preview, text="套用相機設定", command=self.apply_camera).grid(row=2, column=11)
        camera_actions = ttk.Frame(preview); camera_actions.grid(row=3, column=0, columnspan=6, sticky="w")
        ttk.Button(camera_actions, text="顯示預覽", command=lambda: self.set_camera_view("docked")).pack(side="left")
        ttk.Button(camera_actions, text="隱藏預覽", command=lambda: self.set_camera_view("hidden")).pack(side="left")
        ttk.Button(camera_actions, text="分離預覽", command=lambda: self.set_camera_view("detached")).pack(side="left")
        ttk.Button(camera_actions, text="裁切畫面", command=self.crop_camera).pack(side="left")
        ttk.Button(preview, text="完整偵測裝置規格", command=self.probe_capabilities).grid(row=3, column=10, columnspan=2)

        qbox = self._section(self.window, "Ollama 與監控設定", "設定 Ollama 位址與模型、回答逾時、每輪等待時間及三種警報。停止鬧鐘會立即停止目前聲音。", row=1, column=0, columnspan=2, sticky="ew", padx=8)
        self.base_url = self._entry(qbox, 0, "API 位址", ai.get("base_url", "http://127.0.0.1:11434"), 25)
        ttk.Label(qbox, text="模型").grid(row=0, column=2); self.model = ttk.Combobox(qbox, width=20); self.model.set(ai.get("model", "qwen3-vl:8b")); self.model.grid(row=0, column=3)
        ttk.Button(qbox, text="檢查 Ollama", command=self.test_connection).grid(row=0, column=4, padx=4)
        ttk.Button(qbox, text="預先載入模型", command=self.preload_model).grid(row=0, column=5, padx=4)
        self.timeout = self._entry(qbox, 0, "AI 回答逾時", ai.get("timeout", 30), 6, row=1, suffix="秒")
        self.after_answer_delay = self._entry(qbox, 3, "回答完成後等待", ai.get("after_answer_delay", ai.get("interval", 0)), 6, row=1, suffix="秒")
        ttk.Label(qbox, text="模型保留時間").grid(row=1, column=6, sticky="e")
        self.keep_alive = ttk.Combobox(qbox, width=8, values=("5m", "15m", "30m", "1h", "-1"))
        self.keep_alive.set(ai.get("keep_alive", "30m")); self.keep_alive.grid(row=1, column=7, sticky="w")
        self.think = tk.BooleanVar(value=ai.get("think", False))
        ttk.Checkbutton(qbox, text="啟用思考", variable=self.think).grid(row=2, column=6, sticky="w")
        self.num_predict = self._entry(qbox, 6, "最多輸出 token", ai.get("num_predict", 1024), 6, row=3)
        ttk.Button(qbox, text="模型進階參數……", command=self.open_model_options).grid(row=4, column=6, sticky="ew")
        ttk.Button(qbox, text="AI 影像輸入設定……", command=self.open_image_options).grid(row=4, column=7, sticky="ew")
        self._tooltip(self.keep_alive, "只讓模型本體留在記憶體以加快下次回答，不保留任何前次提問、圖片或回答；-1 表示直到 Ollama 停止")
        self._tooltip(self.timeout, "單次送出圖片後，最多等待 AI 回答的時間")
        self._tooltip(self.after_answer_delay, "AI 回答或錯誤處理完成後，再等待幾秒開始下一輪；0 表示立刻繼續")
        self.alarm_on_detected = tk.BooleanVar(value=ai.get("alarm_on_detected", ai.get("sound_enabled", True)))
        self.alarm_on_timeout = tk.BooleanVar(value=ai.get("alarm_on_timeout", True))
        self.alarm_on_error = tk.BooleanVar(value=ai.get("alarm_on_error", False))
        for col, (label, variable, tip) in enumerate((("偵測到 O 時警報", self.alarm_on_detected, "AI 回答 O 時持續警報"), ("AI 回答逾時時警報", self.alarm_on_timeout, "AI 等待超過設定秒數時短促警報"), ("系統錯誤時警報", self.alarm_on_error, "相機、Ollama 或圖片錯誤時提示；相同錯誤 10 秒內一次"))):
            widget = ttk.Checkbutton(qbox, text=label, variable=variable); widget.grid(row=2, column=col, columnspan=2, sticky="w"); self._tooltip(widget, tip)
        ttk.Label(qbox, text="警報聲音").grid(row=3, column=0, sticky="e")
        self.sound_mode = tk.StringVar(value=ai.get("sound_mode", "system_alarm")); self.sound_mode_box = ttk.Combobox(qbox, state="readonly", width=14, textvariable=self.sound_mode,
            values=("系統警報聲", "系統提示音", "自訂聲音檔", "靜音")); self.sound_mode_box.grid(row=3, column=1, sticky="w"); self.sound_mode_box.bind("<<ComboboxSelected>>", self._sound_mode_changed)
        mode_labels = {"system_alarm": "系統警報聲", "system_notice": "系統提示音", "custom": "自訂聲音檔", "mute": "靜音"}
        self.sound_mode.set(mode_labels.get(ai.get("sound_mode", "system_alarm"), "系統警報聲"))
        self.sound_file_label = ttk.Label(qbox, text=os.path.basename(ai.get("sound_path", "")) or "尚未選擇檔案", width=22)
        self.sound_file_label.grid(row=3, column=2, sticky="w"); self.sound_file_label._full_path = ai.get("sound_path", "")
        self.sound_path = tk.StringVar(value=ai.get("sound_path", ""))
        self.pick_sound_button = ttk.Button(qbox, text="選擇聲音檔", command=self.pick_sound); self.pick_sound_button.grid(row=3, column=3)
        self.test_sound_button = ttk.Button(qbox, text="測試聲音", command=self.test_sound); self.test_sound_button.grid(row=3, column=4)
        ttk.Button(qbox, text="停止鬧鐘", command=self.alarm.stop).grid(row=3, column=5)
        self.sound_status = ttk.Label(qbox, text=""); self.sound_status.grid(row=4, column=0, columnspan=6, sticky="w")
        self._sound_mode_changed()

        self.main_paned = tk.PanedWindow(self.window, orient=tk.HORIZONTAL, sashrelief=tk.RAISED)
        self.main_paned.grid(row=2, column=0, columnspan=2, sticky="nsew", padx=8)
        left = ttk.Frame(self.main_paned); left.columnconfigure(0, weight=1); left.rowconfigure(0, weight=1)
        self._layout_panes = {"left": left}
        self.left_paned = tk.PanedWindow(left, orient=tk.VERTICAL, sashrelief=tk.RAISED,
                                         sashwidth=6, showhandle=True)
        self.left_paned.grid(row=0, column=0, sticky="nsew")
        sysbox = self._section(self.left_paned, "共用系統提示詞（選填）", "用來設定 AI 永遠都要遵守的角色、回答格式或共同規則。送出手動提問與 AI Monitor 每一輪自動判斷時都會套用；若提問配置已包含全部規則，這裡可以留白。")
        system_scroll = ttk.Scrollbar(sysbox, orient="vertical")
        self.system = tk.Text(sysbox, height=4, yscrollcommand=system_scroll.set)
        system_scroll.configure(command=self.system.yview)
        system_scroll.pack(side="right", fill="y")
        self.system.pack(side="left", fill="both", expand=True)
        self.system.insert("1.0", ai.get("system_prompt", ""))
        self.system.bind("<KeyRelease>", self._system_changed)
        pbox = self._section(self.left_paned, "提問範本（手動與自動共用）", "這裡可保存多組常用命令，不必記住以前的寫法。啟用一組就是選用該範本；也可同時啟用多組，系統會依畫面順序合併。手動送出與 AI Monitor 都會使用目前啟用的範本。")
        pbox.rowconfigure(0, weight=1)
        self.profile_tree = ttk.Treeview(pbox, columns=("enabled", "name"), show="headings", height=4); self.profile_tree.heading("enabled", text="狀態"); self.profile_tree.heading("name", text="名稱")
        profile_scroll = ttk.Scrollbar(pbox, orient="vertical", command=self.profile_tree.yview)
        self.profile_tree.configure(yscrollcommand=profile_scroll.set)
        self.profile_tree.grid(row=0, column=0, sticky="nsew")
        profile_scroll.grid(row=0, column=1, sticky="ns")
        # Keep the actions in the block when its lower sash is dragged upward.
        # A two-row grid also avoids clipping the last action in a narrow pane.
        bar = ttk.Frame(pbox); bar.grid(row=1, column=0, sticky="ew")
        profile_actions = (("新增提問配置", self.add_profile), ("編輯", self.edit_profile), ("複製", self.copy_profile), ("刪除", self.delete_profile), ("上移", lambda: self.move_profile(-1)), ("下移", lambda: self.move_profile(1)), ("啟用／停用", self.toggle_profile))
        for index, (label, command) in enumerate(profile_actions):
            row, column = divmod(index, 4)
            bar.columnconfigure(column, weight=1, uniform="profile-action")
            ttk.Button(bar, text=label, command=command).grid(row=row, column=column, sticky="ew")
        self._refresh_profiles()
        images = self._section(self.left_paned, "圖片庫／附件", "選擇手動提問的圖片，或管理從相機保存的圖片；雙擊開啟後，可用滾輪以滑鼠位置為中心縮放。")
        # Do not put these unequal columns in one ``uniform`` group.  Tk uses
        # the widgets' requested widths when sizing a uniform group, which can
        # make the library column wider than its pane and clip its right edge.
        images.columnconfigure(0, weight=2)
        images.columnconfigure(1, weight=1)
        images.rowconfigure(0, weight=1)
        attachment = ttk.Frame(images); attachment.grid(row=0, column=0, sticky="nsew", padx=(0, 4))
        attachment.columnconfigure(0, weight=1); attachment.rowconfigure(1, weight=1)
        self.image_info = ttk.Label(attachment, text="本次提問尚未附加圖片"); self.image_info.grid(row=0, column=0, sticky="ew")
        self.attachment_crop_status = ttk.Label(attachment, text="裁切：未套用（手動附件每次重新選圖都會重設）")
        self.attachment_crop_status.grid(row=3, column=0, sticky="w")
        self.attachment_preview = ttk.Label(attachment, text="無附件", anchor="center"); self.attachment_preview.grid(row=1, column=0, sticky="nsew")
        self.attachment_preview.bind("<Double-Button-1>", self._on_attachment_double_click)
        image_actions = ttk.Frame(attachment); image_actions.grid(row=2, column=0, sticky="ew")
        for index, (label, command) in enumerate((("選取檔案", self.pick_image), ("使用最新畫面", self.use_latest), ("附件裁切", self.crop_attachment), ("保存目前圖片", self.save_image))):
            row, column = divmod(index, 2)
            image_actions.columnconfigure(column, weight=1, uniform="image-action")
            ttk.Button(image_actions, text=label, command=command).grid(row=row, column=column, sticky="ew")
        library = ttk.Frame(images)
        library.grid(row=0, column=1, sticky="nsew")
        library.columnconfigure(0, weight=1); library.rowconfigure(0, weight=1)
        self.library_tree = ttk.Treeview(library, columns=("favorite", "time", "source"), show="headings", height=5)
        for key, label in (("favorite", "最愛"), ("time", "保存時間"), ("source", "來源")): self.library_tree.heading(key, text=label)
        self._configure_library_columns(self.library_tree)
        library_scroll = ttk.Scrollbar(library, orient="horizontal", command=self.library_tree.xview)
        library_vertical_scroll = ttk.Scrollbar(library, orient="vertical", command=self.library_tree.yview)
        self.library_tree.configure(xscrollcommand=library_scroll.set, yscrollcommand=library_vertical_scroll.set)
        self.library_tree.grid(row=0, column=0, sticky="nsew")
        library_vertical_scroll.grid(row=0, column=1, sticky="ns")
        library_scroll.grid(row=1, column=0, sticky="ew")
        self.library_tree.bind("<<TreeviewSelect>>", self.select_library_image)
        self.library_tree.bind("<Double-Button-1>", self._on_library_double_click); self.refresh_library()
        for content, minimum in ((sysbox, 75), (pbox, 150), (images, 160)):
            self.left_paned.add(content._section_outer, minsize=minimum, stretch="always")
            self._configure_section_pane(content, minimum)

        right = ttk.Frame(self.main_paned); right.columnconfigure(0, weight=1); right.rowconfigure(0, weight=1)
        self._layout_panes["right"] = right
        # Left and right now have stable meanings.  The sash is the only horizontal
        # layout control, avoiding the accidental pane swaps caused by drag handles.
        self.main_paned.add(left, minsize=320, stretch="always")
        self.main_paned.add(right, minsize=320, stretch="always")
        self.right_paned = tk.PanedWindow(right, orient=tk.VERTICAL, sashrelief=tk.RAISED,
                                          sashwidth=6, showhandle=True)
        self.right_paned.grid(row=0, column=0, sticky="nsew")
        question = self._section(self.right_paned, "手動追加提問", "僅在手動送出時追加；按 Ctrl+Enter 可直接送出。")
        question_text = ttk.Frame(question); question_text.pack(fill="both", expand=True)
        question_scroll = ttk.Scrollbar(question_text, orient="vertical")
        self.user_text = tk.Text(question_text, height=7, yscrollcommand=question_scroll.set)
        question_scroll.configure(command=self.user_text.yview)
        question_scroll.pack(side="right", fill="y"); self.user_text.pack(side="left", fill="both", expand=True); self.user_text.insert("1.0", ai.get("user_prompt_draft", ""))
        self.user_text.bind("<KeyRelease>", self._draft_changed); self.user_text.bind("<Control-Return>", self._send_shortcut)
        # Actions intentionally live in their own pane.  Keeping them inside the
        # editor made every manual/monitor action disappear when the question
        # pane was collapsed or its sash was dragged upward.
        controls = self._section(self.right_paned, "提問與監控操作", "這些操作獨立於手動追加提問區；收合或縮小提問內容後仍可使用。")
        manual_actions = (("送出提問", self.send_test), ("清除提問", self.clear_draft),
                          ("清除 Console", self.clear_console), ("啟用 AI Monitor", self.toggle_monitor),
                          ("停止鬧鐘", self.alarm.stop))
        for index, (label, command) in enumerate(manual_actions):
            controls.columnconfigure(index, weight=1, uniform="manual-action")
            button = ttk.Button(controls, text=label, command=command)
            button.grid(row=0, column=index, sticky="ew")
            if index == 0: self.send_button = button
            elif index == 3: self.monitor_button = button
        console = self._section(self.right_paned, "AI Monitor／Console", "顯示監控狀態、系統訊息與 AI 回答。")
        self.monitor_status = ttk.Label(console, text="AI Monitor：已停止"); self.monitor_status.pack(fill="x")
        console_text = ttk.Frame(console); console_text.pack(fill="both", expand=True)
        console_scroll = ttk.Scrollbar(console_text, orient="vertical")
        self.response = tk.Text(console_text, state="disabled", yscrollcommand=console_scroll.set)
        console_scroll.configure(command=self.response.yview)
        console_scroll.pack(side="right", fill="y"); self.response.pack(side="left", fill="both", expand=True)
        history_box = self._section(self.right_paned, "最近 1000 次提問歷史", "包含 AI Monitor 自動判斷與「送出提問」的手動測試；雙擊可查看當次保存的圖片。")
        history_box.rowconfigure(0, weight=1)
        self.history_tree = ttk.Treeview(history_box, columns=("summary",), show="headings", height=5); self.history_tree.heading("summary", text="時間｜tag｜圖片｜結果")
        history_vertical_scroll = ttk.Scrollbar(history_box, orient="vertical", command=self.history_tree.yview)
        history_horizontal_scroll = ttk.Scrollbar(history_box, orient="horizontal", command=self.history_tree.xview)
        self.history_tree.configure(yscrollcommand=history_vertical_scroll.set, xscrollcommand=history_horizontal_scroll.set)
        self.history_tree.grid(row=0, column=0, sticky="nsew")
        history_vertical_scroll.grid(row=0, column=1, sticky="ns")
        history_horizontal_scroll.grid(row=1, column=0, sticky="ew")
        self.history_tree.bind("<Double-Button-1>", self._on_history_double_click); self._refresh_history()
        for content, minimum in ((question, 90), (controls, 62), (console, 120), (history_box, 110)):
            self.right_paned.add(content._section_outer, minsize=minimum, stretch="always")
            self._configure_section_pane(content, minimum)
        bottom = ttk.Frame(self.window); bottom.grid(row=3, column=0, columnspan=2, sticky="e", padx=8, pady=5)
        ttk.Button(bottom, text="保存設定", command=self.save_settings).pack(side="left")
        ttk.Button(bottom, text="保存 UI 配置", command=self.save_ui_layout).pack(side="left")
        ttk.Button(bottom, text="套用並重新連接相機", command=lambda: self.apply_camera(save=True)).pack(side="left")
        ttk.Button(bottom, text="關閉", command=self.close).pack(side="left")

    @staticmethod
    def _configure_library_columns(tree):
        """Keep every image-library heading reachable in a narrow pane."""
        tree.column("favorite", width=45, minwidth=40, stretch=False)
        tree.column("time", width=105, minwidth=75, stretch=True)
        tree.column("source", width=60, minwidth=45, stretch=True)

    def _section(self, parent, title, help_text, **grid_options):
        """Create a compact section whose explanation is available on demand."""
        outer = ttk.Frame(parent, relief="groove", borderwidth=1)
        if grid_options: outer.grid(**grid_options)
        outer.columnconfigure(0, weight=1)
        header = ttk.Frame(outer)
        header.grid(row=0, column=0, sticky="ew", padx=5, pady=(2, 0))
        collapsed = bool(self.ai_layout.get("collapsed", {}).get(title, False))
        toggle = ttk.Button(header, text="▶" if collapsed else "▼", width=2)
        toggle.pack(side="left")
        ttk.Label(header, text=title).pack(side="left", padx=(3, 0))
        help_button = ttk.Button(header, text="?", width=2, takefocus=True)
        help_button.pack(side="left", padx=(4, 0))
        self._tooltip(help_button, help_text)
        help_button.configure(command=lambda: self._show_help(title, help_text, help_button))
        content = ttk.Frame(outer)
        content.grid(row=1, column=0, sticky="nsew", padx=4, pady=(1, 4))
        content.columnconfigure(0, weight=1)
        outer.rowconfigure(1, weight=1)
        content._section_outer = outer
        self._sections[title] = (content, toggle)
        toggle.configure(command=lambda key=title: self._toggle_section(key))
        if collapsed: content.grid_remove(); outer.rowconfigure(1, weight=0)
        return content

    def _toggle_section(self, title):
        """Collapse a block to its title, leaving neighbouring content available."""
        content, button = self._sections[title]
        collapsed = content.winfo_manager() != ""
        if collapsed:
            content.grid_remove(); content._section_outer.rowconfigure(1, weight=0)
        else:
            content.grid(); content._section_outer.rowconfigure(1, weight=1)
        button.configure(text="▶" if collapsed else "▼")
        self.ai_layout.setdefault("collapsed", {})[title] = collapsed
        outer = content._section_outer
        if isinstance(outer.master, tk.PanedWindow):
            minimum = 32 if collapsed else getattr(outer, "_expanded_minsize", 80)
            try: outer.master.paneconfigure(outer, minsize=minimum)
            except tk.TclError: pass

    @staticmethod
    def _configure_section_pane(content, minimum):
        """Remember an expanded minimum while allowing a collapsed title-only pane."""
        outer = content._section_outer
        outer._expanded_minsize = minimum
        if content.winfo_manager() == "":
            try: outer.master.paneconfigure(outer, minsize=32)
            except tk.TclError: pass

    def _show_help(self, title, text, anchor):
        messagebox.showinfo(title, text, parent=self.window)

    @staticmethod
    def _entry(parent, column, label, value, width=20, row=0, suffix=""):
        ttk.Label(parent, text=label).grid(row=row, column=column); entry = ttk.Entry(parent, width=width); entry.insert(0, str(value)); entry.grid(row=row, column=column + 1)
        if suffix: ttk.Label(parent, text=suffix).grid(row=row, column=column + 2, sticky="w")
        return entry
    @staticmethod
    def _tooltip(widget, text):
        tip = {"window": None}
        def show(_event):
            win = tk.Toplevel(widget); win.wm_overrideredirect(True); win.geometry("+{}+{}".format(widget.winfo_rootx() + 12, widget.winfo_rooty() + widget.winfo_height() + 4)); ttk.Label(win, text=text, padding=5).pack(); tip["window"] = win
        def hide(_event):
            if tip["window"]: tip["window"].destroy(); tip["window"] = None
        widget.bind("<Enter>", show, add="+"); widget.bind("<Leave>", hide, add="+")
    @staticmethod
    def _resolution_text(w, h): return "{} × {}".format(w, h) if w and h else "自動"
    def _schedule(self, delay, callback):
        if self._closed: return
        holder = {}
        def run(): self._after_ids.discard(holder.get("id")); callback()
        holder["id"] = self.window.after(delay, run); self._after_ids.add(holder["id"])
    def _worker(self, operation, work, success=None, failure=None):
        if self._busy: self._append("系統", "已有工作進行中，請稍候"); return
        self._busy = True; self._generation += 1; generation = self._generation; self._append("系統", operation + "中……")
        def run():
            try: result = True, work()
            except Exception as exc: result = False, str(exc)
            def done():
                if self._closed or generation != self._generation: return
                self._busy = False
                if result[0]: success(result[1]) if success else self._append("系統", operation + "完成")
                elif failure: failure(result[1])
                else: self._append("錯誤", result[1])
            self._schedule(0, done)
        threading.Thread(target=run, daemon=True, name="ai-{}".format(operation)).start()

    def scan_cameras(self):
        def scan():
            if not self.camera.stop(): raise RuntimeError(self.camera.error)
            return CameraCapture.discover(10, backend=self.config["ai"].get("camera_backend"))
        self._worker("掃描相機", scan, self._scanned)
    def _scanned(self, devices):
        self.devices = devices; labels = [d["display_name"] + "（index {}）".format(d["index"]) for d in devices]; self.camera_box["values"] = labels
        ai = self.config["ai"]; selected = CameraCapture.select_device(devices, ai.get("camera_device_id", ""), ai.get("camera_name", ""), ai.get("camera_index"))
        if selected:
            self.device_choice.set(labels[devices.index(selected)]); self.apply_camera()
        else:
            self.camera_status.config(text="狀態：找不到先前選擇的 USB 相機，請重新選擇（不會自動切換至虛擬相機）")
    def _current_device(self):
        try: return self.devices[self.camera_box.current()]
        except (IndexError, TypeError): return None
    def _device_changed(self, _event=None):
        device = self._current_device()
        if device: self.camera_status.config(text="已選擇 {}；按「套用相機設定」後才會開啟".format(device["name"]))
    def probe_capabilities(self):
        device = self._current_device()
        if not device: return
        self._probe_cancel.clear()
        def probe():
            if not self.camera.stop(): raise RuntimeError(self.camera.error)
            if device.get("backend") == "dshow":
                try:
                    modes = CameraCapture.enumerate_dshow_capabilities(device["name"], self.config["ai"].get("ffmpeg_path", ""))
                    if modes: return modes
                except FileNotFoundError: pass
            modes = CameraCapture.probe_capabilities(device, cancelled=self._probe_cancel.is_set)
            for item in modes:
                item.update(capability_source="部分探測", validation_status="OpenCV 已驗證",
                            backend=device.get("backend", "auto"), fourcc="", pixel_format="")
            return modes
        self._worker("完整偵測裝置規格", probe, self._capabilities)
    def _capabilities(self, modes):
        self.capabilities = modes
        resolutions = list(dict.fromkeys(["自動", "640 × 480", "1280 × 720", "1920 × 1080"] + [self._resolution_text(x["width"], x["height"]) for x in modes]))
        fps = list(dict.fromkeys(["自動", "15", "30", "60"] + [str(int(x["fps"])) if float(x["fps"]).is_integer() else str(x["fps"]) for x in modes]))
        self.res_box["values"], self.fps_box["values"] = resolutions, fps
        if self.resolution.get() not in resolutions: self.resolution.set("自動")
        if self.fps.get() not in fps: self.fps.set("自動")
        complete = bool(modes and all(x.get("capability_source") == "ffmpeg_dshow" for x in modes))
        self._append("系統", ("DirectShow 已回報全部 {} 種規格" if complete else
            "部分探測共 {} 種；設定 FFmpeg 路徑後可讀取 DirectShow 回報的完整規格").format(len(modes)))
        # Selection/probing is background work; only the quick threaded reader
        # is started here, so opening the dialog never enables AI Monitor.
        self.apply_camera()
    def apply_camera(self, save=False):
        device = self._current_device()
        if not device: messagebox.showwarning("相機", "請先選擇相機裝置", parent=self.window); return
        parts = self.resolution.get().replace(" ", "").split("×"); w, h = (map(int, parts) if len(parts) == 2 else (None, None))
        fps = None if self.fps.get() == "自動" else float(self.fps.get()); fourcc = None if self.fourcc.get() == "自動" else self.fourcc.get()
        self.camera.device_name = device["name"]
        if not self.camera.configure(device_id=device["device_id"], index=device["index"], backend=device["backend"], width=w, height=h, fps=fps, fourcc=fourcc):
            messagebox.showwarning("相機", self.camera.error, parent=self.window); return
        if save: self.save_settings()

    def _poll_preview(self):
        d = self._current_device() or {}; actual = self.camera.actual
        requested = self._resolution_text(self.camera.width, self.camera.height) + " @ " + (str(self.camera.fps) if self.camera.fps else "自動") + " FPS"
        actual_fps = "{:.1f}".format(actual["fps"]) if actual.get("fps") else "驅動未回報"
        actual_text = self._resolution_text(actual.get("width"), actual.get("height")) + " @ " + actual_fps + " FPS / " + (actual.get("fourcc") or "驅動未回報")
        requested += " / " + (self.fourcc.get() if self.fourcc.get() != "自動" else "自動")
        used = self.camera.fallback
        used_text = self._resolution_text(used.get("width"), used.get("height")) + " @ " + (str(used.get("fps")) if used.get("fps") else "自動") + " FPS / " + (used.get("fourcc") or "自動")
        status = "畫面讀取正常" if self.camera.state == "connected" else (self.camera.error or self.camera.state)
        self.camera_status.config(text="裝置：{}\n類型：{}相機　連線方式：{}　要求規格：{}\nfallback 後使用值：{}　實際規格：{}\n狀態：{}".format(d.get("name", "未選擇"), TYPE_LABELS.get(d.get("device_type"), "未知"), BACKEND_LABELS.get(used.get("backend", self.camera.backend), used.get("backend", self.camera.backend)), requested, used_text, actual_text, status))
        # A visible preview is always live.  The old "manual" branch had no
        # refresh action and therefore left the opening snapshot on screen.
        if getattr(self, "camera_view_state", "docked") != "hidden": self.show_latest()
        if self.monitor:
            while not self.monitor.results.empty():
                _, kind, value = self.monitor.results.get_nowait()
                self._record_history(value)
                ended_at = value.get("ended_at") if isinstance(value, dict) else None
                if kind == "answer": self._append("AI", value["text"], ended_at)
                elif kind == "timeout":
                    waited = value.get("timeout_sec", self.timeout.get()); self._append("錯誤", "AI 回答逾時：已等待 {} 秒".format(_number_text(waited)), ended_at); self.monitor_status.config(text="AI Monitor：AI 回答逾時（繼續運行）")
                else: self._append("錯誤", value.get("error", "未知錯誤"), ended_at); self.monitor_status.config(text="AI Monitor：發生錯誤（繼續運行）")
        self._schedule(150, self._poll_preview)
    def show_latest(self):
        frame = self.camera.latest_frame()
        if frame is None: return
        self.displayed_frame = frame.copy()
        try:
            from PIL import Image, ImageTk
            import cv2
            image = Image.fromarray(cv2.cvtColor(frame, cv2.COLOR_BGR2RGB))
            if self.camera_view_state == "docked":
                image.thumbnail((560, 250)); photo = ImageTk.PhotoImage(image)
                self.preview_label.config(image=photo, text=""); self.preview_label.image = photo
            elif self.camera_view_state == "detached" and self.detached_preview:
                self._render_detached_preview(image)
        except Exception as exc: self.preview_label.config(text="預覽失敗：{}".format(exc), image="")

    def _detached_fit_scale(self, image_width, image_height):
        return min(max(1, self.detached_preview.winfo_width()) / image_width,
                   max(1, self.detached_preview.winfo_height()) / image_height)

    def _render_detached_preview(self, image):
        """Render a live frame using the detached preview's zoom transform."""
        from PIL import Image, ImageTk
        canvas_width = max(1, self.detached_preview.winfo_width())
        canvas_height = max(1, self.detached_preview.winfo_height())
        fit_scale = self._detached_fit_scale(image.width, image.height)
        if self._detached_scale is None or self._detached_scale < fit_scale:
            self._detached_scale = fit_scale
            self._detached_offset = ((canvas_width - image.width * fit_scale) / 2,
                                     (canvas_height - image.height * fit_scale) / 2)
        size = (max(1, round(image.width * self._detached_scale)),
                max(1, round(image.height * self._detached_scale)))
        resized = image.resize(size, Image.Resampling.LANCZOS)
        self._detached_photo = ImageTk.PhotoImage(resized)
        self.detached_preview.delete("frame")
        self.detached_preview.create_image(*self._detached_offset, anchor="nw",
                                           image=self._detached_photo, tags="frame")
        if hasattr(self, "detached_status"):
            self.detached_status.config(text="滾輪或 ＋／－ 縮放｜按住左鍵拖曳｜{:.0f}%".format(self._detached_scale * 100))

    def _zoom_detached(self, factor, pointer=None):
        if self._detached_scale is None or self.displayed_frame is None: return "break"
        height, width = self.displayed_frame.shape[:2]
        fit_scale = self._detached_fit_scale(width, height)
        if pointer is None:
            pointer = (self.detached_preview.winfo_width() / 2,
                       self.detached_preview.winfo_height() / 2)
        self._detached_scale, self._detached_offset = self._zoom_at(
            self._detached_scale, self._detached_offset, pointer, factor, fit_scale)
        return "break"

    def _on_detached_wheel(self, event):
        direction = event.delta if getattr(event, "delta", 0) else (1 if event.num == 4 else -1)
        return self._zoom_detached(1.15 if direction > 0 else 1 / 1.15, (event.x, event.y))

    def _start_detached_drag(self, event): self._detached_drag = (event.x, event.y)
    def _drag_detached(self, event):
        if self._detached_drag is None or self._detached_offset is None: return
        self._detached_offset = (self._detached_offset[0] + event.x - self._detached_drag[0],
                                 self._detached_offset[1] + event.y - self._detached_drag[1])
        self._detached_drag = (event.x, event.y)

    def set_camera_view(self, state):
        """Switch only the live picture; camera controls always remain docked."""
        if state not in ("docked", "hidden", "detached"): return
        if self.detached_window:
            self.ai_layout["detached_geometry"] = self.detached_window.geometry()
            self.detached_window.destroy(); self.detached_window = self.detached_preview = None
            self._detached_photo = self._detached_scale = self._detached_offset = self._detached_drag = None
        self.camera_view_state = state; self.ai_layout["camera_state"] = state
        if state == "docked":
            self.preview_label.grid()
        else:
            self.preview_label.grid_remove()
        if state == "detached":
            win = self.detached_window = tk.Toplevel(self.window); win.title("相機預覽")
            win.geometry(self.ai_layout.get("detached_geometry", "800x600")); win.minsize(400, 300)
            self.detached_status = ttk.Label(win, text="滾輪或 ＋／－ 縮放｜按住左鍵拖曳")
            self.detached_status.pack(fill="x", padx=8, pady=4)
            self.detached_preview = tk.Canvas(win, highlightthickness=0, background="#202020", cursor="fleur")
            self.detached_preview.pack(fill="both", expand=True)
            self.detached_preview.bind("<Double-Button-1>", self._on_camera_double_click)
            self.detached_preview.bind("<MouseWheel>", self._on_detached_wheel)
            self.detached_preview.bind("<Button-4>", self._on_detached_wheel)
            self.detached_preview.bind("<Button-5>", self._on_detached_wheel)
            self.detached_preview.bind("<ButtonPress-1>", self._start_detached_drag)
            self.detached_preview.bind("<B1-Motion>", self._drag_detached)
            actions = ttk.Frame(win); actions.pack(pady=5)
            ttk.Button(actions, text="－", width=4, command=lambda: self._zoom_detached(1 / 1.15)).pack(side="left")
            ttk.Button(actions, text="＋", width=4, command=lambda: self._zoom_detached(1.15)).pack(side="left", padx=4)
            ttk.Button(actions, text="重新 Dock", command=lambda: self.set_camera_view("docked")).pack(side="left")
            win.protocol("WM_DELETE_WINDOW", lambda: self.set_camera_view("docked"))
        self._save_ai_layout()

    def _restore_ai_layout(self):
        try:
            sash = self.ai_layout.get("main_sash")
            if sash is not None: self.main_paned.sash_place(0, int(sash), 0)
        except (tk.TclError, TypeError, ValueError): pass
        self._restore_sashes(self.left_paned, self.ai_layout.get("left_sashes", []))
        self._restore_sashes(self.right_paned, self.ai_layout.get("right_sashes", []))
        self.set_camera_view(self.camera_view_state)

    @staticmethod
    def _restore_sashes(paned, positions):
        for index, position in enumerate(positions):
            try: paned.sash_place(index, 0, int(position))
            except (tk.TclError, TypeError, ValueError): break

    @staticmethod
    def _sash_positions(paned):
        positions = []
        for index in range(max(0, len(paned.panes()) - 1)):
            try: positions.append(paned.sash_coord(index)[1])
            except (tk.TclError, AttributeError): break
        return positions

    def _save_ai_layout(self):
        if self._closed: return
        try: self.ai_layout["geometry"] = self.window.geometry()
        except tk.TclError: pass
        try: self.ai_layout["main_sash"] = self.main_paned.sash_coord(0)[0]
        except (tk.TclError, AttributeError): pass
        if hasattr(self, "left_paned"):
            self.ai_layout["left_sashes"] = self._sash_positions(self.left_paned)
        if hasattr(self, "right_paned"):
            self.ai_layout["right_sashes"] = self._sash_positions(self.right_paned)
        self.ai_layout.pop("main_order", None)
        self.ai_layout["camera_state"] = self.camera_view_state
        if self.detached_window:
            try: self.ai_layout["detached_geometry"] = self.detached_window.geometry()
            except tk.TclError: pass

    def _draft_changed(self, _event=None):
        if self._draft_after:
            try: self.window.after_cancel(self._draft_after)
            except Exception: pass
        self._draft_after = self.window.after(700, self._autosave_draft)
    def _autosave_draft(self): self._draft_after = None; self._save(announce=False)
    def _system_changed(self, _event=None):
        if self._system_after:
            try: self.window.after_cancel(self._system_after)
            except Exception: pass
        self._system_after = self.window.after(700, self._autosave_system)
    def _autosave_system(self): self._system_after = None; self._save(announce=False)
    def clear_draft(self): self.user_text.delete("1.0", "end"); self._save(announce=False)
    def clear_console(self):
        self.response.config(state="normal")
        self.response.delete("1.0", "end")
        self.response.config(state="disabled")
    def _save(self, announce=True):
        self._save_ai_layout()
        ai = self.config.setdefault("ai", {}); device = self._current_device()
        try: timeout, delay, num_predict = float(self.timeout.get()), float(self.after_answer_delay.get()), int(self.num_predict.get())
        except ValueError: raise ValueError("AI 回答逾時、回答完成後等待與最多輸出 token 必須是數字")
        if timeout <= 0 or delay < 0: raise ValueError("AI 回答逾時必須大於 0，回答完成後等待不可小於 0")
        if num_predict < 1024: raise ValueError("最多輸出 token 不可小於 1024（Qwen3-VL 可能先使用內部思考 token）")
        parts = self.resolution.get().replace(" ", "").split("×"); w, h = (map(int, parts) if len(parts) == 2 else (None, None))
        ai.update({"enabled": bool(self.monitor and self.monitor.enabled), "camera_device_id": device.get("device_id", "") if device else ai.get("camera_device_id", ""), "camera_name": device.get("name", "") if device else ai.get("camera_name", ""), "camera_index": device.get("index", self.camera.index) if device else self.camera.index, "camera_backend": device.get("backend", self.camera.backend) if device else self.camera.backend, "camera_width": w, "camera_height": h, "camera_fps": None if self.fps.get() == "自動" else float(self.fps.get()), "camera_fourcc": "" if self.fourcc.get() == "自動" else self.fourcc.get(), "preview_mode": self.mode.get(), "base_url": self.base_url.get().strip(), "model": self.model.get().strip(), "timeout": timeout, "keep_alive": self.keep_alive.get().strip() or "30m", "think": self.think.get(), "num_predict": num_predict, "system_prompt": self.system.get("1.0", "end-1c"), "user_prompt_draft": self.user_text.get("1.0", "end-1c"), "after_answer_delay": delay, "alarm_on_detected": self.alarm_on_detected.get(), "alarm_on_timeout": self.alarm_on_timeout.get(), "alarm_on_error": self.alarm_on_error.get(), "sound_mode": self._sound_mode_key(), "sound_path": self.sound_path.get().strip(), "prompt_profiles": [dict(x) for x in self.profiles]})
        ai.pop("interval", None); ai.pop("sound_enabled", None)
        self.on_save(self.config)
        if announce: self._append("系統", "設定已保存")
    def save_settings(self):
        try: self._save()
        except ValueError as exc: messagebox.showerror("無法保存", str(exc), parent=self.window)

    def save_ui_layout(self):
        """Persist only window, divider, collapse and preview layout settings."""
        self._save_ai_layout()
        self.on_save(self.config)
        self._append("系統", "UI 配置已保存")

    def _refresh_profiles(self):
        for row in self.profile_tree.get_children(): self.profile_tree.delete(row)
        for p in self.profiles: self.profile_tree.insert("", "end", iid=p["id"], values=("啟用" if p.get("enabled") else "停用", p.get("name", "未命名")))
    def _profile_index(self):
        selected = self.profile_tree.selection(); return next((i for i, p in enumerate(self.profiles) if selected and p["id"] == selected[0]), None)
    def add_profile(self): self._profile_editor()
    def edit_profile(self):
        i = self._profile_index()
        if i is not None: self._profile_editor(i)
    def copy_profile(self):
        i = self._profile_index()
        if i is not None:
            copy = dict(self.profiles[i]); copy.update(id=uuid.uuid4().hex, name=copy.get("name", "") + "（複製）"); self.profiles.insert(i + 1, copy); self._refresh_profiles(); self._save(announce=False)
    def delete_profile(self):
        i = self._profile_index()
        if i is not None: del self.profiles[i]; self._refresh_profiles(); self._save(announce=False)
    def move_profile(self, delta):
        i = self._profile_index()
        if i is None or not 0 <= i + delta < len(self.profiles): return
        self.profiles[i], self.profiles[i + delta] = self.profiles[i + delta], self.profiles[i]; self._refresh_profiles(); self.profile_tree.selection_set(self.profiles[i + delta]["id"]); self._save(announce=False)
    def toggle_profile(self):
        i = self._profile_index()
        if i is not None: self.profiles[i]["enabled"] = not self.profiles[i].get("enabled"); self._refresh_profiles(); self._save(announce=False)
    def _profile_editor(self, index=None):
        old = self.profiles[index] if index is not None else {"id": uuid.uuid4().hex, "name": "", "prompt": "", "enabled": True}
        win = tk.Toplevel(self.window); win.title("編輯提問配置"); win.transient(self.window); win.grab_set()
        ttk.Label(win, text="配置名稱").pack(anchor="w"); name = ttk.Entry(win, width=60); name.pack(fill="x"); name.insert(0, old.get("name", ""))
        ttk.Label(win, text="Prompt 內容").pack(anchor="w"); prompt = tk.Text(win, width=70, height=12); prompt.pack(fill="both", expand=True); prompt.insert("1.0", old.get("prompt", ""))
        enabled = tk.BooleanVar(value=old.get("enabled", True)); ttk.Checkbutton(win, text="啟用", variable=enabled).pack(anchor="w")
        def save():
            value = {"id": old["id"], "name": name.get().strip() or "未命名", "prompt": prompt.get("1.0", "end-1c"), "enabled": enabled.get()}
            if index is None: self.profiles.append(value)
            else: self.profiles[index] = value
            self._refresh_profiles(); self._save(announce=False); win.destroy()
        ttk.Button(win, text="保存", command=save).pack(side="left"); ttk.Button(win, text="取消", command=win.destroy).pack(side="left")

    def _client(self):
        ai = self.config.get("ai", {})
        options, _unknown = build_ollama_options(ai.get("ollama_option_mode", "model_default"),
            ai.get("ollama_enabled_options", {}), ai.get("ollama_custom_options", {}),
            int(self.num_predict.get()))
        return QwenClient(self.base_url.get(), self.model.get(), float(self.timeout.get()),
            keep_alive=self.keep_alive.get(), think=self.think.get(),
            num_predict=int(self.num_predict.get()), options=options,
            option_mode=ai.get("ollama_option_mode", "model_default"))

    def open_model_options(self):
        """Open the optional sampling editor without crowding the main grid."""
        import json
        ai = self.config.setdefault("ai", {}); win = tk.Toplevel(self.window)
        win.title("模型進階參數"); win.transient(self.window)
        mode = tk.StringVar(value=ai.get("ollama_option_mode", "model_default"))
        labels = (("使用模型原生設定", "model_default"),
                  ("Qwen3-VL 8B O／X 建議設定", "qwen_ox"), ("自訂設定", "custom"))
        for text, value in labels: ttk.Radiobutton(win, text=text, variable=mode, value=value).pack(anchor="w", padx=8)
        ttk.Label(win, text="進階 JSON 編輯（相同 key 以此處為準）").pack(anchor="w", padx=8)
        editor = tk.Text(win, width=65, height=10); editor.pack(fill="both", expand=True, padx=8)
        editor.insert("1.0", json.dumps(ai.get("ollama_custom_options", {}), ensure_ascii=False, indent=2))
        preview = tk.Text(win, width=65, height=10, state="disabled"); preview.pack(fill="both", expand=True, padx=8)
        status = ttk.Label(win, text=""); status.pack(anchor="w", padx=8)
        def refresh():
            try:
                options, unknown = build_ollama_options(mode.get(),
                    ai.get("ollama_enabled_options", {}), editor.get("1.0", "end-1c"),
                    int(self.num_predict.get()))
                preview.config(state="normal"); preview.delete("1.0", "end")
                preview.insert("1.0", json.dumps(options, ensure_ascii=False, indent=2)); preview.config(state="disabled")
                status.config(text=("未知參數警告：" + ", ".join(unknown)) if unknown else "JSON 驗證成功")
                return options
            except Exception as exc: status.config(text=str(exc)); return None
        def save():
            options = refresh()
            if options is None: return
            ai["ollama_option_mode"] = mode.get()
            ai["ollama_custom_options"] = json.loads(editor.get("1.0", "end-1c") or "{}")
            self.on_save(self.config); win.destroy()
        bar = ttk.Frame(win); bar.pack(fill="x", padx=8, pady=6)
        ttk.Button(bar, text="更新實際 options 預覽", command=refresh).pack(side="left")
        ttk.Button(bar, text="恢復模型原生設定", command=lambda: (mode.set("model_default"), editor.delete("1.0", "end"), editor.insert("1.0", "{}"), refresh())).pack(side="left")
        ttk.Button(bar, text="保存", command=save).pack(side="right"); refresh()

    def open_image_options(self):
        ai = self.config.setdefault("ai", {}); win = tk.Toplevel(self.window); win.title("AI 影像輸入設定")
        notebook = ttk.Notebook(win); notebook.pack(fill="both", expand=True, padx=8, pady=8)
        variables = {}
        defaults = {"camera_ai_image": {"enabled": False, "crop_enabled": False, "resize_enabled": False,
            "target_width": 1280, "target_height": 720, "resize_mode": "contain", "allow_upscale": False,
            "align_qwen_grid": False, "output_format": "jpeg", "jpeg_quality": 95, "crop": [0, 0, 1, 1]},
            "attachment_ai_image": {"enabled": False, "crop_enabled": False, "resize_enabled": False,
            "target_width": 1280, "target_height": 720, "resize_mode": "contain", "allow_upscale": False,
            "align_qwen_grid": False, "output_format": "png", "jpeg_quality": 95, "crop": [0, 0, 1, 1]}}
        for key, title in (("camera_ai_image", "相機自動擷取"), ("attachment_ai_image", "附件／圖片庫")):
            data = dict(defaults[key]); data.update(ai.get(key, {})); tab = ttk.Frame(notebook); notebook.add(tab, text=title)
            rowvars = variables[key] = {name: tk.BooleanVar(value=data[name]) for name in ("enabled", "crop_enabled", "resize_enabled", "allow_upscale", "align_qwen_grid")}
            for label, name in (("套用影像前處理", "enabled"), ("套用裁切", "crop_enabled"), ("調整傳給 AI 的尺寸", "resize_enabled"), ("允許放大", "allow_upscale"), ("對齊 Qwen3-VL 32-pixel 網格", "align_qwen_grid")):
                ttk.Checkbutton(tab, text=label, variable=rowvars[name]).pack(anchor="w")
            for name in ("target_width", "target_height", "resize_mode", "output_format", "jpeg_quality"):
                row = ttk.Frame(tab); row.pack(fill="x"); ttk.Label(row, text=name, width=18).pack(side="left")
                var = tk.StringVar(value=str(data[name])); rowvars[name] = var
                values = {"resize_mode": ("contain", "cover", "stretch"), "output_format": ("png", "jpeg"),
                          "jpeg_quality": (70, 80, 85, 88, 90, 92, 95, 100)}.get(name)
                (ttk.Combobox(row, textvariable=var, values=values, state="readonly") if values else ttk.Entry(row, textvariable=var)).pack(side="left")
        help_tab = ttk.Frame(notebook); notebook.add(help_tab, text="使用說明")
        tips = "【裁切】裁掉無關背景通常是最有效的加速方式；裁切太小可能移除判斷上下文。\n\n【Resize】縮小尺寸能直接減少 visual tokens。\n\n【PNG】無損，適合遊戲 UI、字幕與小字。\n\n【JPEG】遊戲文字建議 quality 92～95。\n\n【保持比例】除非確定模型不受影響，否則不建議直接拉伸。\n\n【Qwen 尺寸對齊】以補邊對齊，不直接拉伸內容。"
        ttk.Label(help_tab, text=tips, justify="left", wraplength=620).pack(anchor="w", padx=8, pady=8)
        def save():
            for key, fields in variables.items():
                current = dict(defaults[key]); current.update(ai.get(key, {}))
                for name, var in fields.items(): current[name] = var.get()
                current["target_width"] = int(current["target_width"]); current["target_height"] = int(current["target_height"]); current["jpeg_quality"] = int(current["jpeg_quality"])
                ai[key] = current
            self.on_save(self.config); win.destroy()
        ttk.Button(win, text="保存", command=save).pack(side="right", padx=8, pady=6)

    def crop_camera(self):
        if self.displayed_frame is None:
            messagebox.showwarning("裁切畫面", "相機尚無可用畫面", parent=self.window); return
        from PIL import Image
        import cv2
        image = Image.fromarray(cv2.cvtColor(self.displayed_frame, cv2.COLOR_BGR2RGB))
        self._open_crop_dialog(image, "camera_ai_image", "裁切相機畫面")

    def crop_attachment(self):
        if not self.selected_image: messagebox.showwarning("附件裁切", "請先選取附件", parent=self.window); return
        from PIL import Image
        with Image.open(self.selected_image) as opened: image = opened.convert("RGB")
        self._open_crop_dialog(image, "attachment_ai_image", "附件裁切")

    def _open_crop_dialog(self, image, settings_key, title):
        """Select a real ROI and persist it for the matching AI image source."""
        from PIL import ImageTk
        win = tk.Toplevel(self.window); win.title(title); win.geometry("900x650"); win.minsize(600, 450)
        ttk.Label(win, text="按住左鍵拖曳裁切範圍；儲存後會自動啟用此來源的影像前處理與裁切。",
                  anchor="w").pack(fill="x", padx=8, pady=6)
        canvas = tk.Canvas(win, background="#202020", cursor="crosshair", highlightthickness=0)
        canvas.pack(fill="both", expand=True, padx=8)
        state = {"start": None, "end": None, "scale": 1.0, "offset": (0.0, 0.0), "photo": None}

        def render(_event=None):
            width, height = max(1, canvas.winfo_width()), max(1, canvas.winfo_height())
            scale = min(width / image.width, height / image.height)
            size = (max(1, round(image.width * scale)), max(1, round(image.height * scale)))
            state["scale"], state["offset"] = scale, ((width - size[0]) / 2, (height - size[1]) / 2)
            state["photo"] = ImageTk.PhotoImage(image.resize(size))
            canvas.delete("all"); canvas.create_image(*state["offset"], anchor="nw", image=state["photo"])
            if state["start"] and state["end"]:
                canvas.create_rectangle(*state["start"], *state["end"], outline="#ff4040", width=3, tags="roi")

        def press(event): state["start"] = state["end"] = (event.x, event.y); render()
        def drag(event): state["end"] = (event.x, event.y); render()
        def save_crop():
            if not state["start"] or not state["end"] or state["start"] == state["end"]:
                messagebox.showwarning(title, "請先拖曳出非空白的裁切範圍", parent=win); return
            first = map_canvas_point_to_image(*state["start"], image.width, image.height,
                                              state["scale"], state["offset"], relative=True)
            second = map_canvas_point_to_image(*state["end"], image.width, image.height,
                                               state["scale"], state["offset"], relative=True)
            from ai_image_processing import normalize_crop_roi
            roi = normalize_crop_roi(first, second)
            settings = self.config.setdefault("ai", {}).setdefault(settings_key, {})
            settings.update({"enabled": True, "crop_enabled": True, "crop": list(roi)})
            self.on_save(self.config)
            if settings_key == "camera_ai_image" and self.monitor:
                self.monitor.update_image_settings(settings)
                self._append("系統", "相機裁切已更新；AI Monitor 下一張畫面立即套用新範圍")
            if settings_key == "attachment_ai_image":
                self.attachment_crop_roi = roi
                self._select_image(self.selected_image, "附件", reset_crop=False)
                self._update_attachment_crop_status()
            win.destroy()

        canvas.bind("<Configure>", render); canvas.bind("<ButtonPress-1>", press); canvas.bind("<B1-Motion>", drag)
        buttons = ttk.Frame(win); buttons.pack(fill="x", padx=8, pady=8)
        ttk.Button(buttons, text="取消", command=win.destroy).pack(side="right")
        ttk.Button(buttons, text="儲存裁切範圍", command=save_crop).pack(side="right", padx=6)
    def test_connection(self): self._worker("檢查 Ollama", self._client().test_connection, self._models_loaded)
    def preload_model(self): self._worker("預先載入模型", self._client().preload)
    def _models_loaded(self, names): self.model["values"] = names; self.model.set(choose_model(names, self.model.get())); self._append("系統", "Ollama 連線成功")
    def _send_shortcut(self, _event): self.send_test(); return "break"
    def send_test(self):
        self._save(announce=False)
        try: configured = combine_prompt_profiles(self.profiles)
        except ValueError: configured = ""
        user = self.user_text.get("1.0", "end-1c").strip(); prompt = "\n\n".join(x for x in (configured, user) if x)
        if not prompt: messagebox.showwarning("提問", "請輸入文字或啟用提問配置", parent=self.window); return
        started = time.time()
        image_input, image_meta = self.selected_image, None
        settings = dict(self.config.get("ai", {}).get("attachment_ai_image", {}))
        # Resize/format settings are persistent, but a manual crop is only valid
        # after the user explicitly crops the currently selected attachment.
        attachment_crop_roi = getattr(self, "attachment_crop_roi", None)
        settings["crop_enabled"] = attachment_crop_roi is not None
        if attachment_crop_roi is not None:
            settings["crop"] = list(attachment_crop_roi)
        if self.selected_image and settings.get("enabled"):
            try:
                image_input, image_meta = prepare_and_encode_ai_image(self.selected_image, settings, "attachment")
            except Exception as exc:
                messagebox.showerror("附件處理失敗", str(exc), parent=self.window); return
        filename, path = self._archive_manual_image(image_input, started)
        profile_tags = [str(p.get("id", "")) for p in self.profiles if p.get("enabled")]
        tags = ", ".join(str(p.get("name", "")) for p in self.profiles if p.get("enabled"))
        base_history = {"history_id": uuid.uuid4().hex, "captured_at": started,
                        "sent_at": started, "filename": filename, "path": path,
                        "tag": tags, "profile_tags": profile_tags, "mode": "manual",
                        "prompt": prompt, "image_metadata": image_meta or {}}

        def completed(answer):
            ended = time.time()
            self._append("AI", answer, ended)
            self._record_history(dict(base_history, ended_at=ended,
                                      elapsed_sec=ended - started, answer=answer,
                                      text=answer, error=""))

        def failed(error):
            ended = time.time()
            self._append("錯誤", error, ended)
            self._record_history(dict(base_history, ended_at=ended,
                                      elapsed_sec=ended - started, answer="",
                                      text="", error=error))

        self._append("使用者", prompt)
        if image_meta:
            crop_note = "裁切 {}×{}".format(*image_meta["cropped_size"]) if settings.get("crop_enabled") else "未裁切"
            resize_note = "{}，目標框 {}×{}，{}放大".format(
                settings.get("resize_mode", "contain"), settings.get("target_width", 1280),
                settings.get("target_height", 720), "允許" if settings.get("allow_upscale") else "禁止")
            self._append("系統", "附件影像處理：原圖 {}×{} → {} → 輸出 {}×{}（{}）；{}，{} bytes".format(
                *image_meta["original_size"], crop_note, *image_meta["output_size"],
                resize_note, image_meta["format"], image_meta["bytes"]))
        self._worker("等待 AI 回答", lambda: self._client().chat(
            prompt, image_input, self.system.get("1.0", "end-1c")), completed, failed)

    @staticmethod
    def _manual_filename(started, extension="jpg"):
        milliseconds = int(started * 1000) % 1000
        return "manual_{}_{:03d}.{}".format(
            time.strftime("%Y%m%d_%H%M%S", time.localtime(started)), milliseconds,
            str(extension).lower().lstrip(".") or "jpg")

    def _archive_manual_image(self, source_path, started):
        """Keep the exact image used by a manual test beside monitor captures."""
        if not source_path:
            return "", ""
        from PIL import Image
        history_dir = os.path.join(os.path.dirname(__file__), "saved_ai_images", "monitor")
        os.makedirs(history_dir, exist_ok=True)
        extension = "png" if isinstance(source_path, bytes) and source_path.startswith(b"\x89PNG\r\n\x1a\n") else "jpg"
        filename = self._manual_filename(started, extension)
        path = os.path.join(history_dir, filename)
        if isinstance(source_path, bytes):
            with open(path, "wb") as stream: stream.write(source_path)
        else:
            with Image.open(source_path) as image:
                image.convert("RGB").save(path, "JPEG")
        return filename, path
    def toggle_monitor(self):
        if self.monitor and self.monitor.enabled:
            self.monitor.stop(); self.monitor_button.config(text="啟用 AI Monitor"); self.monitor_status.config(text="AI Monitor：已停止"); self.on_state and self.on_state(False); return
        try:
            self._save(announce=False); self._configure_alarm()
            self.monitor = AIMonitor(self.camera, self._client(), self.profiles, float(self.after_answer_delay.get()), self.alarm,
                self.alarm_on_detected.get(), self.alarm_on_timeout.get(), self.alarm_on_error.get(),
                self.config["ai"].get("stop_on_timeout", False), os.path.join(os.path.dirname(__file__), "saved_ai_images", "monitor"),
                system_prompt=self.system.get("1.0", "end-1c"),
                image_settings=self.config["ai"].get("camera_ai_image", {})); self.monitor.start()
            self.monitor_button.config(text="停用 AI Monitor"); self.monitor_status.config(text="AI Monitor：運行中"); self.on_state and self.on_state(True)
        except Exception as exc: self.monitor_status.config(text="AI Monitor：發生錯誤"); messagebox.showerror("無法啟用 AI", str(exc), parent=self.window)

    def _sound_mode_key(self):
        return {"系統警報聲": "system_alarm", "系統提示音": "system_notice", "自訂聲音檔": "custom", "靜音": "mute"}.get(self.sound_mode.get(), "system_alarm")
    def _configure_alarm(self): self.alarm.sound_mode = self._sound_mode_key(); self.alarm.sound_path = self.sound_path.get().strip()
    def _sound_mode_changed(self, _event=None):
        custom = self._sound_mode_key() == "custom"; state = "normal" if custom else "disabled"
        if hasattr(self, "pick_sound_button"):
            self.pick_sound_button.config(state=state); self.sound_file_label.config(state=state)
    def pick_sound(self):
        path = filedialog.askopenfilename(parent=self.window, filetypes=[("聲音檔", "*.wav *.mp3 *.ogg"), ("所有檔案", "*.*")])
        if path: self.sound_path.set(path); self.sound_file_label.config(text=os.path.basename(path)); self.sound_file_label._full_path = path; self.sound_status.config(text=path)
    def test_sound(self):
        self._configure_alarm()
        if self.alarm.sound_mode == "custom" and not os.path.isfile(self.alarm.sound_path):
            self.sound_status.config(text="自訂聲音檔不存在；測試時將改用系統警報聲")
        self.alarm.stop(); self.alarm.play("timeout")
        self._schedule(100, lambda: self.sound_status.config(text=self.alarm.error or self.sound_status.cget("text")))
    def _record_history(self, value):
        if not isinstance(value, dict): return
        item = dict(value); item.setdefault("history_id", uuid.uuid4().hex)
        item.setdefault("mode", "auto"); item.setdefault("profile_tags", [])
        ai = self.config.setdefault("ai", {})
        history = ai.get("question_history")
        if not isinstance(history, list): history = ai["question_history"] = []
        history.append(item); del history[:-QUESTION_HISTORY_LIMIT]
        self.on_save(self.config); self._refresh_history()
    def _refresh_history(self):
        if not hasattr(self, "history_tree"): return
        for row in self.history_tree.get_children(): self.history_tree.delete(row)
        self._history_by_id = {}
        history = self.config.get("ai", {}).get("question_history", [])
        if not isinstance(history, list): history = []
        history = history[-QUESTION_HISTORY_LIMIT:]
        for item in history:
            if isinstance(item, dict) and not item.get("history_id"): item["history_id"] = uuid.uuid4().hex
        for item in reversed(history):
            if not isinstance(item, dict): continue
            history_id = str(item["history_id"]); self._history_by_id[history_id] = item
            stamp = time.strftime("%H:%M:%S", time.localtime(item.get("ended_at", item.get("captured_at", 0))))
            result = "逾時 {} 秒".format(_number_text(item.get("timeout_sec", item.get("elapsed_sec", 0)))) if item.get("error") == "AI 回答逾時" else (item.get("error") or item.get("answer", ""))
            summary = "{}｜{}｜{}｜{}".format(stamp, item.get("tag", ""), item.get("filename", ""), result)
            self.history_tree.insert("", "end", iid=history_id, values=(summary,))

    def _on_history_double_click(self, event):
        row_id = self.history_tree.identify_row(event.y)
        item = self._history_by_id.get(row_id)
        if not item: return
        path = item.get("path", "")
        if not os.path.isfile(path):
            messagebox.showwarning("歷史圖片", "找不到歷史圖片：\n{}".format(item.get("filename") or os.path.basename(path)), parent=self.window)
            return
        self._open_image_viewer("history", path=path, title="歷史圖片", metadata=item)

    def _append(self, role, value, timestamp=None):
        try: stamp = time.strftime("%H:%M:%S", time.localtime(float(timestamp) if timestamp is not None else time.time()))
        except (TypeError, ValueError, OverflowError, OSError): stamp = time.strftime("%H:%M:%S", time.localtime())
        self.response.config(state="normal"); self.response.insert("end", "[{}] {}：{}\n\n".format(stamp, role, value)); self.response.config(state="disabled"); self.response.see("end")
    def save_image(self):
        if self.displayed_frame is not None: self.library.save(self.displayed_frame); self.refresh_library()
    def pick_image(self):
        path = filedialog.askopenfilename(parent=self.window, filetypes=[("圖片", "*.jpg *.jpeg *.png")])
        if path: self._select_image(path, "檔案")
    def use_latest(self):
        if self.displayed_frame is not None:
            item = self.library.save(self.displayed_frame, source="camera"); self.refresh_library(); self._select_image(item["path"], "最新相機畫面")
    def clear_image(self):
        self.selected_image = None; self.attachment_crop_roi = None
        self.image_info.config(text="本次提問尚未附加圖片"); self.attachment_preview.config(image="", text="無附件")
        self._update_attachment_crop_status()
    def _update_attachment_crop_status(self):
        if not hasattr(self, "attachment_crop_status"): return
        if self.attachment_crop_roi is None:
            text = "裁切：未套用（手動附件每次重新選圖都會重設）"
        else:
            left, top, right, bottom = self.attachment_crop_roi
            text = "裁切：已套用到本次附件（範圍 {:.1%}～{:.1%} × {:.1%}～{:.1%}）".format(left, right, top, bottom)
        self.attachment_crop_status.config(text=text)
    def _select_image(self, path, source, reset_crop=True):
        try:
            from PIL import Image, ImageTk
            with Image.open(path) as im: size = "{}×{}".format(*im.size); thumb = im.convert("RGB"); thumb.thumbnail((400, 130)); photo = ImageTk.PhotoImage(thumb)
            self.selected_image = path
            if reset_crop: self.attachment_crop_roi = None
            self.image_info.config(text="已附加：{}（{}，{}）".format(os.path.basename(path), size, source)); self.attachment_preview.config(image=photo, text=""); self.attachment_preview.image = photo
            self._update_attachment_crop_status()
        except Exception as exc: messagebox.showerror("圖片無法開啟", str(exc), parent=self.window)
    def refresh_library(self):
        if not hasattr(self, "library_tree"): return
        for row in self.library_tree.get_children(): self.library_tree.delete(row)
        for item in reversed(self.library.list()): self.library_tree.insert("", "end", iid=item["id"], values=("★" if item.get("favorite") else "", time_text(item.get("saved_at")), item.get("source", "")))
    def select_library_image(self, _event=None):
        selected = self.library_tree.selection()
        if selected:
            item = next((x for x in self.library.list() if x["id"] == selected[0]), None)
            if item: self._selected_item = item; self._select_image(item["path"], "圖片庫")

    def _on_camera_double_click(self, _event=None):
        if self.displayed_frame is not None:
            self._open_image_viewer("camera", frame=self.displayed_frame.copy(), title="相機即時畫面")
    def _on_attachment_double_click(self, _event=None):
        if self.selected_image: self._open_image_viewer("attachment", path=self.selected_image, title="提問附件")
    def _on_library_double_click(self, event):
        row_id = self.library_tree.identify_row(event.y)
        item = next((x for x in self.library.list() if x.get("id") == row_id), None)
        if item: self._open_image_viewer("library", path=item.get("path"), title="圖片庫", metadata=item)

    def _open_image_viewer(self, source_type, path=None, frame=None, title="", metadata=None):
        """Open an explicit image source; never infer it from attachment state."""
        self._close_image_viewer()
        try:
            from PIL import Image
            if source_type == "camera":
                if frame is None: return
                import cv2
                image = Image.fromarray(cv2.cvtColor(frame, cv2.COLOR_BGR2RGB))
                filename = "相機即時畫面"
            else:
                if not path or not os.path.isfile(path): raise FileNotFoundError(path or "")
                with Image.open(path) as opened: image = opened.convert("RGB")
                filename = os.path.basename(path)
        except Exception as exc:
            messagebox.showerror("圖片無法開啟", str(exc), parent=self.window); return
        self._viewer_image = image
        viewer = self.zoom_window = tk.Toplevel(self.window); viewer.title(title or "圖片檢視")
        viewer.geometry("900x700"); viewer.minsize(800, 600)
        source_labels = {"camera": "相機", "attachment": "附件", "library": "圖片庫", "history": "最近提問"}
        ttk.Label(viewer, text="{}   原始解析度：{}×{}   來源：{}".format(filename, image.width, image.height, source_labels.get(source_type, source_type))).pack(fill="x", padx=8, pady=5)
        self.viewer_status = ttk.Label(viewer, text="滾輪縮放（以滑鼠位置為中心）｜按住左鍵拖曳圖片")
        self.viewer_status.pack(fill="x", padx=8)
        self.zoom_canvas = tk.Canvas(viewer, highlightthickness=0, background="#202020", cursor="fleur")
        self.zoom_canvas.pack(fill="both", expand=True)
        ttk.Button(viewer, text="關閉", command=self._close_image_viewer).pack(pady=5)
        viewer.bind("<Escape>", lambda _e: self._close_image_viewer())
        self.zoom_canvas.bind("<Configure>", self._queue_viewer_render)
        self.zoom_canvas.bind("<MouseWheel>", self._on_viewer_wheel)
        self.zoom_canvas.bind("<Button-4>", self._on_viewer_wheel)
        self.zoom_canvas.bind("<Button-5>", self._on_viewer_wheel)
        self.zoom_canvas.bind("<ButtonPress-1>", self._start_viewer_drag)
        self.zoom_canvas.bind("<B1-Motion>", self._drag_viewer)
        viewer.protocol("WM_DELETE_WINDOW", self._close_image_viewer)
        self._viewer_scale = None; self._viewer_offset = None; self._viewer_drag = None
        self._queue_viewer_render()

    def _queue_viewer_render(self, _event=None):
        if not self.zoom_window: return
        if self._viewer_after:
            try: self.zoom_window.after_cancel(self._viewer_after)
            except tk.TclError: pass
        self._viewer_after = self.zoom_window.after(100, self._render_viewer)
    def _render_viewer(self):
        self._viewer_after = None
        if not self.zoom_window or self._viewer_image is None: return
        from PIL import Image, ImageTk
        canvas_width = max(1, self.zoom_canvas.winfo_width())
        canvas_height = max(1, self.zoom_canvas.winfo_height())
        fit_scale = min(canvas_width / self._viewer_image.width,
                        canvas_height / self._viewer_image.height)
        if self._viewer_scale is None:
            self._viewer_scale = fit_scale
            self._viewer_offset = (
                (canvas_width - self._viewer_image.width * fit_scale) / 2,
                (canvas_height - self._viewer_image.height * fit_scale) / 2,
            )
        elif self._viewer_scale < fit_scale:
            self._viewer_scale = fit_scale
            self._viewer_offset = (
                (canvas_width - self._viewer_image.width * fit_scale) / 2,
                (canvas_height - self._viewer_image.height * fit_scale) / 2,
            )
        size = (max(1, round(self._viewer_image.width * self._viewer_scale)),
                max(1, round(self._viewer_image.height * self._viewer_scale)))
        image = self._viewer_image.resize(size, Image.Resampling.LANCZOS)
        self._viewer_photo = ImageTk.PhotoImage(image)
        self.zoom_canvas.delete("all")
        self.zoom_canvas.create_image(*self._viewer_offset, anchor="nw", image=self._viewer_photo)
        self.viewer_status.config(text="滾輪縮放（以滑鼠位置為中心）｜按住左鍵拖曳圖片｜{:.0f}%".format(self._viewer_scale * 100))

    @staticmethod
    def _zoom_at(scale, offset, pointer, factor, minimum, maximum=8.0):
        """Return a zoom transform that keeps the pixel below ``pointer`` fixed."""
        maximum = max(maximum, minimum)
        new_scale = min(max(scale * factor, minimum), maximum)
        ratio = new_scale / scale
        return new_scale, (pointer[0] - (pointer[0] - offset[0]) * ratio,
                           pointer[1] - (pointer[1] - offset[1]) * ratio)

    def _on_viewer_wheel(self, event):
        if self._viewer_scale is None or self._viewer_image is None: return "break"
        direction = event.delta if getattr(event, "delta", 0) else (1 if event.num == 4 else -1)
        factor = 1.15 if direction > 0 else 1 / 1.15
        fit_scale = min(self.zoom_canvas.winfo_width() / self._viewer_image.width,
                        self.zoom_canvas.winfo_height() / self._viewer_image.height)
        self._viewer_scale, self._viewer_offset = self._zoom_at(
            self._viewer_scale, self._viewer_offset, (event.x, event.y), factor, fit_scale)
        self._render_viewer()
        return "break"

    def _start_viewer_drag(self, event):
        self._viewer_drag = (event.x, event.y)

    def _drag_viewer(self, event):
        if self._viewer_drag is None or self._viewer_offset is None: return
        dx, dy = event.x - self._viewer_drag[0], event.y - self._viewer_drag[1]
        self._viewer_offset = (self._viewer_offset[0] + dx, self._viewer_offset[1] + dy)
        self._viewer_drag = (event.x, event.y)
        self._render_viewer()
    def _close_image_viewer(self):
        if self._viewer_after and self.zoom_window:
            try: self.zoom_window.after_cancel(self._viewer_after)
            except tk.TclError: pass
        self._viewer_after = None
        if self.zoom_window:
            try: self.zoom_window.destroy()
            except tk.TclError: pass
        if self._viewer_image is not None:
            try: self._viewer_image.close()
            except Exception: pass
        self.zoom_window = self._viewer_image = self._viewer_photo = None
        self._viewer_scale = self._viewer_offset = self._viewer_drag = None
    def close(self):
        if self._closed: return
        self._save_ai_layout(); self._close_image_viewer()
        try: self._save(announce=False)
        except Exception: pass
        self._closed = True; self._generation += 1
        self._probe_cancel.set()
        if self.monitor: self.monitor.stop()
        self.alarm.stop(); self.camera.stop()
        for after_id in list(self._after_ids):
            try: self.window.after_cancel(after_id)
            except Exception: pass
        self.config["ai"]["enabled"] = False; self.on_save(self.config); self.on_state and self.on_state(False); self.window.destroy()


def time_text(timestamp):
    try: return time.strftime("%Y-%m-%d %H:%M", time.localtime(float(timestamp)))
    except Exception: return ""


def _number_text(value):
    try:
        number = float(value)
        return str(int(number)) if number.is_integer() else "{:.1f}".format(number)
    except (TypeError, ValueError): return str(value)
