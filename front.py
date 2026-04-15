# -*- coding: utf-8 -*-
import json
import os
import re
import socket
import time
import threading 
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from pynput import keyboard

DEFAULT_PI_HOST = "192.168.100.140"
PI_PORT = 5000
BUFF_SKIP_MODE_WALK = "walk"
BUFF_SKIP_MODE_COMPRESS = "compress"

SAVE_DIR = "saved_timelines"
CONFIG_FILE = "front_config.json"

KEY_MAP = {
    "Key.up": "up",
    "Key.down": "down",
    "Key.left": "left",
    "Key.right": "right",
    "Key.space": "space",
    "'x'": "x",
    "'c'": "c",
    "'v'": "v",
    "'g'": "g",
    "'f'": "f",
    "'d'": "d",
    "'6'": "6",
    "Key.shift": "shift",
    "Key.shift_r": "shift",
    "Key.ctrl_l": "ctrl",
    "Key.ctrl_r": "ctrl",
    "Key.alt_l": "alt",
    "Key.alt_r": "alt",
    "Key.alt_gr": "alt",
    "Key.cmd": "fn",
    "Key.cmd_r": "fn",
}

events = []
recording = False
recording_start = None
pressed_keys = set()
listener = None


def ensure_dirs():
    if not os.path.exists(SAVE_DIR):
        os.makedirs(SAVE_DIR)


def load_config():
    if not os.path.exists(CONFIG_FILE):
        return {
            "pi_host": DEFAULT_PI_HOST,
            "last_selected_name": "",
            "buff_skip_mode": BUFF_SKIP_MODE_COMPRESS,
            "ui_layout": {
                "paned_sash_x": None,
                "window_size": None,
                "tree_column_widths": {}
            }
        }

    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        ui_layout = data.get("ui_layout", {})
        if not isinstance(ui_layout, dict):
            ui_layout = {}
        tree_column_widths = ui_layout.get("tree_column_widths", {})
        if not isinstance(tree_column_widths, dict):
            tree_column_widths = {}
        return {
            "pi_host": data.get("pi_host", DEFAULT_PI_HOST),
            "last_selected_name": data.get("last_selected_name", ""),
            "buff_skip_mode": data.get("buff_skip_mode", BUFF_SKIP_MODE_COMPRESS),
            "ui_layout": {
                "paned_sash_x": ui_layout.get("paned_sash_x"),
                "paned_ratio": ui_layout.get("paned_ratio"),
                "window_size": ui_layout.get("window_size"),
                "tree_column_widths": tree_column_widths
            }
        }
    except Exception:
        return {
            "pi_host": DEFAULT_PI_HOST,
            "last_selected_name": "",
            "buff_skip_mode": BUFF_SKIP_MODE_COMPRESS,
            "ui_layout": {
                "paned_sash_x": None,
                "window_size": None,
                "tree_column_widths": {}
            }
        }


def save_config(config):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)


def normalize_key(key):
    try:
        return str(key)
    except Exception:
        return None


def sanitize_filename(name):
    name = name.strip()
    name = re.sub(r'[\\/:*?"<>|]', "_", name)
    return name


def timeline_file_path(name):
    safe_name = sanitize_filename(name)
    return os.path.join(SAVE_DIR, safe_name + ".json")


def list_saved_timeline_names():
    ensure_dirs()
    names = []
    for fn in os.listdir(SAVE_DIR):
        if fn.lower().endswith(".json"):
            names.append(os.path.splitext(fn)[0])
    names.sort()
    return names


def save_named_timeline(name, timeline, meta=None):
    path = timeline_file_path(name)
    data = {
        "name": name,
        "saved_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        "events": timeline,
        "_meta": meta or {}
    }
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def load_named_timeline(name):
    path = timeline_file_path(name)
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def delete_named_timeline(name):
    path = timeline_file_path(name)
    if os.path.exists(path):
        os.remove(path)


def on_press(key):
    global events
    if not recording:
        return

    key_name = normalize_key(key)
    mapped = KEY_MAP.get(key_name)
    if not mapped:
        return

    if key_name in pressed_keys:
        return

    pressed_keys.add(key_name)
    now = time.time()
    events.append({
        "raw_key": key_name,
        "button": mapped,
        "type": "press",
        "time": now
    })


def on_release(key):
    global events
    if not recording:
        return

    key_name = normalize_key(key)
    mapped = KEY_MAP.get(key_name)
    if not mapped:
        return

    if key_name in pressed_keys:
        pressed_keys.remove(key_name)

    now = time.time()
    events.append({
        "raw_key": key_name,
        "button": mapped,
        "type": "release",
        "time": now
    })


def send_to_pi(payload, pi_host, timeout=8):
    raw_payload = json.dumps(payload, ensure_ascii=False).encode("utf-8")

    with socket.create_connection((pi_host, PI_PORT), timeout=timeout) as s:
        s.settimeout(timeout)
        s.sendall(raw_payload)
        try:
            s.shutdown(socket.SHUT_WR)
        except OSError:
            pass

        chunks = []
        while True:
            part = s.recv(1024 * 1024)
            if not part:
                break
            chunks.append(part)

    if not chunks:
        raise ConnectionError("Pi 沒有回應資料")

    data = b"".join(chunks)
    return json.loads(data.decode("utf-8"))


def build_timeline(raw_events, start_time):
    if not raw_events:
        return []

    timeline = []
    for ev in raw_events:
        timeline.append({
            "type": ev["type"],
            "button": ev["button"],
            "at": round(max(0.0, ev["time"] - start_time), 4),
            "at_jitter": 0.0,
            "buff_group": "",
            "buff_cycle_sec": 0.0,
            "buff_jitter_sec": 0.0
        })

    timeline.sort(key=lambda x: x["at"])
    return timeline


def build_overlap_summary(timeline):
    if not timeline:
        return []

    groups = []
    current = [timeline[0]]
    threshold = 0.05

    for ev in timeline[1:]:
        if abs(ev["at"] - current[-1]["at"]) <= threshold:
            current.append(ev)
        else:
            groups.append(current)
            current = [ev]
    groups.append(current)

    summary = []
    for i, grp in enumerate(groups):
        summary.append({
            "group": i,
            "at_range": [grp[0]["at"], grp[-1]["at"]],
            "events": grp
        })
    return summary


class App:
    def normalize_event_schema(self, ev):
        row = dict(ev)
        row.setdefault("type", "")
        row.setdefault("button", "")
        row["at"] = round(max(0.0, float(row.get("at", 0.0))), 4)
        row["at_jitter"] = round(abs(float(row.get("at_jitter", 0.0))), 4)
        row["buff_group"] = str(row.get("buff_group", "")).strip()
        row["buff_cycle_sec"] = round(max(0.0, float(row.get("buff_cycle_sec", 0.0))), 4)
        row["buff_jitter_sec"] = round(abs(float(row.get("buff_jitter_sec", 0.0))), 4)
        return row

    def new_meta(self):
        return {
            "original_events": [],
            "history": []
        }

    def normalize_meta(self, meta):
        if not isinstance(meta, dict):
            return self.new_meta()
        raw_original = meta.get("original_events", [])
        raw_history = meta.get("history", [])
        original_events = []
        if isinstance(raw_original, list):
            original_events = [self.normalize_event_schema(ev) for ev in raw_original]

        history = []
        if isinstance(raw_history, list):
            for item in raw_history:
                if not isinstance(item, dict):
                    continue
                events = item.get("events", [])
                if not isinstance(events, list):
                    continue
                history.append({
                    "ts": str(item.get("ts", "")),
                    "reason": str(item.get("reason", "")),
                    "events": [self.normalize_event_schema(ev) for ev in events]
                })
        return {
            "original_events": original_events,
            "history": history[-5:]
        }

    def copy_events(self, rows):
        return [self.normalize_event_schema(ev) for ev in rows]

    def push_history(self, reason):
        snap = {
            "ts": time.strftime("%Y-%m-%d %H:%M:%S"),
            "reason": reason,
            "events": self.copy_events(self.timeline)
        }
        history = self.timeline_meta.setdefault("history", [])
        history.append(snap)
        if len(history) > 5:
            del history[:-5]

    def restore_original_timeline(self):
        original = self.timeline_meta.get("original_events", [])
        if not original:
            return False
        self.timeline = self.copy_events(original)
        self.current_loaded_from_saved = False
        self.update_current_labels()
        self.refresh_tree()
        self.refresh_preview()
        return True

    def persist_current_timeline_if_named(self):
        if not self.current_name:
            return
        save_named_timeline(self.current_name, self.timeline, self.timeline_meta)

    def recalculate_timeline_for_runtime(self):
        events = self.copy_events(self.timeline)
        if not events:
            return events

        seen_negative_group = {}
        i = 0
        offset = 0.0
        while i < len(events):
            ev = events[i]
            group_name = str(ev.get("buff_group", "")).strip()
            is_negative_group = group_name.startswith("-")
            if not is_negative_group:
                events[i]["at"] = round(events[i]["at"] + offset, 4)
                i += 1
                continue

            if group_name in seen_negative_group:
                raise ValueError("負群 {} 必須成段，不可分散".format(group_name))
            seen_negative_group[group_name] = True

            start = i
            end = i
            while end < len(events):
                g = str(events[end].get("buff_group", "")).strip()
                if g != group_name:
                    break
                end += 1

            segment = events[start:end]
            base_at = min(item["at"] for item in segment)
            for row in segment:
                row["at"] = round(max(0.0, row["at"] - base_at), 4)

            duration = max(row["at"] for row in segment) if segment else 0.0
            offset = round(offset + duration, 4)
            i = end

        return events

    def validate_negative_group_monotonic_by_index(self, events):
        if not events:
            return

        i = 0
        while i < len(events):
            group_name = str(events[i].get("buff_group", "")).strip()
            if not group_name.startswith("-"):
                i += 1
                continue

            start = i
            end = i
            while end < len(events):
                g = str(events[end].get("buff_group", "")).strip()
                if g != group_name:
                    break
                end += 1

            prev_at = float(events[start].get("at", 0.0))
            for cur_idx in range(start + 1, end):
                cur_at = float(events[cur_idx].get("at", 0.0))
                if cur_at < prev_at:
                    raise ValueError(
                        "負群 {} 的 at 時間必須依 idx 非遞減：idx {} 的 at={}，但 idx {} 的 at={}。"
                        "請先調整同群組列順序或時間，避免後列比前列更早。".format(
                            group_name,
                            cur_idx - 1,
                            round(prev_at, 4),
                            cur_idx,
                            round(cur_at, 4)
                        )
                    )
                prev_at = cur_at

            i = end

    def update_error_text(self, widget, message):
        widget.config(state="normal")
        widget.delete("1.0", tk.END)
        widget.insert(tk.END, message or "無")
        widget.config(state="disabled")

    def set_frontend_error(self, message):
        self.update_error_text(self.frontend_error_text, message)

    def set_backend_error(self, message):
        self.update_error_text(self.backend_error_text, message)

    def clear_errors(self):
        self.set_frontend_error("")
        self.set_backend_error("")

    def set_status(self, message):
        self.status_var.set(message)

    def set_connection_status(self, message):
        self.connection_var.set(message)

    def set_connected(self, connected, message=""):
        self.connected = connected
        if connected:
            self.offline_mode = False
            self.set_connection_status("已連線（{}）".format(self.config["pi_host"]))
            if message:
                self.set_status(message)
        else:
            self.set_connection_status("未連線")
            if message:
                self.set_status(message)

    def mark_timeline_dirty(self):
        self.current_loaded_from_saved = False
        self.update_current_labels()
        self.refresh_tree()
        self.refresh_preview()

    def request_pi(self, payload, success_status=None, write_response=True):
        if self.offline_mode and payload.get("action") != "ping":
            raise ConnectionError("目前為離線模式，請先按「測試連線」")

        last_err = None
        with self.request_lock:
            for retry in range(2):
                try:
                    self.ensure_connection()
                    wire = json.dumps(payload, ensure_ascii=False) + "\n"
                    self.conn.sendall(wire.encode("utf-8"))
                    line = self.conn_file.readline()
                    if not line:
                        raise ConnectionError("連線已中斷，未收到回應")
                    res = json.loads(line)
                    break
                except Exception as e:
                    last_err = e
                    self.close_connection()
                    if retry == 0:
                        time.sleep(0.15)
                        continue
                    err = str(last_err)
                    self.set_frontend_error(err)
                    self.set_connected(False)
                    if success_status:
                        self.set_status(success_status + "（前端連線失敗）")
                    raise

        self.set_connected(True)
        status = res.get("status")
        if status in ("error", "busy"):
            self.set_backend_error(res.get("message", "Pi 回傳錯誤"))
        else:
            self.set_backend_error("")

        if write_response:
            self.write_text({
                "current_name": self.current_name,
                "pi_host": self.config["pi_host"],
                "request": payload,
                "response": res
            })
        return res

    def ensure_connection(self):
        if self.conn is not None and self.conn_file is not None and self.connected:
            return
        self.open_connection()

    def open_connection(self, timeout=3.0):
        self.close_connection(silent=True)
        host = self.config["pi_host"]
        sock = socket.create_connection((host, PI_PORT), timeout=timeout)
        sock.settimeout(timeout)
        self.conn = sock
        self.conn_file = sock.makefile("r", encoding="utf-8")
        self.set_connected(True)

    def close_connection(self, silent=False):
        if self.conn_file is not None:
            try:
                self.conn_file.close()
            except Exception:
                pass
        if self.conn is not None:
            try:
                self.conn.close()
            except Exception:
                pass
        self.conn_file = None
        self.conn = None
        if not silent:
            self.set_connected(False)

    def __init__(self, root):
        ensure_dirs()
        self.config = load_config()

        self.root = root
        self.root.title("鍵盤錄製 / Timeline 分析 / 傳送")
        self.timeline = []
        self.timeline_meta = self.new_meta()
        self.current_name = ""
        self.current_loaded_from_saved = False
        self.connected = False
        self.offline_mode = False
        self.request_lock = threading.Lock()
        self.drag_iid = None
        self.conn = None
        self.conn_file = None

        container = tk.Frame(root)
        container.pack(fill="both", expand=True, padx=10, pady=8)

        body = tk.PanedWindow(container, orient=tk.HORIZONTAL, sashrelief=tk.RAISED)
        body.pack(fill="both", expand=True)
        self.body = body

        left_panel = tk.Frame(body)
        right_panel = tk.Frame(body)
        body.add(left_panel, minsize=440)
        body.add(right_panel, minsize=320)

        top = tk.LabelFrame(left_panel, text="操作區")
        top.pack(fill="x", pady=(0, 6))

        btn_specs = [
            ("重新分析", self.analyze, "#fff4b3"),
            ("開始錄製", self.toggle_record, "#9be58b"),
            ("送出執行", self.send_timeline, "#c8f7c5"),
            ("重複執行", self.send_timeline_loop, "#b7f0ad"),
            ("停止", self.stop_pi, "#ff8c69"),
        ]
        btn_count = len(btn_specs)

        self.status_var = tk.StringVar(value="尚未錄製")
        tk.Label(top, textvariable=self.status_var, anchor="w").grid(
            row=0, column=0, columnspan=btn_count, sticky="we", padx=8, pady=(6, 2)
        )
        self.record_button = None
        for idx, (txt, cmd, color) in enumerate(btn_specs):
            kwargs = {"text": txt, "command": cmd, "width": 10}
            if color:
                kwargs["bg"] = color
            btn = tk.Button(top, **kwargs)
            btn.grid(row=1, column=idx, padx=4, pady=(2, 8))
            if idx == 1:
                self.record_button = btn

        self.current_script_var = tk.StringVar(value="【目前腳本：未命名 / 未儲存】")
        tk.Label(
            top,
            textvariable=self.current_script_var,
            anchor="w",
            fg="#1a4fb8"
        ).grid(row=2, column=0, columnspan=btn_count, sticky="w", padx=8, pady=(0, 8))

        info = tk.LabelFrame(left_panel, text="目前套用資訊")
        row = tk.Frame(info)
        row.pack(fill="x", padx=8, pady=2)

        tk.Label(row, text="Pi IP：").pack(side="left")
        self.pi_ip_entry = tk.Entry(row, width=18)
        self.pi_ip_entry.insert(0, self.config["pi_host"])
        self.pi_ip_entry.pack(side="left", padx=5)
        tk.Button(row, text="保存", command=self.save_pi_ip).pack(side="left", padx=5)

        connection_row = tk.Frame(info)
        connection_row.pack(fill="x", padx=8, pady=(2, 5))
        tk.Label(connection_row, text="連線狀況：").pack(side="left")
        self.connection_var = tk.StringVar(value="未連線")
        tk.Label(connection_row, textvariable=self.connection_var, fg="#1a4fb8").pack(side="left", padx=(0, 10))
        tk.Button(connection_row, text="測試連線", command=self.ping_pi, width=10).pack(side="left", padx=4)
        tk.Button(connection_row, text="我要離線", command=self.go_offline, width=10).pack(side="left", padx=4)

        buff_mode_row = tk.Frame(info)
        buff_mode_row.pack(fill="x", padx=8, pady=(0, 5))
        tk.Label(buff_mode_row, text="buff 略過模式：").pack(side="left")
        self.buff_skip_mode_var = tk.StringVar(value=self.config.get("buff_skip_mode", BUFF_SKIP_MODE_COMPRESS))
        self.buff_skip_mode_combo = ttk.Combobox(
            buff_mode_row,
            state="readonly",
            width=32,
            values=[
                "略過(不等不按，壓縮時間軸)",
                "走過(等但不按，保留時間軸)"
            ]
        )
        if self.buff_skip_mode_var.get() == BUFF_SKIP_MODE_WALK:
            self.buff_skip_mode_combo.current(1)
        else:
            self.buff_skip_mode_combo.current(0)
        self.buff_skip_mode_combo.pack(side="left", padx=5)
        self.buff_skip_mode_combo.bind("<<ComboboxSelected>>", self.on_buff_skip_mode_change)
        info.pack(fill="x", pady=5)

        save_frame = tk.LabelFrame(left_panel, text="儲存 / 載入")
        save_frame.pack(fill="x", pady=5)

        name_row = tk.Frame(save_frame)
        name_row.pack(fill="x", padx=8, pady=(5, 3))
        tk.Label(name_row, text="腳本名稱：").pack(side="left")
        self.name_entry = tk.Entry(name_row, width=25)
        self.name_entry.pack(side="left", padx=5)

        save_btn_row = tk.Frame(save_frame)
        save_btn_row.pack(fill="x", padx=8, pady=(0, 5))
        tk.Button(save_btn_row, text="保存目前 Timeline", command=self.save_current_timeline).pack(side="left", padx=(0, 5))
        tk.Button(save_btn_row, text="重新整理清單", command=self.refresh_saved_list).pack(side="left", padx=5)
        tk.Button(save_btn_row, text="載入選取項目", command=self.load_selected_timeline).pack(side="left", padx=5)
        tk.Button(save_btn_row, text="刪除選取項目", command=self.delete_selected_timeline).pack(side="left", padx=5)

        self.saved_listbox = tk.Listbox(save_frame, height=5, exportselection=False)
        self.saved_listbox.pack(fill="x", padx=5, pady=5)

        error_frame = tk.LabelFrame(left_panel, text="錯誤訊息（前端 / 後端）")
        error_frame.pack(fill="both", expand=True, pady=(5, 0))
        tk.Label(error_frame, text="前端：", width=8, anchor="w").grid(row=0, column=0, padx=6, pady=2, sticky="nw")
        self.frontend_error_text = tk.Text(error_frame, height=4, wrap="word", fg="#b30000")
        self.frontend_error_text.grid(row=0, column=1, sticky="nsew", padx=(0, 6), pady=2)
        tk.Label(error_frame, text="後端：", width=8, anchor="w").grid(row=1, column=0, padx=6, pady=2, sticky="nw")
        self.backend_error_text = tk.Text(error_frame, height=4, wrap="word", fg="#b30000")
        self.backend_error_text.grid(row=1, column=1, sticky="nsew", padx=(0, 6), pady=(2, 6))
        error_frame.grid_columnconfigure(1, weight=1)
        error_frame.grid_rowconfigure(0, weight=1)
        error_frame.grid_rowconfigure(1, weight=1)
        self.clear_errors()

        jitter_frame = tk.LabelFrame(right_panel, text="jitter 設定")
        jitter_frame.pack(fill="x", pady=(0, 8))

        tk.Label(jitter_frame, text="at jitter").grid(row=0, column=0, padx=5, pady=5)
        self.at_jitter_entry = tk.Entry(jitter_frame, width=10)
        self.at_jitter_entry.insert(0, "0.00")
        self.at_jitter_entry.grid(row=0, column=1, padx=5, pady=5)

        tk.Button(jitter_frame, text="套用到選取列", command=self.apply_jitter_to_selected).grid(row=0, column=2, padx=5, pady=5)
        tk.Button(jitter_frame, text="套用到全部 event", command=self.apply_jitter_to_all).grid(row=0, column=3, padx=5, pady=5)
        tk.Button(jitter_frame, text="選取列清成 0", command=self.clear_jitter_selected).grid(row=0, column=4, padx=5, pady=5)
        tk.Button(
            jitter_frame,
            text="UI 保存",
            command=self.save_ui_layout,
            bg="#fff4b3",
            activebackground="#ffe08a"
        ).grid(row=0, column=5, padx=(5, 8), pady=5)

        columns = ("idx", "type", "button", "at", "at_jitter", "buff_group", "buff_cycle_sec", "buff_jitter_sec", "group")
        self.tree_columns = columns
        self.tree = ttk.Treeview(right_panel, columns=columns, show="headings", height=8, selectmode="extended")
        for col in columns:
            self.tree.heading(col, text=col)
            width = 92
            if col in ("buff_group", "buff_cycle_sec", "buff_jitter_sec"):
                width = 100
            self.tree.column(col, width=width, minwidth=40, stretch=False)
        self.tree.pack(fill="both", expand=True, pady=(0, 8))
        self.tree.tag_configure("copied_group", background="#fff4b3")
        self.tree.bind("<Double-1>", self.on_tree_double_click)
        self.tree.bind("<ButtonPress-1>", self.on_tree_drag_start, add="+")
        self.tree.bind("<ButtonRelease-1>", self.on_tree_drag_end, add="+")

        edit_row = tk.Frame(right_panel)
        edit_row.pack(fill="x", pady=(0, 8))
        tk.Button(edit_row, text="上移", command=self.move_selected_up, width=9).pack(side="left", padx=2)
        tk.Button(edit_row, text="下移", command=self.move_selected_down, width=9).pack(side="left", padx=2)
        tk.Button(edit_row, text="複製列", command=self.duplicate_selected_rows, width=9).pack(side="left", padx=2)
        tk.Button(edit_row, text="刪除列", command=self.delete_selected_rows, width=9).pack(side="left", padx=2)

        bottom = tk.Frame(right_panel)
        bottom.pack(fill="both", expand=True)

        tk.Label(bottom, text="JSON 預覽 / 分析結果").pack(anchor="w")
        self.text = tk.Text(bottom, height=12)
        self.text.pack(fill="both", expand=True)

        self.refresh_saved_list()
        self.restore_last_selected()
        self.root.after(50, self.apply_saved_ui_layout)
        self.root.after(150, self.auto_connect)

    def get_current_paned_sash_x(self):
        self.root.update_idletasks()
        panes = self.body.panes()
        if len(panes) < 2:
            return None
        sash_x, _ = self.body.sash_coord(0)
        return max(0, int(sash_x))

    def apply_saved_ui_layout(self):
        ui_layout = self.config.get("ui_layout", {})
        if not isinstance(ui_layout, dict):
            return

        window_size = ui_layout.get("window_size")
        if (
            isinstance(window_size, (list, tuple))
            and len(window_size) == 2
            and all(isinstance(v, (int, float)) for v in window_size)
        ):
            width = max(900, int(window_size[0]))
            height = max(620, int(window_size[1]))
            self.root.geometry(f"{width}x{height}")
            self.root.update_idletasks()

        widths = ui_layout.get("tree_column_widths", {})
        if isinstance(widths, dict):
            for col in self.tree_columns:
                width = widths.get(col)
                if isinstance(width, (int, float)) and width >= 40:
                    self.tree.column(col, width=int(width), minwidth=40, stretch=False)

        self.root.update_idletasks()
        total_width = self.body.winfo_width()
        if total_width <= 0:
            return

        sash_x = ui_layout.get("paned_sash_x")
        if isinstance(sash_x, (int, float)):
            target_x = int(sash_x)
        else:
            ratio = ui_layout.get("paned_ratio")
            if isinstance(ratio, (int, float)):
                target_x = int(total_width * ratio)
            else:
                return

        min_x = 120
        max_x = max(min_x, total_width - 120)
        self.body.sash_place(0, min(max_x, max(min_x, target_x)), 0)

    def save_ui_layout(self):
        sash_x = self.get_current_paned_sash_x()
        if sash_x is None:
            messagebox.showwarning("提醒", "目前無法取得 UI 版面資訊，請稍後再試")
            return

        column_widths = {}
        for col in self.tree_columns:
            width = self.tree.column(col, option="width")
            try:
                column_widths[col] = int(width)
            except Exception:
                pass

        self.config["ui_layout"] = {
            "window_size": [self.root.winfo_width(), self.root.winfo_height()],
            "paned_sash_x": int(sash_x),
            "tree_column_widths": column_widths
        }
        save_config(self.config)
        self.set_status("已保存 UI 版面（PanedWindow 像素位置 + 欄寬）")

    def update_current_labels(self):
        if self.current_name:
            source_note = "已儲存" if self.current_loaded_from_saved else "目前工作中"
            self.current_script_var.set("【目前腳本：{}（{}）】".format(self.current_name, source_note))
        else:
            self.current_script_var.set("【目前腳本：未命名 / 未儲存】")

    def write_text(self, obj):
        self.text.delete("1.0", tk.END)
        self.text.insert(tk.END, json.dumps(obj, ensure_ascii=False, indent=2))

    def refresh_preview(self):
        self.write_text({
            "current_name": self.current_name,
            "pi_host": self.config["pi_host"],
            "action": "run_timeline",
            "events": self.timeline,
            "simultaneous_groups": build_overlap_summary(self.timeline)
        })

    def refresh_tree(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        summary = build_overlap_summary(self.timeline)
        event_to_group = {}

        for grp in summary:
            group_id = grp["group"]
            for ev in grp["events"]:
                key = (ev["type"], ev["button"], ev["at"])
                event_to_group[key] = group_id

        for i, ev in enumerate(self.timeline):
            key = (ev["type"], ev["button"], ev["at"])
            grp = event_to_group.get(key, "")
            buff_group = str(ev.get("buff_group", "")).strip()
            tags = ("copied_group",) if buff_group.startswith("-") else ()
            self.tree.insert("", "end", iid=str(i), values=(
                i,
                ev["type"],
                ev["button"],
                ev["at"],
                ev["at_jitter"],
                buff_group,
                ev.get("buff_cycle_sec", 0.0),
                ev.get("buff_jitter_sec", 0.0),
                grp
            ), tags=tags)

    def refresh_saved_list(self):
        names = list_saved_timeline_names()
        self.saved_listbox.delete(0, tk.END)
        for name in names:
            self.saved_listbox.insert(tk.END, name)

    def restore_last_selected(self):
        last_name = self.config.get("last_selected_name", "")
        if not last_name:
            self.update_current_labels()
            return

        names = list_saved_timeline_names()
        if last_name in names:
            idx = names.index(last_name)
            self.saved_listbox.selection_clear(0, tk.END)
            self.saved_listbox.selection_set(idx)
            self.saved_listbox.activate(idx)

        self.update_current_labels()

    def get_selected_saved_name(self):
        selection = self.saved_listbox.curselection()
        if not selection:
            return ""
        return self.saved_listbox.get(selection[0])

    def save_pi_ip(self):
        ip = self.pi_ip_entry.get().strip()
        if not ip:
            self.set_frontend_error("Pi IP 不可空白")
            messagebox.showwarning("提醒", "Pi IP 不可空白")
            return

        self.config["pi_host"] = ip
        save_config(self.config)
        self.update_current_labels()
        self.set_frontend_error("")
        self.set_status("已保存 Pi IP：{}".format(ip))
        self.set_connected(False, "IP 已更新，請重新測試連線")

    def auto_connect(self):
        self.ping_pi(show_popup=False)

    def go_offline(self):
        self.config["pi_host"] = self.pi_ip_entry.get().strip() or DEFAULT_PI_HOST
        save_config(self.config)
        self.update_current_labels()
        try:
            if self.connected:
                self.request_pi({"action": "stop"}, write_response=False)
        except Exception as e:
            self.set_frontend_error(str(e))
        self.close_connection(silent=True)
        self.offline_mode = True
        self.set_connected(False, "已離線（GPIO 釋放指令已送出）")

    def start_record(self):
        global recording, recording_start, events, pressed_keys
        events = []
        pressed_keys = set()
        recording = True
        recording_start = time.time()
        self.timeline = []
        self.timeline_meta = self.new_meta()
        self.current_loaded_from_saved = False
        self.set_frontend_error("")
        self.set_status("錄製中...")
        self.write_text({"status": "recording"})
        self.refresh_tree()
        self.record_button.config(text="停止錄製")
        self.record_button.config(bg="#f7e37a")

    def stop_record(self):
        global recording
        recording = False
        self.record_button.config(text="開始錄製")
        self.record_button.config(bg="#9be58b")
        if events:
            self.analyze()
        else:
            self.set_status("已停止錄製，共 0 筆事件")
            self.write_text(events)

    def toggle_record(self):
        if recording:
            self.stop_record()
        else:
            self.start_record()

    def analyze(self):
        global events, recording_start
        if not events:
            if self.restore_original_timeline():
                self.set_status("已還原原始 Timeline（重新分析）")
                return
            messagebox.showwarning("提醒", "目前沒有錄到資料，也沒有可還原的原始 Timeline")
            return

        self.timeline = build_timeline(events, recording_start)
        self.timeline_meta = self.new_meta()
        self.mark_timeline_dirty()

        if not self.current_name:
            self.current_name = "未命名_{}".format(time.strftime("%H%M%S"))

        self.name_entry.delete(0, tk.END)
        self.name_entry.insert(0, self.current_name)

        self.set_status("Timeline 分析完成，共 {} 筆 event".format(len(self.timeline)))

    def save_current_timeline(self):
        if not self.timeline:
            messagebox.showwarning("提醒", "目前沒有 timeline 可保存")
            return

        name = self.name_entry.get().strip()
        if not name:
            messagebox.showwarning("提醒", "請先輸入保存名稱")
            return

        name = sanitize_filename(name)
        path = timeline_file_path(name)

        if os.path.exists(path):
            yes = messagebox.askyesno("確認覆蓋", "名稱 '{}' 已存在，是否覆蓋？".format(name))
            if not yes:
                return

        try:
            self.push_history("before_save")
            self.timeline_meta["original_events"] = self.copy_events(self.timeline)
            self.validate_negative_group_monotonic_by_index(self.timeline)
            recalculated = self.recalculate_timeline_for_runtime()
        except Exception as e:
            messagebox.showerror("保存失敗", str(e))
            return

        self.timeline = recalculated
        save_named_timeline(name, self.timeline, self.timeline_meta)
        self.current_name = name
        self.current_loaded_from_saved = True
        self.config["last_selected_name"] = name
        save_config(self.config)

        self.refresh_saved_list()
        self.select_saved_name(name)
        self.mark_timeline_dirty()
        self.set_status("已保存 Timeline：{}".format(name))

    def select_saved_name(self, name):
        names = list_saved_timeline_names()
        self.saved_listbox.selection_clear(0, tk.END)
        if name in names:
            idx = names.index(name)
            self.saved_listbox.selection_set(idx)
            self.saved_listbox.activate(idx)

    def load_selected_timeline(self):
        name = self.get_selected_saved_name()
        if not name:
            messagebox.showwarning("提醒", "請先從清單選一個已保存項目")
            return

        try:
            data = load_named_timeline(name)
        except Exception as e:
            messagebox.showerror("載入失敗", str(e))
            return

        self.timeline = [self.normalize_event_schema(ev) for ev in data.get("events", [])]
        self.timeline_meta = self.normalize_meta(data.get("_meta", {}))
        self.current_name = data.get("name", name)
        self.current_loaded_from_saved = True
        self.config["last_selected_name"] = self.current_name
        save_config(self.config)

        self.name_entry.delete(0, tk.END)
        self.name_entry.insert(0, self.current_name)

        self.refresh_tree()
        self.refresh_preview()
        self.update_current_labels()
        self.set_status("已載入：{}".format(self.current_name))

    def delete_selected_timeline(self):
        name = self.get_selected_saved_name()
        if not name:
            messagebox.showwarning("提醒", "請先選取要刪除的項目")
            return

        yes = messagebox.askyesno("確認刪除", "確定要刪除 '{}' 嗎？".format(name))
        if not yes:
            return

        try:
            delete_named_timeline(name)
        except Exception as e:
            messagebox.showerror("刪除失敗", str(e))
            return

        if self.current_name == name:
            self.current_name = ""
            self.current_loaded_from_saved = False
            self.timeline_meta = self.new_meta()

        if self.config.get("last_selected_name") == name:
            self.config["last_selected_name"] = ""
            save_config(self.config)

        self.refresh_saved_list()
        self.update_current_labels()
        self.set_status("已刪除：{}".format(name))

    def get_selected_indexes(self):
        selection = self.tree.selection()
        return [int(iid) for iid in selection]

    def apply_jitter_to_selected(self):
        if not self.timeline:
            messagebox.showwarning("提醒", "目前沒有 timeline 資料")
            return

        selected = self.get_selected_indexes()
        if not selected:
            messagebox.showwarning("提醒", "請先選取一列或多列")
            return

        try:
            jitter = float(self.at_jitter_entry.get().strip())
        except ValueError:
            messagebox.showerror("錯誤", "at jitter 必須是數字")
            return

        for idx in selected:
            self.timeline[idx]["at_jitter"] = jitter

        self.mark_timeline_dirty()
        for idx in selected:
            self.tree.selection_add(str(idx))
        self.set_status("已套用 jitter 到 {} 筆選取列".format(len(selected)))

    def apply_jitter_to_all(self):
        if not self.timeline:
            messagebox.showwarning("提醒", "目前沒有 timeline 資料")
            return

        try:
            jitter = float(self.at_jitter_entry.get().strip())
        except ValueError:
            messagebox.showerror("錯誤", "at jitter 必須是數字")
            return

        for ev in self.timeline:
            ev["at_jitter"] = jitter

        self.mark_timeline_dirty()
        self.set_status("已套用 jitter 到全部 event")

    def clear_jitter_selected(self):
        if not self.timeline:
            messagebox.showwarning("提醒", "目前沒有 timeline 資料")
            return

        selected = self.get_selected_indexes()
        if not selected:
            messagebox.showwarning("提醒", "請先選取一列或多列")
            return

        for idx in selected:
            self.timeline[idx]["at_jitter"] = 0.0

        self.mark_timeline_dirty()
        for idx in selected:
            self.tree.selection_add(str(idx))
        self.set_status("已將選取列 jitter 清為 0")

    def ping_pi(self, show_popup=True):
        try:
            self.config["pi_host"] = self.pi_ip_entry.get().strip() or DEFAULT_PI_HOST
            save_config(self.config)
            self.update_current_labels()

            self.offline_mode = False
            self.set_frontend_error("")
            self.open_connection(timeout=1.5)
            self.request_pi({"action": "ping"})
            self.set_connected(True, "Pi 連線正常：{}".format(self.config["pi_host"]))
        except Exception as e:
            self.set_frontend_error(str(e))
            self.close_connection(silent=True)
            self.set_connected(False)
            if show_popup:
                messagebox.showerror("Pi 連線失敗", str(e))

    def stop_pi(self):
        try:
            self.config["pi_host"] = self.pi_ip_entry.get().strip() or DEFAULT_PI_HOST
            save_config(self.config)
            self.update_current_labels()

            self.set_frontend_error("")
            self.request_pi({"action": "stop"})
            self.set_status("已停止 Pi：{}".format(self.config["pi_host"]))
        except Exception as e:
            self.set_frontend_error(str(e))
            messagebox.showerror("停止失敗", str(e))

    def send_timeline(self):
        if not self.timeline:
            messagebox.showwarning("提醒", "請先錄製並分析，或載入已保存項目")
            return

        self.config["pi_host"] = self.pi_ip_entry.get().strip() or DEFAULT_PI_HOST
        save_config(self.config)
        self.update_current_labels()

        prepared_events, resolve_note = self.prepare_events_for_send(action_reason="before_send")
        if prepared_events is None:
            return

        payload = {
            "action": "run_timeline",
            "events": prepared_events,
            "buff_skip_mode": self.config.get("buff_skip_mode", BUFF_SKIP_MODE_COMPRESS)
        }

        display_name = self.current_name if self.current_name else "未命名資料"

        try:
            self.set_frontend_error("")
            res = self.request_pi(payload, write_response=False)
            self.write_text({
                "sending_name": display_name,
                "pi_host": self.config["pi_host"],
                "request": payload,
                "response": res
            })

            if res.get("status") == "stopped":
                self.set_status("Pi 已停止執行：{} -> {}".format(display_name, self.config["pi_host"]))
            elif res.get("status") in ("error", "busy"):
                self.set_status("送出失敗：{} -> {}".format(display_name, self.config["pi_host"]))
            else:
                suffix = ""
                if resolve_note:
                    suffix = "（{}）".format(resolve_note)
                self.set_status("已送出：{} -> {}{}".format(display_name, self.config["pi_host"], suffix))
        except Exception as e:
            self.set_frontend_error(str(e))
            messagebox.showerror("傳送失敗", str(e))

    def send_timeline_loop(self):
        if not self.timeline:
            messagebox.showwarning("提醒", "請先錄製並分析，或載入已保存項目")
            return

        self.config["pi_host"] = self.pi_ip_entry.get().strip() or DEFAULT_PI_HOST
        save_config(self.config)
        self.update_current_labels()

        prepared_events, resolve_note = self.prepare_events_for_send(action_reason="before_send_loop")
        if prepared_events is None:
            return

        payload = {
            "action": "run_timeline_loop",
            "events": prepared_events,
            "buff_skip_mode": self.config.get("buff_skip_mode", BUFF_SKIP_MODE_COMPRESS)
        }

        display_name = self.current_name if self.current_name else "未命名資料"

        try:
            self.set_frontend_error("")
            res = self.request_pi(payload, write_response=False)
            self.write_text({
                "sending_name": display_name,
                "pi_host": self.config["pi_host"],
                "request": payload,
                "response": res
            })

            if res.get("status") in ("error", "busy"):
                self.set_status("重複送出失敗：{} -> {}".format(display_name, self.config["pi_host"]))
            else:
                suffix = ""
                if resolve_note:
                    suffix = "（{}）".format(resolve_note)
                self.set_status("已開始重複執行：{} -> {}{}".format(display_name, self.config["pi_host"], suffix))
        except Exception as e:
            self.set_frontend_error(str(e))
            messagebox.showerror("重複傳送失敗", str(e))

    def prepare_events_for_send(self, action_reason="before_send"):
        try:
            self.push_history(action_reason)
            self.timeline_meta["original_events"] = self.copy_events(self.timeline)
            self.validate_negative_group_monotonic_by_index(self.timeline)
            self.timeline = self.recalculate_timeline_for_runtime()
            self.current_loaded_from_saved = False
            self.update_current_labels()
            self.refresh_tree()
            self.refresh_preview()
            self.persist_current_timeline_if_named()
        except Exception as e:
            messagebox.showerror("時間重算失敗", str(e))
            return None, ""

        events = [self.normalize_event_schema(ev) for ev in self.timeline]
        group_configs = {}
        for idx, ev in enumerate(events):
            group_name = ev.get("buff_group", "").strip()
            if not group_name:
                continue
            cycle = float(ev.get("buff_cycle_sec", 0.0))
            jitter = float(ev.get("buff_jitter_sec", 0.0))
            if cycle <= 0:
                continue
            group_configs.setdefault(group_name, []).append({
                "idx": idx,
                "cycle": cycle,
                "jitter": jitter
            })

        resolved_groups = []
        for group_name, cfg_rows in sorted(group_configs.items()):
            uniq = {}
            for row in cfg_rows:
                key = (row["cycle"], row["jitter"])
                if key not in uniq:
                    uniq[key] = []
                uniq[key].append(row["idx"])

            if len(uniq) == 1:
                (chosen_cycle, chosen_jitter), _ = next(iter(uniq.items()))
            else:
                options = []
                option_map = {}
                for cfg_idx, ((cycle, jitter), idx_list) in enumerate(sorted(uniq.items()), start=1):
                    sample_idx = idx_list[0]
                    options.append("{}. idx{} => cycle={}, jitter={}".format(cfg_idx, sample_idx, cycle, jitter))
                    option_map[str(cfg_idx)] = (cycle, jitter)

                ans = simpledialog.askstring(
                    "buff 參數衝突",
                    "buff_group {} 有多組秒數設定：\n{}\n請輸入要採用的選項編號：".format(
                        group_name,
                        "\n".join(options)
                    ),
                    initialvalue="1"
                )
                if ans is None:
                    return None, ""
                ans = ans.strip()
                if ans not in option_map:
                    messagebox.showerror("輸入錯誤", "選項不存在：{}".format(ans))
                    return None, ""
                chosen_cycle, chosen_jitter = option_map[ans]
                resolved_groups.append(str(group_name))

            for ev in events:
                if ev.get("buff_group", "").strip() == group_name:
                    ev["buff_cycle_sec"] = chosen_cycle
                    ev["buff_jitter_sec"] = chosen_jitter

        resolve_note = ""
        if resolved_groups:
            resolve_note = "已解決衝突群組: {}".format("、".join(resolved_groups))
        return events, resolve_note

    def on_buff_skip_mode_change(self, _event=None):
        label = self.buff_skip_mode_combo.get().strip()
        if label.startswith("走過"):
            mode = BUFF_SKIP_MODE_WALK
        else:
            mode = BUFF_SKIP_MODE_COMPRESS
        self.buff_skip_mode_var.set(mode)
        self.config["buff_skip_mode"] = mode
        save_config(self.config)

    def on_tree_double_click(self, event):
        row_id = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)
        if not row_id or not col_id:
            return

        idx = int(row_id)
        col_index = int(col_id[1:]) - 1
        columns = ("idx", "type", "button", "at", "at_jitter", "buff_group", "buff_cycle_sec", "buff_jitter_sec", "group")
        field = columns[col_index]
        if field not in ("type", "button", "at", "at_jitter", "buff_group", "buff_cycle_sec", "buff_jitter_sec"):
            return

        current_value = str(self.timeline[idx].get(field, ""))
        new_value = simpledialog.askstring("修改欄位", "請輸入 {}:".format(field), initialvalue=current_value)
        if new_value is None:
            return

        try:
            if field in ("at", "at_jitter", "buff_cycle_sec", "buff_jitter_sec"):
                value = float(new_value)
                if field == "at":
                    value = max(0.0, value)
                else:
                    value = abs(value)
                self.timeline[idx][field] = round(value, 4)
            else:
                value = new_value.strip().lower()
                if field == "type" and value not in ("press", "release"):
                    raise ValueError("type 只能是 press / release")
                if field == "button" and not value:
                    raise ValueError("button 不可空白")
                if field == "buff_group":
                    value = new_value.strip()
                self.timeline[idx][field] = value
        except Exception as e:
            messagebox.showerror("修改失敗", str(e))
            return

        self.mark_timeline_dirty()
        self.tree.selection_set(str(idx))

    def on_tree_drag_start(self, event):
        self.drag_iid = self.tree.identify_row(event.y)

    def on_tree_drag_end(self, event):
        if not self.drag_iid:
            return
        target_iid = self.tree.identify_row(event.y)
        if not target_iid or target_iid == self.drag_iid:
            self.drag_iid = None
            return

        src = int(self.drag_iid)
        dst = int(target_iid)
        ev = self.timeline.pop(src)
        self.timeline.insert(dst, ev)
        self.mark_timeline_dirty()
        self.tree.selection_set(str(dst))
        self.drag_iid = None

    def move_selected_up(self):
        selected = sorted(self.get_selected_indexes())
        if not selected:
            messagebox.showwarning("提醒", "請先選取列")
            return
        if selected[0] == 0:
            return
        for idx in selected:
            self.timeline[idx - 1], self.timeline[idx] = self.timeline[idx], self.timeline[idx - 1]
        self.mark_timeline_dirty()
        self.tree.selection_set([str(i - 1) for i in selected])

    def move_selected_down(self):
        selected = sorted(self.get_selected_indexes(), reverse=True)
        if not selected:
            messagebox.showwarning("提醒", "請先選取列")
            return
        if selected[0] == len(self.timeline) - 1:
            return
        for idx in selected:
            self.timeline[idx + 1], self.timeline[idx] = self.timeline[idx], self.timeline[idx + 1]
        self.mark_timeline_dirty()
        self.tree.selection_set([str(i + 1) for i in sorted(selected)])

    def duplicate_selected_rows(self):
        selected = sorted(self.get_selected_indexes())
        if not selected:
            messagebox.showwarning("提醒", "請先選取列")
            return
        offset = 0
        new_ids = []
        for idx in selected:
            insert_at = idx + 1 + offset
            copied = dict(self.timeline[idx + offset])
            self.timeline.insert(insert_at, copied)
            new_ids.append(insert_at)
            offset += 1
        self.mark_timeline_dirty()
        self.tree.selection_set([str(i) for i in new_ids])

    def delete_selected_rows(self):
        selected = sorted(self.get_selected_indexes(), reverse=True)
        if not selected:
            messagebox.showwarning("提醒", "請先選取列")
            return
        for idx in selected:
            self.timeline.pop(idx)
        self.mark_timeline_dirty()
        self.set_status("已刪除 {} 列".format(len(selected)))


def start_listener():
    global listener
    listener = keyboard.Listener(on_press=on_press, on_release=on_release)
    listener.start()


def main():
    start_listener()
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
