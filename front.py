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
NEGATIVE_GROUP_ANCHOR_GAP_SEC = 0.2

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
            "send_delay_sec": 1.0,
            "last_selected_name": "",
            "buff_skip_mode": BUFF_SKIP_MODE_COMPRESS,
            "manual_offset_sec": 0.2,
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
            "send_delay_sec": float(data.get("send_delay_sec", 1.0)),
            "last_selected_name": data.get("last_selected_name", ""),
            "buff_skip_mode": data.get("buff_skip_mode", BUFF_SKIP_MODE_COMPRESS),
            "manual_offset_sec": float(data.get("manual_offset_sec", NEGATIVE_GROUP_ANCHOR_GAP_SEC)),
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
            "send_delay_sec": 1.0,
            "last_selected_name": "",
            "buff_skip_mode": BUFF_SKIP_MODE_COMPRESS,
            "manual_offset_sec": 0.2,
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


def recalculate_runtime_events_by_index(events, anchor_gap_sec):
    if not events:
        return events

    original_at_list = [float(ev.get("at", 0.0)) for ev in events]
    seen_negative_group = {}
    cursor_at = 0.0
    i = 0

    while i < len(events):
        ev = events[i]
        group_name = str(ev.get("buff_group", "")).strip()
        is_negative_group = group_name.startswith("-")
        if not is_negative_group:
            original_at = original_at_list[i]
            prev_anchor_idx = i - 1
            while prev_anchor_idx >= 0:
                prev_group = str(events[prev_anchor_idx].get("buff_group", "")).strip()
                if not prev_group.startswith("-"):
                    break
                prev_anchor_idx -= 1

            if i == 0:
                shifted_at = max(0.0, original_at)
            elif prev_anchor_idx >= 0:
                prev_anchor_original_at = original_at_list[prev_anchor_idx]
                follow_gap_sec = max(0.0, original_at - prev_anchor_original_at)
                shifted_at = max(0.0, cursor_at + follow_gap_sec)
            else:
                follow_gap_sec = max(0.0, original_at - original_at_list[0])
                shifted_at = max(0.0, cursor_at + follow_gap_sec)
            events[i]["at"] = round(shifted_at, 2)
            cursor_at = shifted_at
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
        segment_base_at = original_at_list[start]
        relative_rows = []
        for local_idx, row in enumerate(segment):
            absolute_idx = start + local_idx
            relative_at = round(max(0.0, original_at_list[absolute_idx] - segment_base_at), 4)
            relative_rows.append((row, relative_at))

        if start == 0:
            anchor = max(0.0, segment_base_at)
        else:
            anchor = max(0.0, cursor_at + anchor_gap_sec)

        duration = max(relative_at for _, relative_at in relative_rows) if relative_rows else 0.0
        for row, relative_at in relative_rows:
            row["at"] = round(anchor + relative_at, 2)

        cursor_at = anchor + duration
        i = end

    return events


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
    def _is_negative_buff_group(self, value):
        group = str(value or "").strip()
        return group.startswith("-")

    def _normalize_replicated_row_flag(self, value):
        if isinstance(value, bool):
            return 1 if value else 0
        if isinstance(value, (int, float)):
            return 1 if int(value) != 0 else 0
        text = str(value or "").strip().lower()
        if text in ("1", "true", "yes", "y"):
            return 1
        return 0

    def _sync_replicated_row(self, row):
        buff_group = str(row.get("buff_group", "")).strip()
        replicated = self._normalize_replicated_row_flag(row.get("replicatedRow", 0))
        if self._is_negative_buff_group(buff_group):
            replicated = 1
        row["buff_group"] = buff_group
        row["replicatedRow"] = replicated
        return row

    def normalize_event_schema(self, ev):
        row = dict(ev)
        row.setdefault("type", "")
        row.setdefault("button", "")
        row["at"] = round(max(0.0, float(row.get("at", 0.0))), 2)
        row["at_jitter"] = round(abs(float(row.get("at_jitter", 0.0))), 4)
        row["buff_group"] = str(row.get("buff_group", "")).strip()
        row["buff_cycle_sec"] = round(max(0.0, float(row.get("buff_cycle_sec", 0.0))), 4)
        row["buff_jitter_sec"] = round(abs(float(row.get("buff_jitter_sec", 0.0))), 4)
        return self._sync_replicated_row(row)

    def get_manual_offset_sec(self):
        raw = ""
        if hasattr(self, "offset_sec_entry"):
            raw = self.offset_sec_entry.get().strip()
        if not raw:
            raw = str(self.config.get("manual_offset_sec", NEGATIVE_GROUP_ANCHOR_GAP_SEC))
        try:
            val = float(raw)
        except ValueError:
            raise ValueError("自/手動偏移時間必須是數字")
        self.config["manual_offset_sec"] = val
        save_config(self.config)
        return val

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

    def recalculate_timeline_for_runtime(self, anchor_gap_sec=None):
        events = self.copy_events(self.timeline)
        if not events:
            return events
        if anchor_gap_sec is None:
            anchor_gap_sec = float(self.config.get("manual_offset_sec", NEGATIVE_GROUP_ANCHOR_GAP_SEC))
        return recalculate_runtime_events_by_index(events, anchor_gap_sec)

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
        self.clear_runtime_highlight()
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
        self.conn = None
        self.conn_file = None
        self.timeline_runtime_info = {"events": []}
        self.timeline_runtime_by_index = {}
        self.runtime_recent_ok_indices = []
        self.runtime_recent_skipped_indices = []
        self.runtime_latest_index = None
        self.last_runtime_signature = ""

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
        tk.Label(row, text="送出延遲(秒)：").pack(side="left", padx=(8, 0))
        self.send_delay_entry = tk.Entry(row, width=8)
        self.send_delay_entry.insert(0, str(self.config.get("send_delay_sec", 1.0)))
        self.send_delay_entry.pack(side="left", padx=5)
        tk.Button(row, text="保存", command=self.save_apply_info).pack(side="left", padx=5)

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

        jitter_frame = tk.LabelFrame(right_panel, text="timeline 設定")
        jitter_frame.pack(fill="x", pady=(0, 8))

        tk.Label(jitter_frame, text="tip : 多重選取項目後雙擊title可批量設定").grid(
            row=0, column=0, padx=(8, 5), pady=5, sticky="w"
        )
        tk.Button(
            jitter_frame,
            text="UI 保存",
            command=self.save_ui_layout,
            bg="#fff4b3",
            activebackground="#ffe08a"
        ).grid(row=0, column=1, padx=5, pady=5)
        tk.Button(jitter_frame, text="計算偏移量", command=self.calculate_offsets_only).grid(
            row=0, column=2, padx=(5, 8), pady=5
        )
        tk.Label(
            jitter_frame,
            text="【註: buff_goup 如果設定 -1 可計算偏移量，不同複製按鈕群請用不同負值 】"
        ).grid(row=1, column=0, columnspan=3, padx=(8, 8), pady=(0, 6), sticky="w")

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
        self.tree.tag_configure("runtime_ok_1", background="#bfe8bf")
        self.tree.tag_configure("runtime_ok_2", background="#d9f2d9")
        self.tree.tag_configure("runtime_ok_3", background="#edf9ed")
        self.tree.bind("<Double-1>", self.on_tree_double_click)
        self.tree.bind("<Control-a>", self.on_tree_select_all, add="+")
        self.tree.bind("<Control-A>", self.on_tree_select_all, add="+")
        self.tree.bind("<Control-c>", self.on_tree_copy, add="+")
        self.tree.bind("<Control-C>", self.on_tree_copy, add="+")
        self.tree.bind("<Control-v>", self.on_tree_paste, add="+")
        self.tree.bind("<Control-V>", self.on_tree_paste, add="+")
        self.tree.bind("<Configure>", self._schedule_tree_overlay_refresh, add="+")
        self.tree.bind("<MouseWheel>", self._schedule_tree_overlay_refresh, add="+")
        self.tree.bind("<Button-4>", self._schedule_tree_overlay_refresh, add="+")
        self.tree.bind("<Button-5>", self._schedule_tree_overlay_refresh, add="+")
        self.tree.bind("<Expose>", self._schedule_tree_overlay_refresh, add="+")
        self.tree_overlay_labels = []

        edit_row = tk.Frame(right_panel)
        edit_row.pack(fill="x", pady=(0, 8))
        tk.Button(edit_row, text="上移", command=self.move_selected_up, width=9).pack(side="left", padx=2)
        tk.Button(edit_row, text="下移", command=self.move_selected_down, width=9).pack(side="left", padx=2)
        tk.Button(edit_row, text="複製列", command=self.duplicate_selected_rows, width=9).pack(side="left", padx=2)
        tk.Button(edit_row, text="刪除列", command=self.delete_selected_rows, width=9).pack(side="left", padx=2)
        tk.Label(edit_row, text="自/手動偏移 :").pack(side="left", padx=(14, 5))
        self.offset_sec_entry = tk.Entry(edit_row, width=10)
        self.offset_sec_entry.insert(0, "{:.3f}".format(float(self.config.get("manual_offset_sec", NEGATIVE_GROUP_ANCHOR_GAP_SEC))))
        self.offset_sec_entry.pack(side="left", padx=(0, 5))
        tk.Label(edit_row, text="秒").pack(side="left", padx=(0, 5))
        tk.Button(edit_row, text="套用偏移", command=self.apply_offset_from_selected).pack(side="left", padx=2)

        bottom = tk.Frame(right_panel)
        bottom.pack(fill="both", expand=True)

        tk.Label(bottom, text="JSON 預覽 / 分析結果").pack(anchor="w")
        self.text = tk.Text(bottom, height=12)
        self.text.pack(fill="both", expand=True)

        self.refresh_saved_list()
        self.restore_last_selected()
        self.root.after(50, self.apply_saved_ui_layout)
        self.root.after(150, self.auto_connect)
        self.root.after(200, self.poll_runtime_status)

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

    def update_runtime_from_status(self, status_response):
        runtime = status_response.get("timeline_runtime", {})
        if not isinstance(runtime, dict):
            runtime = {}
        events = runtime.get("events", [])
        if not isinstance(events, list):
            events = []

        runtime_by_index = {}
        recent_ok = []
        recent_skipped = []
        latest_idx = None
        for ev in events:
            if not isinstance(ev, dict):
                continue
            try:
                idx = int(ev.get("original_index"))
            except Exception:
                continue
            runtime_by_index[idx] = ev
            latest_idx = idx

        for ev in reversed(events):
            if not isinstance(ev, dict):
                continue
            try:
                idx = int(ev.get("original_index"))
            except Exception:
                continue
            status = str(ev.get("status", "")).strip().lower()
            if status == "ok" and idx not in recent_ok:
                recent_ok.append(idx)
                if len(recent_ok) >= 3:
                    break

        for ev in reversed(events):
            if not isinstance(ev, dict):
                continue
            try:
                idx = int(ev.get("original_index"))
            except Exception:
                continue
            status = str(ev.get("status", "")).strip().lower()
            if status == "skipped_by_cooldown" and idx not in recent_skipped:
                recent_skipped.append(idx)
                if len(recent_skipped) >= 2:
                    break

        runtime["events"] = events
        signature = json.dumps(
            {
                "run_id": runtime.get("run_id"),
                "state": runtime.get("state"),
                "processed_count": runtime.get("processed_count"),
                "last_event": runtime.get("last_event"),
                "recent_ok": recent_ok,
                "recent_skipped": recent_skipped
            },
            ensure_ascii=False,
            sort_keys=True
        )
        changed = (
            signature != self.last_runtime_signature
            or runtime_by_index != self.timeline_runtime_by_index
        )
        self.timeline_runtime_info = runtime
        self.timeline_runtime_by_index = runtime_by_index
        self.runtime_recent_ok_indices = recent_ok
        self.runtime_recent_skipped_indices = recent_skipped
        self.runtime_latest_index = latest_idx
        self.last_runtime_signature = signature
        return changed

    def poll_runtime_status(self):
        try:
            if not self.offline_mode and self.conn is not None and self.connected:
                res = self.request_pi({"action": "status"}, write_response=False)
                if isinstance(res, dict) and self.update_runtime_from_status(res):
                    self.refresh_tree()
                    state = str(self.timeline_runtime_info.get("state", "")).strip().lower()
                    if state == "running":
                        self.focus_latest_runtime_row()
        except Exception:
            pass
        finally:
            self.root.after(200, self.poll_runtime_status)

    def clear_runtime_highlight(self):
        self.timeline_runtime_info = {"events": []}
        self.timeline_runtime_by_index = {}
        self.runtime_recent_ok_indices = []
        self.runtime_recent_skipped_indices = []
        self.runtime_latest_index = None
        self.last_runtime_signature = ""

    def focus_latest_runtime_row(self):
        if self.runtime_latest_index is None:
            return
        row_id = str(self.runtime_latest_index)
        if not self.tree.exists(row_id):
            return
        self.tree.see(row_id)
        self._schedule_tree_overlay_refresh()

    def _schedule_tree_overlay_refresh(self, _event=None):
        if not hasattr(self, "tree"):
            return
        if hasattr(self, "_overlay_refresh_after_id") and self._overlay_refresh_after_id is not None:
            try:
                self.root.after_cancel(self._overlay_refresh_after_id)
            except Exception:
                pass
        self._overlay_refresh_after_id = self.root.after(10, self._refresh_tree_buff_group_overlays)

    def _refresh_tree_buff_group_overlays(self):
        self._overlay_refresh_after_id = None
        if not hasattr(self, "tree"):
            return

        for label in getattr(self, "tree_overlay_labels", []):
            try:
                label.destroy()
            except Exception:
                pass
        self.tree_overlay_labels = []

        for iid in self.tree.get_children():
            try:
                idx = int(iid)
            except Exception:
                continue
            if idx < 0 or idx >= len(self.timeline):
                continue

            ev = self.timeline[idx]
            original_buff_group = str(ev.get("buff_group", "")).strip()
            is_replicated = self._normalize_replicated_row_flag(ev.get("replicatedRow", 0)) == 1
            runtime_event = self.timeline_runtime_by_index.get(idx, {})
            runtime_status = str(runtime_event.get("status", "")).strip().lower()

            bg_color = None
            display_text = original_buff_group
            if runtime_status == "skipped_by_cooldown":
                bg_color = "#e3e3e3"
            elif is_replicated:
                bg_color = "#fff4b3"

            if not bg_color:
                continue

            bbox = self.tree.bbox(iid, column="buff_group")
            if not bbox:
                continue
            x, y, w, h = bbox
            label = tk.Label(
                self.tree,
                text=display_text,
                background=bg_color,
                font=ttk.Style().lookup("Treeview", "font"),
                anchor="w",
                padx=4,
                borderwidth=0,
                highlightthickness=0
            )
            label.place(x=x, y=y, width=w, height=h)
            self.tree_overlay_labels.append(label)

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
            self._sync_replicated_row(ev)
            key = (ev["type"], ev["button"], ev["at"])
            grp = event_to_group.get(key, "")
            original_buff_group = str(ev.get("buff_group", "")).strip()
            buff_group = original_buff_group
            is_replicated = self._normalize_replicated_row_flag(ev.get("replicatedRow", 0)) == 1
            at_value = "{:.2f}".format(float(ev["at"]))
            tags = []
            is_recent_ok = i in self.runtime_recent_ok_indices
            if is_recent_ok:
                recent_rank = self.runtime_recent_ok_indices.index(i)
                if recent_rank == 0:
                    tags.append("runtime_ok_1")
                elif recent_rank == 1:
                    tags.append("runtime_ok_2")
                else:
                    tags.append("runtime_ok_3")
            elif is_replicated:
                tags.append("copied_group")
            self.tree.insert("", "end", iid=str(i), values=(
                i,
                ev["type"],
                ev["button"],
                at_value,
                ev["at_jitter"],
                buff_group,
                ev.get("buff_cycle_sec", 0.0),
                ev.get("buff_jitter_sec", 0.0),
                grp
            ), tags=tuple(tags))
        self._schedule_tree_overlay_refresh()

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

    def parse_send_delay_sec(self):
        raw = self.send_delay_entry.get().strip()
        try:
            delay = float(raw)
        except ValueError:
            raise ValueError("送出延遲秒數必須是數字")
        if delay < 0:
            raise ValueError("送出延遲秒數不可小於 0")
        return delay

    def save_apply_info(self):
        ip = self.pi_ip_entry.get().strip()
        if not ip:
            self.set_frontend_error("Pi IP 不可空白")
            messagebox.showwarning("提醒", "Pi IP 不可空白")
            return
        try:
            delay = self.parse_send_delay_sec()
        except ValueError as e:
            self.set_frontend_error(str(e))
            messagebox.showwarning("提醒", str(e))
            return

        self.config["pi_host"] = ip
        self.config["send_delay_sec"] = delay
        save_config(self.config)
        self.update_current_labels()
        self.set_frontend_error("")
        self.set_status("已保存 Pi IP：{}，送出延遲 {} 秒".format(ip, delay))
        self.set_connected(False, "設定已更新，請重新測試連線")

    def apply_send_delay_if_needed(self):
        delay = self.parse_send_delay_sec()
        self.config["send_delay_sec"] = delay
        save_config(self.config)
        if delay <= 0:
            return 0.0
        self.set_status("延遲 {} 秒後送出...".format(delay))
        self.root.update_idletasks()
        time.sleep(delay)
        return delay

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
            offset_sec = self.get_manual_offset_sec()
            self.push_history("before_save")
            self.timeline = [self._sync_replicated_row(ev) for ev in self.timeline]
            self.timeline_meta["original_events"] = self.copy_events(self.timeline)
            self.validate_negative_group_monotonic_by_index(self.timeline)
            recalculated = self.recalculate_timeline_for_runtime(anchor_gap_sec=offset_sec)
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

    def calculate_offsets_only(self):
        if not self.timeline:
            messagebox.showwarning("提醒", "目前沒有 timeline 資料")
            return

        prepared_events, _ = self.prepare_events_for_send(action_reason="calculate_offset_only")
        if prepared_events is None:
            return
        self.set_status("已計算偏移量並套用到 timeline（共 {} 筆）".format(len(prepared_events)))

    def apply_offset_from_selected(self):
        if not self.timeline:
            messagebox.showwarning("提醒", "目前沒有 timeline 資料")
            return

        selected = sorted(self.get_selected_indexes())
        if not selected:
            messagebox.showwarning("提醒", "請先選取一列或多列")
            return

        try:
            offset_sec = self.get_manual_offset_sec()
        except Exception as e:
            messagebox.showerror("錯誤", str(e))
            return

        start_idx = selected[0]
        old_at = [float(ev.get("at", 0.0)) for ev in self.timeline]
        base = old_at[start_idx]

        if start_idx == 0:
            anchor = max(0.0, base + offset_sec)
        else:
            anchor = max(0.0, float(self.timeline[start_idx - 1].get("at", 0.0)) + offset_sec)

        for idx in range(start_idx, len(self.timeline)):
            delta = old_at[idx] - base
            self.timeline[idx]["at"] = round(max(0.0, anchor + delta), 2)

        self.mark_timeline_dirty()
        self.tree.selection_set(str(start_idx))
        self.set_status("已自第 {} 列起偏移 {:.3f} 秒，共 {} 列".format(start_idx, offset_sec, len(self.timeline) - start_idx))

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
            self.clear_runtime_highlight()
            self.refresh_tree()
            self.set_status("已停止 Pi：{}".format(self.config["pi_host"]))
        except Exception as e:
            self.set_frontend_error(str(e))
            messagebox.showerror("停止失敗", str(e))

    def send_timeline(self):
        if not self.timeline:
            messagebox.showwarning("提醒", "請先錄製並分析，或載入已保存項目")
            return

        self.config["pi_host"] = self.pi_ip_entry.get().strip() or DEFAULT_PI_HOST
        try:
            self.config["send_delay_sec"] = self.parse_send_delay_sec()
        except ValueError as e:
            self.set_frontend_error(str(e))
            messagebox.showwarning("提醒", str(e))
            return
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
            self.clear_runtime_highlight()
            self.refresh_tree()
            delay = self.apply_send_delay_if_needed()
            res = self.request_pi(payload, write_response=False)
            self.write_text({
                "sending_name": display_name,
                "pi_host": self.config["pi_host"],
                "send_delay_sec": delay,
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
        try:
            self.config["send_delay_sec"] = self.parse_send_delay_sec()
        except ValueError as e:
            self.set_frontend_error(str(e))
            messagebox.showwarning("提醒", str(e))
            return
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
            self.clear_runtime_highlight()
            self.refresh_tree()
            delay = self.apply_send_delay_if_needed()
            res = self.request_pi(payload, write_response=False)
            self.write_text({
                "sending_name": display_name,
                "pi_host": self.config["pi_host"],
                "send_delay_sec": delay,
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
            offset_sec = self.get_manual_offset_sec()
            self.push_history(action_reason)
            self.timeline = [self._sync_replicated_row(ev) for ev in self.timeline]
            self.timeline_meta["original_events"] = self.copy_events(self.timeline)
            self.validate_negative_group_monotonic_by_index(self.timeline)
            self.timeline = self.recalculate_timeline_for_runtime(anchor_gap_sec=offset_sec)
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
        region = self.tree.identify("region", event.x, event.y)
        col_id = self.tree.identify_column(event.x)
        if not col_id:
            return

        columns = ("idx", "type", "button", "at", "at_jitter", "buff_group", "buff_cycle_sec", "buff_jitter_sec", "group")
        col_index = int(col_id[1:]) - 1
        field = columns[col_index]
        editable_fields = ("type", "button", "at", "at_jitter", "buff_group", "buff_cycle_sec", "buff_jitter_sec")
        if field not in editable_fields:
            return

        if region == "heading":
            selected = sorted(self.get_selected_indexes())
            if not selected:
                messagebox.showwarning("提醒", "請先選取一列或多列，再雙擊欄位標題")
                return
            initial_value = str(self.timeline[selected[0]].get(field, ""))
            new_value = simpledialog.askstring(
                "批量修改欄位",
                "請輸入 {}（將套用到 {} 筆選取列）:".format(field, len(selected)),
                initialvalue=initial_value
            )
            if new_value is None:
                return

            try:
                for idx in selected:
                    self._apply_tree_field_value(idx, field, new_value)
            except Exception as e:
                messagebox.showerror("修改失敗", str(e))
                return

            self.mark_timeline_dirty()
            self.tree.selection_set([str(idx) for idx in selected])
            self.set_status("已將 {} 套用到 {} 筆選取列".format(field, len(selected)))
            return

        if region != "cell":
            return

        row_id = self.tree.identify_row(event.y)
        if not row_id:
            return

        idx = int(row_id)
        current_value = str(self.timeline[idx].get(field, ""))
        new_value = simpledialog.askstring("修改欄位", "請輸入 {}:".format(field), initialvalue=current_value)
        if new_value is None:
            return

        try:
            self._apply_tree_field_value(idx, field, new_value)
        except Exception as e:
            messagebox.showerror("修改失敗", str(e))
            return

        self.mark_timeline_dirty()
        self.tree.selection_set(str(idx))

    def _apply_tree_field_value(self, idx, field, raw_value):
        self._apply_tree_field_value_to_event(self.timeline[idx], field, raw_value)
        if field == "buff_group":
            self._sync_replicated_row(self.timeline[idx])

    def on_tree_select_all(self, _event=None):
        if not self.timeline:
            return "break"
        self.tree.selection_set([str(i) for i in range(len(self.timeline))])
        self.set_status("已全選 {} 列".format(len(self.timeline)))
        return "break"

    def on_tree_copy(self, _event=None):
        selected = sorted(self.get_selected_indexes())
        if not selected:
            return "break"

        lines = ["\t".join(self.tree_columns)]
        for idx in selected:
            values = self.tree.item(str(idx), "values")
            lines.append("\t".join(str(v) for v in values))

        self.root.clipboard_clear()
        self.root.clipboard_append("\n".join(lines))
        self.set_status("已複製 {} 列（含欄位標題，可貼到 Excel）".format(len(selected)))
        return "break"

    def on_tree_paste(self, _event=None):
        try:
            raw = self.root.clipboard_get()
        except tk.TclError:
            messagebox.showwarning("提醒", "剪貼簿沒有可貼上的內容")
            return "break"

        lines = [line for line in raw.splitlines() if line.strip()]
        if not lines:
            return "break"

        start_idx = 0
        selected = sorted(self.get_selected_indexes())
        if selected:
            start_idx = selected[0]

        fields = ("type", "button", "at", "at_jitter", "buff_group", "buff_cycle_sec", "buff_jitter_sec")
        full_header = [col.lower() for col in self.tree_columns]
        short_header = [f.lower() for f in fields]
        parsed_rows = []
        for line in lines:
            values = [v.strip() for v in line.split("\t")]
            lowered = [v.lower() for v in values]
            if values and (lowered == short_header or lowered == full_header):
                continue
            parsed_rows.append(values)

        if not parsed_rows:
            messagebox.showwarning("提醒", "剪貼簿只有標題列，沒有可貼上的資料")
            return "break"

        try:
            parsed_rows = [self._normalize_paste_row(values) for values in parsed_rows]
        except Exception as e:
            messagebox.showerror("貼上失敗", str(e))
            return "break"

        if len(parsed_rows) >= 2 and len(selected) == 1:
            anchor = selected[0]
            ask = messagebox.askyesnocancel(
                "貼上位置",
                "你選了第 {} 列。\n要貼在：\n"
                "Yes：第 {}~{} 列之間（插入在上方）\n"
                "No：第 {}~{} 列之間（插入在下方）\n"
                "Cancel：取消貼上".format(
                    anchor,
                    max(0, anchor - 1),
                    anchor,
                    anchor,
                    anchor + 1
                )
            )
            if ask is None:
                self.set_status("已取消貼上")
                return "break"

            insert_at = anchor if ask else anchor + 1
            new_indexes = []
            try:
                for offset, row_values in enumerate(parsed_rows):
                    ev = self._build_timeline_event_from_values(row_values)
                    idx = insert_at + offset
                    self.timeline.insert(idx, ev)
                    new_indexes.append(idx)
            except Exception as e:
                messagebox.showerror("貼上失敗", str(e))
                return "break"

            self.mark_timeline_dirty()
            self.tree.selection_set([str(i) for i in new_indexes])
            self.set_status("已插入貼上 {} 列（從 idx {} 開始）".format(len(new_indexes), insert_at))
            return "break"

        changed_indexes = []
        for offset, row_values in enumerate(parsed_rows):
            row_idx = start_idx + offset
            if row_idx >= len(self.timeline):
                break
            for col_idx, raw_value in enumerate(row_values):
                field = fields[col_idx]
                self._apply_tree_field_value(row_idx, field, raw_value)
            changed_indexes.append(row_idx)

        if not changed_indexes:
            messagebox.showwarning("提醒", "貼上範圍超出目前列數，未更新任何資料")
            return "break"

        self.mark_timeline_dirty()
        self.tree.selection_set([str(i) for i in changed_indexes])
        self.set_status("已貼上 {} 列（從 idx {} 開始）".format(len(changed_indexes), changed_indexes[0]))
        return "break"

    def _normalize_paste_row(self, values):
        if not values:
            raise ValueError("貼上資料有空白列")
        if len(values) < 7:
            raise ValueError("每列至少需要 7 欄（type 到 buff_jitter_sec）")
        # 優先支援 7 欄格式；若是完整表格 9 欄（idx + 7 欄 + group），則忽略 idx/group。
        if len(values) >= 9 and values[1].strip().lower() in ("press", "release"):
            normalized = values[1:8]
        else:
            normalized = values[:7]
        for i in range(7):
            normalized[i] = normalized[i].strip()
        return normalized

    def _build_timeline_event_from_values(self, values):
        event = {
            "type": "press",
            "button": "",
            "at": 0.0,
            "at_jitter": 0.0,
            "buff_group": "",
            "buff_cycle_sec": 0.0,
            "buff_jitter_sec": 0.0
        }
        fields = ("type", "button", "at", "at_jitter", "buff_group", "buff_cycle_sec", "buff_jitter_sec")
        for i, field in enumerate(fields):
            self._apply_tree_field_value_to_event(event, field, values[i])
        return self._sync_replicated_row(event)

    def _apply_tree_field_value_to_event(self, event, field, raw_value):
        if field in ("at", "at_jitter", "buff_cycle_sec", "buff_jitter_sec"):
            value = float(raw_value)
            if field == "at":
                value = max(0.0, value)
                event[field] = round(value, 2)
            else:
                value = abs(value)
                event[field] = round(value, 4)
            return

        value = raw_value.strip().lower()
        if field == "type" and value not in ("press", "release"):
            raise ValueError("type 只能是 press / release")
        if field == "button" and not value:
            raise ValueError("button 不可空白")
        if field == "buff_group":
            value = raw_value.strip()
        event[field] = value

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
            copied["replicatedRow"] = 1
            copied = self._sync_replicated_row(copied)
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
