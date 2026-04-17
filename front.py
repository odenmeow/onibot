# -*- coding: utf-8 -*-
import json
import os
import re
import socket
import time
import threading 
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, colorchooser
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
SUPPORTED_BUTTONS = {
    "fn", "g", "shift", "f", "c", "v", "d", "alt", "ctrl",
    "left", "up", "down", "right", "x", "space", "6"
}
DEFAULT_HINT_NOTE_TEXT = (
    "前端可接受按鍵(button)：fn, g, shift, f, c, v, d, alt, ctrl, left, up, down, right, x, space, 6。\n"
    "注意：多數電腦的實體 Fn 鍵無法直接被前端鍵盤監聽。\n"
    "建議先用可錄到的替代鍵（例如 Win/Cmd/Alt）錄製，再到 timeline 的 button 欄位手動改成 fn。\n"
    "補充：『套用偏移』是手動把選取列之後的時間整段平移；『糾正複製體』是依 buff_group 負值重算複製體群組的正確 at。"
)

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
            "hint_note_text": DEFAULT_HINT_NOTE_TEXT,
            "ui_recent_colors": [],
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
        ui_recent_colors = data.get("ui_recent_colors", [])
        if not isinstance(ui_recent_colors, list):
            ui_recent_colors = []
        normalized_recent = []
        for raw in ui_recent_colors:
            text = str(raw or "").strip().lower()
            if re.fullmatch(r"#[0-9a-f]{6}", text) and text not in normalized_recent:
                normalized_recent.append(text)
        normalized_recent = normalized_recent[:7]
        return {
            "pi_host": data.get("pi_host", DEFAULT_PI_HOST),
            "send_delay_sec": float(data.get("send_delay_sec", 1.0)),
            "last_selected_name": data.get("last_selected_name", ""),
            "buff_skip_mode": data.get("buff_skip_mode", BUFF_SKIP_MODE_COMPRESS),
            "manual_offset_sec": float(data.get("manual_offset_sec", NEGATIVE_GROUP_ANCHOR_GAP_SEC)),
            "hint_note_text": str(data.get("hint_note_text", DEFAULT_HINT_NOTE_TEXT)),
            "ui_recent_colors": normalized_recent,
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
            "hint_note_text": DEFAULT_HINT_NOTE_TEXT,
            "ui_recent_colors": [],
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


def fallback_button_name(key_name):
    if not key_name:
        return ""
    raw = str(key_name).strip()
    if len(raw) >= 2 and raw[0] == "'" and raw[-1] == "'":
        return raw[1:-1].lower()
    if raw.startswith("Key."):
        return raw[4:].lower()
    return raw.lower()


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


def rename_named_timeline(old_name, new_name, overwrite=False):
    old_path = timeline_file_path(old_name)
    new_path = timeline_file_path(new_name)

    if not os.path.exists(old_path):
        raise FileNotFoundError("找不到來源檔案：{}".format(old_name))
    if os.path.exists(new_path) and not overwrite:
        raise FileExistsError("名稱 '{}' 已存在".format(new_name))

    data = load_named_timeline(old_name)
    data["name"] = new_name
    with open(new_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    if old_path != new_path and os.path.exists(old_path):
        os.remove(old_path)


def on_press(key):
    global events
    if not recording:
        return

    key_name = normalize_key(key)
    mapped = KEY_MAP.get(key_name) or fallback_button_name(key_name)
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
    mapped = KEY_MAP.get(key_name) or fallback_button_name(key_name)
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
        row["buff_group"] = buff_group
        row["replicatedRow"] = replicated
        return row

    def _normalize_row_color(self, value):
        text = str(value or "").strip().lower()
        if re.fullmatch(r"#[0-9a-f]{6}", text):
            return text
        return ""

    def normalize_event_schema(self, ev):
        row = dict(ev)
        row.setdefault("type", "")
        row.setdefault("button", "")
        row["at"] = round(max(0.0, float(row.get("at", 0.0))), 2)
        row["at_jitter"] = round(abs(float(row.get("at_jitter", 0.0))), 4)
        row["buff_group"] = str(row.get("buff_group", "")).strip()
        row["buff_cycle_sec"] = round(max(0.0, float(row.get("buff_cycle_sec", 0.0))), 4)
        row["buff_jitter_sec"] = round(abs(float(row.get("buff_jitter_sec", 0.0))), 4)
        row["row_color"] = self._normalize_row_color(row.get("row_color", ""))
        replicated = self._normalize_replicated_row_flag(row.get("replicatedRow", 0))
        if replicated == 1 and not row["row_color"]:
            row["row_color"] = "#fff4b3"
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
            "latest_saved_events": []
        }

    def normalize_meta(self, meta):
        if not isinstance(meta, dict):
            return self.new_meta()
        raw_original = meta.get("original_events", [])
        raw_latest_saved = meta.get("latest_saved_events", [])
        raw_history = meta.get("history", [])
        original_events = []
        if isinstance(raw_original, list):
            original_events = [self.normalize_event_schema(ev) for ev in raw_original]

        latest_saved_events = []
        if isinstance(raw_latest_saved, list):
            latest_saved_events = [self.normalize_event_schema(ev) for ev in raw_latest_saved]
        elif isinstance(raw_history, list):
            # 舊格式相容：history 最後一筆視為新版快照。
            for item in reversed(raw_history):
                if not isinstance(item, dict):
                    continue
                events = item.get("events", [])
                if isinstance(events, list):
                    latest_saved_events = [self.normalize_event_schema(ev) for ev in events]
                    break
        return {
            "original_events": original_events,
            "latest_saved_events": latest_saved_events
        }

    def copy_events(self, rows):
        return [self.normalize_event_schema(ev) for ev in rows]

    def _history_key(self):
        name = str(getattr(self, "current_name", "") or "").strip()
        return name if name else "__unsaved__"

    def _history_bucket(self):
        if not hasattr(self, "timeline_histories") or not isinstance(self.timeline_histories, dict):
            self.timeline_histories = {}
        key = self._history_key()
        if key not in self.timeline_histories:
            self.timeline_histories[key] = {"undo": [], "redo": []}
        return self.timeline_histories[key]

    def _timeline_signature(self, rows):
        return json.dumps(rows, ensure_ascii=False, sort_keys=True)

    def _begin_timeline_change(self):
        return self.copy_events(self.timeline)

    def _finalize_timeline_change(self, before_snapshot):
        if self._timeline_signature(before_snapshot) == self._timeline_signature(self.timeline):
            return False
        bucket = self._history_bucket()
        bucket["undo"].append(before_snapshot)
        if len(bucket["undo"]) > 100:
            bucket["undo"] = bucket["undo"][-100:]
        bucket["redo"].clear()
        return True

    def _reset_timeline_history(self):
        if not hasattr(self, "timeline_histories") or not isinstance(self.timeline_histories, dict):
            self.timeline_histories = {}
        self.timeline_histories[self._history_key()] = {"undo": [], "redo": []}

    def undo_timeline(self, _event=None):
        if not self._ensure_runtime_editable():
            return "break"
        bucket = self._history_bucket()
        if not bucket["undo"]:
            self.set_status("沒有可復原的上一步")
            return "break"
        before = bucket["undo"].pop()
        bucket["redo"].append(self.copy_events(self.timeline))
        self.timeline = self.copy_events(before)
        self.mark_timeline_dirty()
        self.set_status("已復原上一步")
        return "break"

    def redo_timeline(self, _event=None):
        if not self._ensure_runtime_editable():
            return "break"
        bucket = self._history_bucket()
        if not bucket["redo"]:
            self.set_status("沒有可重做的下一步")
            return "break"
        after = bucket["redo"].pop()
        bucket["undo"].append(self.copy_events(self.timeline))
        self.timeline = self.copy_events(after)
        self.mark_timeline_dirty()
        self.set_status("已重做下一步")
        return "break"

    def restore_original_timeline(self):
        original = self.timeline_meta.get("original_events", [])
        if not original:
            return False
        self.timeline = self.copy_events(original)
        self._reset_timeline_history()
        self.current_loaded_from_saved = False
        self.update_current_labels()
        self.refresh_tree()
        self.refresh_preview()
        return True

    def restore_latest_saved_timeline(self):
        latest_saved = self.timeline_meta.get("latest_saved_events", [])
        if not latest_saved:
            return False
        self.timeline = self.copy_events(latest_saved)
        self._reset_timeline_history()
        self.current_loaded_from_saved = False
        self.update_current_labels()
        self.refresh_tree()
        self.refresh_preview()
        return True

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

    def get_unsupported_buttons(self, events):
        unknown = sorted({
            str(ev.get("button", "")).strip().lower()
            for ev in events
            if str(ev.get("button", "")).strip()
            and str(ev.get("button", "")).strip().lower() not in SUPPORTED_BUTTONS
        })
        return unknown

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
        self.pre_run_timeline_snapshot = []
        self.runtime_working_timeline = []
        self.has_pre_run_snapshot = False
        self.runtime_display_frozen = False
        self.runtime_manual_restore_active = False
        self.tree_row_color_tags = set()
        self.last_tree_copy_payload = ""
        self.timeline_histories = {}

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
        tk.Button(save_btn_row, text="保存目前 Timeline", command=self.save_current_timeline).pack(side="left", padx=(0, 3))
        tk.Button(save_btn_row, text="重新整理", command=self.refresh_saved_list).pack(side="left", padx=3)
        tk.Button(save_btn_row, text="重新命名", command=self.rename_selected_timeline).pack(side="left", padx=3)
        tk.Button(save_btn_row, text="載入選取項目", command=self.load_selected_timeline).pack(side="left", padx=3)
        tk.Button(save_btn_row, text="刪除選取項目", command=self.delete_selected_timeline).pack(side="left", padx=3)

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
        tk.Button(
            jitter_frame,
            text="提示",
            command=self.open_hint_note_dialog,
            bg="#fff4b3",
            activebackground="#ffe08a"
        ).grid(row=0, column=2, padx=(0, 5), pady=5)
        tk.Button(jitter_frame, text="糾正複製體", command=self.calculate_offsets_only).grid(
            row=0, column=3, padx=(5, 8), pady=5
        )
        tk.Button(jitter_frame, text="改顏色", command=self.open_row_color_dialog).grid(
            row=0, column=4, padx=(0, 8), pady=5
        )
        self.restore_pre_run_btn = tk.Button(
            jitter_frame,
            text="恢復執行前狀態",
            command=self.restore_pre_run_state,
            state="disabled"
        )
        self.restore_pre_run_btn.grid(row=0, column=5, padx=(0, 8), pady=5)
        tk.Label(
            jitter_frame,
            text="【註：buff_group 為負值時，請用「糾正複製體」重算；「套用偏移」僅手動平移時間。】"
        ).grid(row=1, column=0, columnspan=6, padx=(8, 8), pady=(0, 6), sticky="w")

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
        self.tree.tag_configure("runtime_ok_1", background="#bfe8bf")
        self.tree.tag_configure("runtime_ok_2", background="#d9f2d9")
        self.tree.tag_configure("runtime_ok_3", background="#edf9ed")
        self.tree.tag_configure("unsupported_button", background="#ffcaca")
        self.tree.bind("<Double-1>", self.on_tree_double_click)
        self.tree.bind("<Control-a>", self.on_tree_select_all, add="+")
        self.tree.bind("<Control-A>", self.on_tree_select_all, add="+")
        self.tree.bind("<Control-c>", self.on_tree_copy, add="+")
        self.tree.bind("<Control-C>", self.on_tree_copy, add="+")
        self.tree.bind("<Control-v>", self.on_tree_paste, add="+")
        self.tree.bind("<Control-V>", self.on_tree_paste, add="+")
        self.tree.bind("<Control-z>", self.undo_timeline, add="+")
        self.tree.bind("<Control-Z>", self.undo_timeline, add="+")
        self.tree.bind("<Control-y>", self.redo_timeline, add="+")
        self.tree.bind("<Control-Y>", self.redo_timeline, add="+")

        edit_row = tk.Frame(right_panel)
        edit_row.pack(fill="x", pady=(0, 8))
        tk.Button(edit_row, text="上一步", command=self.undo_timeline, width=9).pack(side="left", padx=2)
        tk.Button(edit_row, text="下一步", command=self.redo_timeline, width=9).pack(side="left", padx=2)
        tk.Button(edit_row, text="上移", command=self.move_selected_up, width=9).pack(side="left", padx=2)
        tk.Button(edit_row, text="下移", command=self.move_selected_down, width=9).pack(side="left", padx=2)
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
        self.config["ui_recent_colors"] = self._get_recent_colors()
        save_config(self.config)
        self.set_status("已保存 UI 版面（PanedWindow 像素位置 + 欄寬）")

    def get_hint_note_text(self):
        text = str(self.config.get("hint_note_text", "")).strip()
        return text or DEFAULT_HINT_NOTE_TEXT

    def open_hint_note_dialog(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("提示說明")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.geometry("700x420")

        tk.Label(dialog, text="此內容會與 UI 版面一起保存到 front_config.json", anchor="w").pack(
            fill="x", padx=12, pady=(10, 4)
        )
        text_widget = tk.Text(dialog, wrap="word")
        text_widget.pack(fill="both", expand=True, padx=12, pady=(0, 8))
        text_widget.insert("1.0", self.get_hint_note_text())

        btn_row = tk.Frame(dialog)
        btn_row.pack(fill="x", padx=12, pady=(0, 12))

        def save_note():
            note = text_widget.get("1.0", tk.END).strip() or DEFAULT_HINT_NOTE_TEXT
            self.config["hint_note_text"] = note
            save_config(self.config)
            self.write_text({"action": "hint_note_saved", "hint_note_text": note})
            self.set_status("提示說明已保存（與 UI 設定同檔）")
            dialog.destroy()

        def reset_default():
            text_widget.delete("1.0", tk.END)
            text_widget.insert("1.0", DEFAULT_HINT_NOTE_TEXT)

        tk.Button(btn_row, text="恢復預設", command=reset_default).pack(side="left")
        tk.Button(btn_row, text="取消", command=dialog.destroy).pack(side="right", padx=(6, 0))
        tk.Button(
            btn_row,
            text="保存",
            bg="#fff4b3",
            activebackground="#ffe08a",
            command=save_note
        ).pack(side="right")

    def _get_recent_colors(self):
        raw = self.config.get("ui_recent_colors", [])
        if not isinstance(raw, list):
            return []
        normalized = []
        for color in raw:
            code = self._normalize_row_color(color)
            if code and code not in normalized:
                normalized.append(code)
            if len(normalized) >= 7:
                break
        return normalized

    def _remember_recent_color(self, color_code):
        color = self._normalize_row_color(color_code)
        if not color:
            return
        recent = [c for c in self._get_recent_colors() if c != color]
        recent.insert(0, color)
        self.config["ui_recent_colors"] = recent[:7]
        save_config(self.config)

    def _apply_color_to_selected_rows(self, color_code):
        if not self._ensure_runtime_editable():
            return False
        color = self._normalize_row_color(color_code)
        selected = sorted(self.get_selected_indexes())
        if not selected:
            messagebox.showwarning("提醒", "請先選取列")
            return False
        if not color:
            messagebox.showwarning("提醒", "請選擇有效顏色")
            return False
        before = self._begin_timeline_change()
        for idx in selected:
            self.timeline[idx]["row_color"] = color
        self._finalize_timeline_change(before)
        self.mark_timeline_dirty()
        self.tree.selection_set([str(i) for i in selected])
        self._remember_recent_color(color)
        self.set_status("已套用顏色 {} 到 {} 列".format(color, len(selected)))
        return True

    def open_row_color_dialog(self):
        selected = sorted(self.get_selected_indexes())
        if not selected:
            messagebox.showwarning("提醒", "請先選取列")
            return
        if not self._ensure_runtime_editable():
            return

        dialog = tk.Toplevel(self.root)
        dialog.title("改顏色")
        dialog.transient(self.root)
        dialog.grab_set()
        tk.Label(dialog, text="套用到目前選取列（含 Shift 多選）").pack(anchor="w", padx=12, pady=(10, 8))

        recent_frame = tk.Frame(dialog)
        recent_frame.pack(fill="x", padx=12, pady=(0, 8))
        tk.Label(recent_frame, text="Recent：").pack(side="left")
        recent = self._get_recent_colors()
        if recent:
            for color in recent:
                tk.Button(
                    recent_frame,
                    text=color,
                    bg=color,
                    width=10,
                    command=lambda c=color: (
                        self._apply_color_to_selected_rows(c) and dialog.destroy()
                    )
                ).pack(side="left", padx=2)
        else:
            tk.Label(recent_frame, text="（尚無）").pack(side="left", padx=4)

        action_row = tk.Frame(dialog)
        action_row.pack(fill="x", padx=12, pady=(0, 12))

        def pick_from_system():
            initial = self.timeline[selected[0]].get("row_color", "") if selected else ""
            _rgb, code = colorchooser.askcolor(color=initial or "#fff4b3", parent=dialog, title="選擇列顏色")
            if not code:
                return
            if self._apply_color_to_selected_rows(code):
                dialog.destroy()

        tk.Button(action_row, text="系統選色器", command=pick_from_system).pack(side="left")
        tk.Button(action_row, text="取消", command=dialog.destroy).pack(side="right")

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
                if isinstance(res, dict):
                    changed = self.update_runtime_from_status(res)
                    if changed:
                        state = str(self.timeline_runtime_info.get("state", "")).strip().lower()
                        if state == "running":
                            self.runtime_manual_restore_active = False
                            self.runtime_display_frozen = False
                            self.refresh_tree()
                            self.focus_latest_runtime_row()
                        elif state in ("stopped", "idle"):
                            if self.has_pre_run_snapshot and not self.runtime_manual_restore_active:
                                self.runtime_display_frozen = True
                            else:
                                self.runtime_display_frozen = False
                        else:
                            self.runtime_display_frozen = False
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

    def _ensure_tree_color_tag(self, color_code):
        color = self._normalize_row_color(color_code)
        if not color:
            return ""
        tag = "row_color_{}".format(color.replace("#", ""))
        if tag not in self.tree_row_color_tags:
            self.tree.tag_configure(tag, background=color)
            self.tree_row_color_tags.add(tag)
        return tag

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
            runtime_event = self.timeline_runtime_by_index.get(i, {})
            runtime_status = str(runtime_event.get("status", "")).strip().lower()
            if runtime_status == "skipped_by_cooldown":
                buff_group = "{}（冷卻中）".format(original_buff_group) if original_buff_group else "冷卻中"
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
            else:
                row_color_tag = self._ensure_tree_color_tag(ev.get("row_color", ""))
                if row_color_tag:
                    tags.append(row_color_tag)
            if str(ev.get("button", "")).strip().lower() not in SUPPORTED_BUTTONS:
                tags.append("unsupported_button")
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

    def _is_runtime_readonly(self):
        state = str(self.timeline_runtime_info.get("state", "")).strip().lower()
        return state == "running" or bool(self.runtime_display_frozen)

    def _ensure_runtime_editable(self):
        if not self.has_pre_run_snapshot and self.runtime_display_frozen:
            self.runtime_display_frozen = False
        if not self._is_runtime_readonly():
            return True
        messagebox.showwarning("提醒", "執行中或停止後凍結中，請先「恢復執行前狀態」再編輯")
        return False

    def _set_restore_pre_run_button_state(self):
        btn = getattr(self, "restore_pre_run_btn", None)
        if btn is None:
            return
        btn.config(state=("normal" if self.has_pre_run_snapshot else "disabled"))

    def restore_pre_run_state(self):
        if not self.has_pre_run_snapshot:
            messagebox.showwarning("提醒", "沒有可恢復內容")
            return
        self.timeline = self.copy_events(self.pre_run_timeline_snapshot)
        self._reset_timeline_history()
        self.runtime_manual_restore_active = True
        self.runtime_display_frozen = False
        self.clear_runtime_highlight()
        self.refresh_tree()
        self.refresh_preview()

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
        self._reset_timeline_history()
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
            has_original = bool(self.timeline_meta.get("original_events"))
            has_latest_saved = bool(self.timeline_meta.get("latest_saved_events"))
            if has_original and has_latest_saved:
                selected = messagebox.askyesnocancel(
                    "重新分析",
                    "目前無新錄製資料。\n是否回到「初版 Timeline」？\n"
                    "按「是」= 初版；按「否」= 上次保存新版；按「取消」= 不變更。"
                )
                if selected is None:
                    self.set_status("已取消重新分析")
                    return
                restored = self.restore_original_timeline() if selected else self.restore_latest_saved_timeline()
                if restored:
                    if selected:
                        self.set_status("已還原初版 Timeline（重新分析）")
                    else:
                        self.set_status("已還原上次保存新版 Timeline（重新分析）")
                return
            if has_original:
                messagebox.showinfo("重新分析", "目前無新錄製資料。\n目前僅有可回復的「初版 Timeline」。")
                self.restore_original_timeline()
                self.set_status("已還原初版 Timeline（重新分析）")
                return
            if has_latest_saved:
                messagebox.showinfo("重新分析", "目前無新錄製資料。\n目前僅有可回復的「上次保存新版 Timeline」。")
                self.restore_latest_saved_timeline()
                self.set_status("已還原上次保存新版 Timeline（重新分析）")
                return
            messagebox.showwarning("提醒", "目前沒有錄到資料，也沒有可還原的初版/新版 Timeline")
            return

        self.timeline = build_timeline(events, recording_start)
        self._reset_timeline_history()
        self.timeline_meta = self.new_meta()
        self.mark_timeline_dirty()

        if not self.current_name:
            self.current_name = "未命名_{}".format(time.strftime("%H%M%S"))

        self.name_entry.delete(0, tk.END)
        self.name_entry.insert(0, self.current_name)

        unsupported = self.get_unsupported_buttons(self.timeline)
        if unsupported:
            self.set_frontend_error(
                "錄製完成，但包含前端可錄製、後端未支援的按鍵：{}\n"
                "這些列已在 timeline 以紅底顯示；請改成可用 button 後再送出。".format(", ".join(unsupported))
            )
        else:
            self.set_frontend_error("")
        self.set_status("Timeline 分析完成，共 {} 筆 event".format(len(self.timeline)))

    def ask_save_existing_timeline_mode(self, name):
        dialog = tk.Toplevel(self.root)
        dialog.title("保存同名 Timeline")
        dialog.transient(self.root)
        dialog.resizable(False, False)
        dialog.grab_set()

        result = {"mode": None}

        msg = (
            "名稱 '{}' 已存在。\n\n"
            "請選擇保存方式：\n"
            "・存為新版：保留初版 original_events，只更新 latest_saved_events。\n"
            "・完全取代：重建初版與新版，會覆蓋整份版本資訊。"
        ).format(name)
        tk.Label(dialog, text=msg, justify="left", anchor="w").pack(padx=18, pady=(16, 8))

        btn_row = tk.Frame(dialog)
        btn_row.pack(padx=12, pady=(8, 14), fill="x")

        def choose(mode):
            result["mode"] = mode
            dialog.destroy()

        btn_new = tk.Button(btn_row, text="存為新版", width=12, command=lambda: choose("new_version"))
        btn_new.pack(side="left", padx=5)
        tk.Button(btn_row, text="完全取代", width=12, command=lambda: choose("full_replace")).pack(side="left", padx=5)
        tk.Button(btn_row, text="取消", width=10, command=lambda: choose(None)).pack(side="left", padx=5)

        dialog.protocol("WM_DELETE_WINDOW", lambda: choose(None))
        btn_new.focus_set()
        dialog.bind("<Return>", lambda _e: choose("new_version"))
        dialog.bind("<Escape>", lambda _e: choose(None))
        self.root.wait_window(dialog)
        return result["mode"]

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
        save_mode = "new_version"

        if os.path.exists(path):
            save_mode = self.ask_save_existing_timeline_mode(name)
            if save_mode is None:
                self.set_status("已取消保存")
                return

        try:
            offset_sec = self.get_manual_offset_sec()
            working_timeline = self.copy_events(self.timeline)
            self.validate_negative_group_monotonic_by_index(working_timeline)
            recalculated = recalculate_runtime_events_by_index(self.copy_events(working_timeline), offset_sec)
            if save_mode == "full_replace":
                self.timeline_meta["original_events"] = self.copy_events(working_timeline)
            elif not self.timeline_meta.get("original_events"):
                self.timeline_meta["original_events"] = self.copy_events(working_timeline)
            self.timeline_meta["latest_saved_events"] = self.copy_events(recalculated)
        except Exception as e:
            messagebox.showerror("保存失敗", str(e))
            return

        save_payload = self.copy_events(recalculated)
        if save_mode == "full_replace":
            self.timeline = self.copy_events(recalculated)
            save_payload = self.copy_events(self.timeline)
        save_named_timeline(name, save_payload, self.timeline_meta)
        self.current_name = name
        self.current_loaded_from_saved = True
        self.config["last_selected_name"] = name
        save_config(self.config)

        self.refresh_saved_list()
        self.select_saved_name(name)
        self.mark_timeline_dirty()
        if save_mode == "full_replace":
            self.set_status("已完全取代保存 Timeline：{}".format(name))
        else:
            self.set_status("已保存新版 Timeline（保留初版）：{}".format(name))

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
        if not self.timeline_meta.get("original_events") and self.timeline:
            self.timeline_meta["original_events"] = self.copy_events(self.timeline)
        self.current_name = data.get("name", name)
        self._reset_timeline_history()
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

    def rename_selected_timeline(self):
        old_name = self.get_selected_saved_name()
        if not old_name:
            messagebox.showwarning("提醒", "請先從清單選一個已保存項目")
            return

        new_name = sanitize_filename(self.name_entry.get().strip())
        if not new_name:
            messagebox.showwarning("提醒", "請先輸入新的腳本名稱")
            return

        if new_name == old_name:
            self.set_status("名稱未變更：{}".format(old_name))
            return

        overwrite = False
        if os.path.exists(timeline_file_path(new_name)):
            overwrite = messagebox.askyesno("確認覆蓋", "名稱 '{}' 已存在，是否覆蓋？".format(new_name))
            if not overwrite:
                return

        try:
            rename_named_timeline(old_name, new_name, overwrite=overwrite)
        except Exception as e:
            messagebox.showerror("重新命名失敗", str(e))
            return

        if self.current_name == old_name:
            self.current_name = new_name
            self.current_loaded_from_saved = True
        if self.config.get("last_selected_name") == old_name:
            self.config["last_selected_name"] = new_name
            save_config(self.config)

        self.name_entry.delete(0, tk.END)
        self.name_entry.insert(0, new_name)
        self.refresh_saved_list()
        self.select_saved_name(new_name)
        self.update_current_labels()
        self.set_status("已重新命名：{} → {}".format(old_name, new_name))

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

        before = self._begin_timeline_change()
        for idx in selected:
            self.timeline[idx]["at_jitter"] = jitter

        self._finalize_timeline_change(before)
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

        before = self._begin_timeline_change()
        for ev in self.timeline:
            ev["at_jitter"] = jitter

        self._finalize_timeline_change(before)
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

        before = self._begin_timeline_change()
        for idx in selected:
            self.timeline[idx]["at_jitter"] = 0.0

        self._finalize_timeline_change(before)
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
        for ev in prepared_events:
            if self._is_negative_buff_group(ev.get("buff_group", "")):
                ev["buff_group"] = ""
        before = self._begin_timeline_change()
        self.runtime_working_timeline = self.copy_events(prepared_events)
        self.timeline = self.copy_events(prepared_events)
        self._finalize_timeline_change(before)
        self.current_loaded_from_saved = False
        self.update_current_labels()
        self.refresh_tree()
        self.refresh_preview()
        self.set_status("已糾正複製體並更新 table（尚未保存，共 {} 筆）".format(len(prepared_events)))

    def apply_offset_from_selected(self):
        if not self._ensure_runtime_editable():
            return
        if not self.timeline:
            messagebox.showwarning("提醒", "目前沒有 timeline 資料")
            return

        selected = sorted(self.get_selected_indexes())
        if not selected:
            messagebox.showwarning("提醒", "請先選取一列或多列")
            return
        if any(self._is_negative_buff_group(self.timeline[idx].get("buff_group", "")) for idx in selected):
            messagebox.showwarning("提醒", "選到複製體負群（buff_group < 0），請改用「糾正複製體」。")
            return

        try:
            offset_sec = self.get_manual_offset_sec()
        except Exception as e:
            messagebox.showerror("錯誤", str(e))
            return

        before = self._begin_timeline_change()
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

        self._finalize_timeline_change(before)
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
            self.runtime_manual_restore_active = False
            self.runtime_display_frozen = bool(self.has_pre_run_snapshot)
            self._set_restore_pre_run_button_state()
            self.set_status("已停止 Pi：{}".format(self.config["pi_host"]))
        except Exception as e:
            self.set_frontend_error(str(e))
            messagebox.showerror("停止失敗", str(e))

    def send_timeline(self):
        if not self.timeline:
            messagebox.showwarning("提醒", "請先錄製並分析，或載入已保存項目")
            return
        self.pre_run_timeline_snapshot = self.copy_events(self.timeline)
        self.has_pre_run_snapshot = True
        self.runtime_manual_restore_active = False
        self.runtime_display_frozen = False
        self._set_restore_pre_run_button_state()

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
        self.pre_run_timeline_snapshot = self.copy_events(self.timeline)
        self.has_pre_run_snapshot = True
        self.runtime_manual_restore_active = False
        self.runtime_display_frozen = False
        self._set_restore_pre_run_button_state()

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
            base_events = [self.normalize_event_schema(ev) for ev in self.timeline]
            self.validate_negative_group_monotonic_by_index(base_events)
            events = recalculate_runtime_events_by_index(base_events, offset_sec)
        except Exception as e:
            messagebox.showerror("時間重算失敗", str(e))
            return None, ""
        unsupported = self.get_unsupported_buttons(events)
        if unsupported:
            msg = (
                "前端偵測到不支援的 button：{}\n"
                "已用紅底標記對應列。請先修改 button 欄位後再送出。".format(", ".join(unsupported))
            )
            self.set_frontend_error(msg)
            messagebox.showwarning("送出前檢查", msg)
            self.refresh_tree()
            return None, ""
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
        if not self._ensure_runtime_editable():
            return
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

            before = self._begin_timeline_change()
            try:
                for idx in selected:
                    self._apply_tree_field_value(idx, field, new_value)
            except Exception as e:
                messagebox.showerror("修改失敗", str(e))
                return

            self._finalize_timeline_change(before)
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

        before = self._begin_timeline_change()
        try:
            self._apply_tree_field_value(idx, field, new_value)
        except Exception as e:
            messagebox.showerror("修改失敗", str(e))
            return

        self._finalize_timeline_change(before)
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

        payload = "\n".join(lines)
        self.root.clipboard_clear()
        self.root.clipboard_append(payload)
        self.last_tree_copy_payload = payload
        self.set_status("已複製 {} 列（含欄位標題，可貼到 Excel）".format(len(selected)))
        return "break"

    def _allocate_auto_negative_group(self):
        used = set()
        for ev in self.timeline:
            group = str(ev.get("buff_group", "")).strip()
            if not group.startswith("-"):
                continue
            try:
                used.add(int(group))
            except ValueError:
                continue

        candidate = -1
        while candidate in used:
            candidate -= 1
        return str(candidate)

    def on_tree_paste(self, _event=None):
        if not self._ensure_runtime_editable():
            return "break"
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
        copied_from_tree = raw.strip() == self.last_tree_copy_payload.strip() and bool(self.last_tree_copy_payload.strip())
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

        if copied_from_tree:
            auto_group = self._allocate_auto_negative_group()
            for row_values in parsed_rows:
                row_values[4] = auto_group

        if len(selected) == 1:
            anchor = selected[0]
            position = self.ask_paste_position(anchor, len(parsed_rows))
            if position is None:
                self.set_status("已取消貼上")
                return "break"
            insert_at = anchor if position == "above" else anchor + 1
            new_indexes = []
            before = self._begin_timeline_change()
            try:
                for offset, row_values in enumerate(parsed_rows):
                    ev = self._build_timeline_event_from_values(row_values)
                    ev["row_color"] = "#fff4b3"
                    idx = insert_at + offset
                    self.timeline.insert(idx, ev)
                    new_indexes.append(idx)
            except Exception as e:
                messagebox.showerror("貼上失敗", str(e))
                return "break"

            self._finalize_timeline_change(before)
            self.mark_timeline_dirty()
            self.tree.selection_set([str(i) for i in new_indexes])
            self.set_status("已插入貼上 {} 列（從 idx {} 開始）".format(len(new_indexes), insert_at))
            return "break"

        before = self._begin_timeline_change()
        changed_indexes = []
        for offset, row_values in enumerate(parsed_rows):
            row_idx = start_idx + offset
            if row_idx >= len(self.timeline):
                break
            for col_idx, raw_value in enumerate(row_values):
                field = fields[col_idx]
                self._apply_tree_field_value(row_idx, field, raw_value)
            self.timeline[row_idx]["row_color"] = "#fff4b3"
            changed_indexes.append(row_idx)

        if not changed_indexes:
            messagebox.showwarning("提醒", "貼上範圍超出目前列數，未更新任何資料")
            return "break"

        self._finalize_timeline_change(before)
        self.mark_timeline_dirty()
        self.tree.selection_set([str(i) for i in changed_indexes])
        self.set_status("已貼上 {} 列（從 idx {} 開始）".format(len(changed_indexes), changed_indexes[0]))
        return "break"

    def ask_paste_position(self, anchor_idx, row_count):
        choice = {"value": None}
        dialog = tk.Toplevel(self.root)
        dialog.title("貼上位置")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)

        msg = (
            "你選了第 {} 列，要把 {} 列資料貼在目標的哪裡？"
            .format(anchor_idx, row_count)
        )
        tk.Label(dialog, text=msg, justify="left", anchor="w").pack(
            fill="x", padx=14, pady=(12, 8)
        )

        btn_row = tk.Frame(dialog)
        btn_row.pack(fill="x", padx=14, pady=(0, 12))

        def choose(value):
            choice["value"] = value
            dialog.destroy()

        tk.Button(btn_row, text="上方", width=10, command=lambda: choose("above")).pack(side="left")
        tk.Button(btn_row, text="下方", width=10, command=lambda: choose("below")).pack(side="left", padx=(6, 0))
        tk.Button(btn_row, text="取消", width=10, command=dialog.destroy).pack(side="right")

        dialog.update_idletasks()
        dialog_w = dialog.winfo_width()
        dialog_h = dialog.winfo_height()

        try:
            tree = self.tree
            tree_x = tree.winfo_rootx()
            tree_y = tree.winfo_rooty()
            tree_w = tree.winfo_width()
            tree_h = tree.winfo_height()
        except Exception:
            tree_x = self.root.winfo_rootx()
            tree_y = self.root.winfo_rooty()
            tree_w = self.root.winfo_width()
            tree_h = self.root.winfo_height()

        x = int(tree_x + (tree_w - dialog_w) / 2)
        y = int(tree_y + (tree_h - dialog_h) / 2 + 45)
        screen_w = dialog.winfo_screenwidth()
        screen_h = dialog.winfo_screenheight()
        x = max(0, min(x, screen_w - dialog_w))
        y = max(0, min(y, screen_h - dialog_h))
        dialog.geometry("{}x{}+{}+{}".format(dialog_w, dialog_h, x, y))

        self.root.wait_window(dialog)
        return choice["value"]

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
            "buff_jitter_sec": 0.0,
            "row_color": ""
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
        if not self._ensure_runtime_editable():
            return
        selected = sorted(self.get_selected_indexes())
        if not selected:
            messagebox.showwarning("提醒", "請先選取列")
            return
        if selected[0] == 0:
            return
        before = self._begin_timeline_change()
        for idx in selected:
            self.timeline[idx - 1], self.timeline[idx] = self.timeline[idx], self.timeline[idx - 1]
        self._finalize_timeline_change(before)
        self.mark_timeline_dirty()
        self.tree.selection_set([str(i - 1) for i in selected])

    def move_selected_down(self):
        if not self._ensure_runtime_editable():
            return
        selected = sorted(self.get_selected_indexes(), reverse=True)
        if not selected:
            messagebox.showwarning("提醒", "請先選取列")
            return
        if selected[0] == len(self.timeline) - 1:
            return
        before = self._begin_timeline_change()
        for idx in selected:
            self.timeline[idx + 1], self.timeline[idx] = self.timeline[idx], self.timeline[idx + 1]
        self._finalize_timeline_change(before)
        self.mark_timeline_dirty()
        self.tree.selection_set([str(i + 1) for i in sorted(selected)])

    def delete_selected_rows(self):
        if not self._ensure_runtime_editable():
            return
        selected = sorted(self.get_selected_indexes(), reverse=True)
        if not selected:
            messagebox.showwarning("提醒", "請先選取列")
            return
        before = self._begin_timeline_change()
        for idx in selected:
            self.timeline.pop(idx)
        self._finalize_timeline_change(before)
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
