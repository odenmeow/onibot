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
            "last_selected_name": ""
        }

    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        return {
            "pi_host": data.get("pi_host", DEFAULT_PI_HOST),
            "last_selected_name": data.get("last_selected_name", "")
        }
    except Exception:
        return {
            "pi_host": DEFAULT_PI_HOST,
            "last_selected_name": ""
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


def save_named_timeline(name, timeline):
    path = timeline_file_path(name)
    data = {
        "name": name,
        "saved_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        "events": timeline
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
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.settimeout(timeout)
    s.connect((pi_host, PI_PORT))
    s.sendall(json.dumps(payload, ensure_ascii=False).encode("utf-8"))
    data = s.recv(1024 * 1024)
    s.close()
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
            "at_jitter": 0.0
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

    def request_pi(self, payload, success_status=None, write_response=True):
        try:
            res = send_to_pi(payload, self.config["pi_host"])
        except Exception as e:
            err = str(e)
            self.set_frontend_error(err)
            if success_status:
                self.set_status(success_status + "（前端連線失敗）")
            raise

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

    def poll_pi_status(self):
        if self.polling_status:
            self.root.after(1000, self.poll_pi_status)
            return

        self.polling_status = True
        pi_host = self.pi_ip_entry.get().strip() or DEFAULT_PI_HOST

        def worker():
            frontend_err = ""
            backend_err = None
            status_msg = None
            try:
                res = send_to_pi({"action": "status"}, pi_host, timeout=0.8)

                run_status = res.get("run_status", {})
                state = run_status.get("state", "unknown")
                mode = run_status.get("mode", "")
                message = run_status.get("message", "")

                if state == "running":
                    status_msg = "Pi 執行中：{} / {}".format(mode, message)
                    backend_err = ""
                elif state == "stopped":
                    status_msg = "Pi 已停止：{}".format(message)
                    backend_err = message
                elif state == "error":
                    status_msg = "Pi 錯誤：{}".format(message)
                    backend_err = message
                else:
                    backend_err = ""
            except Exception as e:
                frontend_err = str(e)

            def apply_result():
                if status_msg:
                    self.set_status(status_msg)
                if backend_err is not None:
                    self.set_backend_error(backend_err)
                if frontend_err:
                    self.set_frontend_error(frontend_err)
                self.polling_status = False
                self.root.after(1000, self.poll_pi_status)

            self.root.after(0, apply_result)

        threading.Thread(target=worker, daemon=True).start()

    def __init__(self, root):
        ensure_dirs()
        self.config = load_config()

        self.root = root
        self.root.title("鍵盤錄製 / Timeline 分析 / 傳送")
        self.timeline = []
        self.current_name = ""
        self.current_loaded_from_saved = False
        self.polling_status = False

        container = tk.Frame(root)
        container.pack(fill="both", expand=True, padx=10, pady=8)

        body = tk.PanedWindow(container, orient=tk.HORIZONTAL, sashrelief=tk.RAISED)
        body.pack(fill="both", expand=True)

        left_panel = tk.Frame(body)
        right_panel = tk.Frame(body)
        body.add(left_panel, minsize=560)
        body.add(right_panel, minsize=420)

        top = tk.LabelFrame(left_panel, text="操作區")
        top.pack(fill="x", pady=(0, 6))

        self.status_var = tk.StringVar(value="尚未錄製")
        tk.Label(top, textvariable=self.status_var, anchor="w").grid(
            row=0, column=0, columnspan=6, sticky="we", padx=8, pady=(6, 2)
        )

        btn_specs = [
            ("開始錄製", self.start_record, None),
            ("停止錄製", self.stop_record, None),
            ("分析 Timeline", self.analyze, None),
            ("送出執行", self.send_timeline, "#c8f7c5"),
            ("停止", self.stop_pi, "#ff8c69"),
            ("測試連線", self.ping_pi, "#d9d9d9"),
        ]
        for idx, (txt, cmd, color) in enumerate(btn_specs):
            kwargs = {"text": txt, "command": cmd, "width": 10}
            if color:
                kwargs["bg"] = color
            tk.Button(top, **kwargs).grid(row=1, column=idx, padx=4, pady=(2, 8))

        self.current_script_var = tk.StringVar(value="【目前腳本：未命名 / 未儲存】")
        tk.Label(
            top,
            textvariable=self.current_script_var,
            anchor="w",
            fg="#1a4fb8"
        ).grid(row=2, column=0, columnspan=6, sticky="w", padx=8, pady=(0, 8))

        info = tk.LabelFrame(left_panel, text="目前套用資訊")
        row = tk.Frame(info)
        row.pack(fill="x", padx=8, pady=2)

        tk.Label(row, text="Pi IP：").pack(side="left")
        self.pi_ip_entry = tk.Entry(row, width=18)
        self.pi_ip_entry.insert(0, self.config["pi_host"])
        self.pi_ip_entry.pack(side="left", padx=5)
        tk.Button(row, text="保存", command=self.save_pi_ip).pack(side="left", padx=5)
        info.pack(fill="x", pady=5)

        save_frame = tk.LabelFrame(left_panel, text="儲存 / 載入")
        save_frame.pack(fill="x", pady=5)

        tk.Label(save_frame, text="名稱").grid(row=0, column=0, padx=5, pady=5)
        self.name_entry = tk.Entry(save_frame, width=25)
        self.name_entry.grid(row=0, column=1, padx=5, pady=5)

        tk.Button(save_frame, text="保存目前 Timeline", command=self.save_current_timeline).grid(row=0, column=2, padx=5, pady=5)
        tk.Button(save_frame, text="重新整理清單", command=self.refresh_saved_list).grid(row=0, column=3, padx=5, pady=5)
        tk.Button(save_frame, text="載入選取項目", command=self.load_selected_timeline).grid(row=0, column=4, padx=5, pady=5)
        tk.Button(save_frame, text="刪除選取項目", command=self.delete_selected_timeline).grid(row=0, column=5, padx=5, pady=5)

        self.saved_listbox = tk.Listbox(save_frame, height=5, exportselection=False)
        self.saved_listbox.grid(row=1, column=0, columnspan=6, sticky="we", padx=5, pady=5)

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

        columns = ("idx", "type", "button", "at", "at_jitter", "group")
        self.tree = ttk.Treeview(right_panel, columns=columns, show="headings", height=8, selectmode="extended")
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=92)
        self.tree.pack(fill="both", expand=True, pady=(0, 8))

        bottom = tk.Frame(right_panel)
        bottom.pack(fill="both", expand=True)

        tk.Label(bottom, text="JSON 預覽 / 分析結果").pack(anchor="w")
        self.text = tk.Text(bottom, height=12)
        self.text.pack(fill="both", expand=True)

        self.refresh_saved_list()
        self.restore_last_selected()
        self.poll_pi_status()

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
            self.tree.insert("", "end", iid=str(i), values=(
                i,
                ev["type"],
                ev["button"],
                ev["at"],
                ev["at_jitter"],
                grp
            ))

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

    def start_record(self):
        global recording, recording_start, events, pressed_keys
        events = []
        pressed_keys = set()
        recording = True
        recording_start = time.time()
        self.timeline = []
        self.current_loaded_from_saved = False
        self.set_frontend_error("")
        self.set_status("錄製中...")
        self.write_text({"status": "recording"})
        self.refresh_tree()

    def stop_record(self):
        global recording
        recording = False
        self.set_status("已停止錄製，共 {} 筆事件".format(len(events)))
        self.write_text(events)

    def analyze(self):
        global events, recording_start
        if not events:
            messagebox.showwarning("提醒", "目前沒有錄到資料")
            return

        self.timeline = build_timeline(events, recording_start)
        self.refresh_tree()
        self.refresh_preview()
        self.current_loaded_from_saved = False

        if not self.current_name:
            self.current_name = "未命名_{}".format(time.strftime("%H%M%S"))

        self.name_entry.delete(0, tk.END)
        self.name_entry.insert(0, self.current_name)

        self.update_current_labels()
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

        save_named_timeline(name, self.timeline)
        self.current_name = name
        self.current_loaded_from_saved = True
        self.config["last_selected_name"] = name
        save_config(self.config)

        self.refresh_saved_list()
        self.select_saved_name(name)
        self.update_current_labels()
        self.refresh_preview()
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

        self.timeline = data.get("events", [])
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

        self.refresh_tree()
        for idx in selected:
            self.tree.selection_add(str(idx))

        self.current_loaded_from_saved = False
        self.update_current_labels()
        self.refresh_preview()
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

        self.current_loaded_from_saved = False
        self.refresh_tree()
        self.refresh_preview()
        self.update_current_labels()
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

        self.current_loaded_from_saved = False
        self.refresh_tree()
        for idx in selected:
            self.tree.selection_add(str(idx))

        self.refresh_preview()
        self.update_current_labels()
        self.set_status("已將選取列 jitter 清為 0")

    def ping_pi(self):
        try:
            self.config["pi_host"] = self.pi_ip_entry.get().strip() or DEFAULT_PI_HOST
            save_config(self.config)
            self.update_current_labels()

            self.set_frontend_error("")
            self.request_pi({"action": "ping"})
            self.set_status("Pi 連線正常：{}".format(self.config["pi_host"]))
        except Exception as e:
            self.set_frontend_error(str(e))
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

        payload = {
            "action": "run_timeline",
            "events": self.timeline
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
                self.set_status("已送出：{} -> {}".format(display_name, self.config["pi_host"]))
        except Exception as e:
            self.set_frontend_error(str(e))
            messagebox.showerror("傳送失敗", str(e))


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
