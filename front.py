# -*- coding: utf-8 -*-
import json
import os
import re
import socket
import time
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


def send_to_pi(payload, pi_host):
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.settimeout(8)
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
    def poll_pi_status(self):
        try:
            pi_host = self.pi_ip_entry.get().strip() or DEFAULT_PI_HOST
            res = send_to_pi({"action": "status"}, pi_host)

            run_status = res.get("run_status", {})
            state = run_status.get("state", "unknown")
            mode = run_status.get("mode", "")
            message = run_status.get("message", "")

            if state == "running":
                self.status_var.set("Pi 執行中：{} / {}".format(mode, message))
            elif state == "idle":
                pass
            elif state == "stopped":
                self.status_var.set("Pi 已停止：{}".format(message))
            elif state == "error":
                self.status_var.set("Pi 錯誤：{}".format(message))
        except Exception:
            pass

        self.root.after(500, self.poll_pi_status)
    def __init__(self, root):
        ensure_dirs()
        self.config = load_config()

        self.root = root
        self.root.title("鍵盤錄製 / Timeline 分析 / 傳送")
        self.timeline = []
        self.current_name = ""
        self.current_loaded_from_saved = False

        top = tk.Frame(root)
        top.pack(fill="x", padx=10, pady=8)

        self.status_var = tk.StringVar(value="尚未錄製")
        tk.Label(top, textvariable=self.status_var).pack(side="left")

        btn_start = tk.Button(top, text="開始錄製", command=self.start_record)
        btn_start.pack(side="left", padx=5)

        btn_stop_record = tk.Button(top, text="停止錄製", command=self.stop_record)
        btn_stop_record.pack(side="left", padx=5)

        btn_analyze = tk.Button(top, text="分析 Timeline", command=self.analyze)
        btn_analyze.pack(side="left", padx=5)

        btn_send = tk.Button(
            top,
            text="送出執行",
            command=self.send_timeline,
            bg="#c8f7c5",      # 淺綠
            activebackground="#b6efb0"
        )
        btn_send.pack(side="left", padx=5)

        btn_stop_pi = tk.Button(
            top,
            text="停止",
            command=self.stop_pi,
            bg="#ff8c69",      # 橘紅
            activebackground="#ff7f50"
        )
        btn_stop_pi.pack(side="left", padx=5)

        btn_ping = tk.Button(
            top,
            text="測試連線",
            command=self.ping_pi,
            bg="#d9d9d9",      # 灰色
            activebackground="#cfcfcf"
        )
        btn_ping.pack(side="left", padx=5)
        info = tk.LabelFrame(root, text="目前套用資訊")
        row = tk.Frame(info)
        row.pack(fill="x", padx=8, pady=2)

        tk.Label(row, text="Pi IP：").pack(side="left")

        self.pi_ip_entry = tk.Entry(row, width=18)
        self.pi_ip_entry.insert(0, self.config["pi_host"])
        self.pi_ip_entry.pack(side="left", padx=5)

        tk.Button(row, text="保存", command=self.save_pi_ip).pack(side="left", padx=5)
        info.pack(fill="x", padx=10, pady=5)

        self.current_name_var = tk.StringVar(value="目前資料：未命名 / 未儲存")
        self.current_pi_var = tk.StringVar(value="Pi IP：{}".format(self.config["pi_host"]))

        tk.Label(info, textvariable=self.current_name_var).pack(anchor="w", padx=8, pady=2)
        tk.Label(info, textvariable=self.current_pi_var).pack(anchor="w", padx=8, pady=2)

        
        save_frame = tk.LabelFrame(root, text="儲存 / 載入")
        save_frame.pack(fill="x", padx=10, pady=5)

        tk.Label(save_frame, text="名稱").grid(row=0, column=0, padx=5, pady=5)
        self.name_entry = tk.Entry(save_frame, width=25)
        self.name_entry.grid(row=0, column=1, padx=5, pady=5)

        tk.Button(save_frame, text="保存目前 Timeline", command=self.save_current_timeline).grid(row=0, column=2, padx=5, pady=5)
        tk.Button(save_frame, text="重新整理清單", command=self.refresh_saved_list).grid(row=0, column=3, padx=5, pady=5)
        tk.Button(save_frame, text="載入選取項目", command=self.load_selected_timeline).grid(row=0, column=4, padx=5, pady=5)
        tk.Button(save_frame, text="刪除選取項目", command=self.delete_selected_timeline).grid(row=0, column=5, padx=5, pady=5)

        self.saved_listbox = tk.Listbox(save_frame, height=5, exportselection=False)
        self.saved_listbox.grid(row=1, column=0, columnspan=6, sticky="we", padx=5, pady=5)

        jitter_frame = tk.LabelFrame(root, text="jitter 設定")
        jitter_frame.pack(fill="x", padx=10, pady=5)

        tk.Label(jitter_frame, text="at jitter").grid(row=0, column=0, padx=5, pady=5)
        self.at_jitter_entry = tk.Entry(jitter_frame, width=10)
        self.at_jitter_entry.insert(0, "0.00")
        self.at_jitter_entry.grid(row=0, column=1, padx=5, pady=5)

        tk.Button(jitter_frame, text="套用到選取列", command=self.apply_jitter_to_selected).grid(row=0, column=2, padx=5, pady=5)
        tk.Button(jitter_frame, text="套用到全部 event", command=self.apply_jitter_to_all).grid(row=0, column=3, padx=5, pady=5)
        tk.Button(jitter_frame, text="選取列清成 0", command=self.clear_jitter_selected).grid(row=0, column=4, padx=5, pady=5)

        columns = ("idx", "type", "button", "at", "at_jitter", "group")
        self.tree = ttk.Treeview(root, columns=columns, show="headings", height=8, selectmode="extended")
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=110)
        self.tree.pack(fill="both", expand=True, padx=10, pady=10)

        bottom = tk.Frame(root)
        bottom.pack(fill="both", expand=True, padx=10, pady=5)

        tk.Label(bottom, text="JSON 預覽 / 分析結果").pack(anchor="w")
        self.text = tk.Text(bottom, height=10)
        self.text.pack(fill="both", expand=True)

        self.refresh_saved_list()
        self.restore_last_selected()

    def update_current_labels(self):
        if self.current_name:
            source_note = "已儲存" if self.current_loaded_from_saved else "目前工作中"
            self.current_name_var.set("目前資料：{} ({})".format(self.current_name, source_note))
        else:
            self.current_name_var.set("目前資料：未命名 / 未儲存")

        self.current_pi_var.set("Pi IP：{}".format(self.config["pi_host"]))

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
            messagebox.showwarning("提醒", "Pi IP 不可空白")
            return

        self.config["pi_host"] = ip
        save_config(self.config)
        self.update_current_labels()
        self.status_var.set("已保存 Pi IP：{}".format(ip))

    def start_record(self):
        global recording, recording_start, events, pressed_keys
        events = []
        pressed_keys = set()
        recording = True
        recording_start = time.time()
        self.timeline = []
        self.current_loaded_from_saved = False
        self.status_var.set("錄製中...")
        self.write_text({"status": "recording"})
        self.refresh_tree()

    def stop_record(self):
        global recording
        recording = False
        self.status_var.set("已停止錄製，共 {} 筆事件".format(len(events)))
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
        self.status_var.set("Timeline 分析完成，共 {} 筆 event".format(len(self.timeline)))

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
        self.status_var.set("已保存 Timeline：{}".format(name))

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
        self.status_var.set("已載入：{}".format(self.current_name))

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
        self.status_var.set("已刪除：{}".format(name))

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
        self.status_var.set("已套用 jitter 到 {} 筆選取列".format(len(selected)))

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
        self.status_var.set("已套用 jitter 到全部 event")

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
        self.status_var.set("已將選取列 jitter 清為 0")

    def ping_pi(self):
        try:
            self.config["pi_host"] = self.pi_ip_entry.get().strip() or DEFAULT_PI_HOST
            save_config(self.config)
            self.update_current_labels()

            res = send_to_pi({"action": "ping"}, self.config["pi_host"])
            self.write_text({
                "current_name": self.current_name,
                "pi_host": self.config["pi_host"],
                "response": res
            })
            self.status_var.set("Pi 連線正常：{}".format(self.config["pi_host"]))
        except Exception as e:
            messagebox.showerror("Pi 連線失敗", str(e))

    def stop_pi(self):
        try:
            self.config["pi_host"] = self.pi_ip_entry.get().strip() or DEFAULT_PI_HOST
            save_config(self.config)
            self.update_current_labels()

            res = send_to_pi({"action": "stop"}, self.config["pi_host"])
            self.write_text({
                "current_name": self.current_name,
                "pi_host": self.config["pi_host"],
                "response": res
            })
            self.status_var.set("已停止 Pi：{}".format(self.config["pi_host"]))
        except Exception as e:
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
            res = send_to_pi(payload, self.config["pi_host"])
            self.write_text({
                "sending_name": display_name,
                "pi_host": self.config["pi_host"],
                "response": res
            })

            if res.get("status") == "stopped":
                self.status_var.set("Pi 已停止執行：{} -> {}".format(display_name, self.config["pi_host"]))
            else:
                self.status_var.set("已送出：{} -> {}".format(display_name, self.config["pi_host"]))
        except Exception as e:
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