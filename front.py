# -*- coding: utf-8 -*-
import json
import os
import re
import socket
import errno
import copy
import time
import threading 
import random
import uuid
import tkinter as tk
from tkinter import ttk, simpledialog, colorchooser, messagebox
from pynput import keyboard

DEFAULT_PI_HOST = "192.168.100.140"
PI_PORT = 5000
RUNTIME_CONTRACT_VERSION = "v1"
BUFF_SKIP_MODE_WALK = "walk"
BUFF_SKIP_MODE_COMPRESS = "compress"
NEGATIVE_GROUP_ANCHOR_GAP_SEC = 0.2
RUNTIME_POLL_INTERVAL_MS = 300
CONTROL_REQUEST_TIMEOUT_SEC = 8.0
STATUS_REQUEST_TIMEOUT_SEC = 0.45
FIRST_EVENT_DEADLINE_GRACE_SEC = 0.8
ACK_TIMEOUT_MS = 1500
START_TIMEOUT_MS = 5000
PROGRESS_STALL_MS = 3000
ROUND_TRACE_REASON_LABELS = {
    "randat_disabled": "未啟用 randat（無可用 randat 列）",
    "fallback_no_free_slot": "無可用 idx（候選位置皆被占用）",
    "group_skipped": "此群組略過",
    "random_pick": "從可用候選池公平抽籤",
    "rule_blocked": "規則阻擋",
}

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
    "補充：『套用偏移』是手動把選取列之後的時間整段平移；『糾正複製體』是依 buff_group 負值重算複製體群組的正確 at；『at 交換』可快速對調兩列 at 時間。\n"
    "randat：type=randat 的列是候選落點；每個 buff_group 會整包換位，組內順序不拆。\n"
    "at_jitter：前端每輪重算為 base + random(0~j)，僅延後。\n"
    "最後一欄 group：用於顯示系統重疊群組；若無系統綁定可視為使用者自訂分組語意。\n"
    "淺藍和淺黃都是 buff。\n"
    "【運行中的淺黃】: 代表自己搶到自己原位執行。\n"
    "【運行中的淺藍】: 代表搶到空格或者別人位置。"
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
                "right_paned_sash_y": None,
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
                "right_paned_sash_y": ui_layout.get("right_paned_sash_y"),
                "right_paned_ratio": ui_layout.get("right_paned_ratio"),
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
                "right_paned_sash_y": None,
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


def move_rows_with_ab_gap_compensation(events, selected_indexes, direction):
    direction_text = str(direction or "").strip().lower()
    if direction_text not in {"up", "down"}:
        raise ValueError("direction 只能是 up/down")
    if not isinstance(events, list):
        raise ValueError("events 必須是 list")
    if not selected_indexes:
        return events, [], {}

    selected_sorted = sorted(set(int(i) for i in selected_indexes))
    count = len(events)
    if count <= 1:
        return events, selected_sorted, {}
    if direction_text == "up" and selected_sorted[0] <= 0:
        return events, selected_sorted, {}
    if direction_text == "down" and selected_sorted[-1] >= count - 1:
        return events, selected_sorted, {}

    head = selected_sorted[0]
    tail = selected_sorted[-1]

    def _gap_a(rows, h):
        if h <= 0:
            return None
        return round(_safe_float(rows[h].get("at", 0.0)) - _safe_float(rows[h - 1].get("at", 0.0)), 4)

    def _gap_b(rows, t):
        if t >= len(rows) - 1:
            return None
        return round(_safe_float(rows[t + 1].get("at", 0.0)) - _safe_float(rows[t].get("at", 0.0)), 4)

    gap_a_before = _gap_a(events, head)
    gap_b_before = _gap_b(events, tail)

    if direction_text == "up":
        for idx in selected_sorted:
            events[idx - 1], events[idx] = events[idx], events[idx - 1]
        moved_indexes = [idx - 1 for idx in selected_sorted]
    else:
        for idx in sorted(selected_sorted, reverse=True):
            events[idx + 1], events[idx] = events[idx], events[idx + 1]
        moved_indexes = [idx + 1 for idx in selected_sorted]

    new_head = min(moved_indexes)
    new_tail = max(moved_indexes)
    gap_a_after = _gap_a(events, new_head)
    gap_b_after = _gap_b(events, new_tail)

    preserve_target = None
    if gap_a_before is None and gap_b_before is None:
        preserve_target = None
    elif gap_a_before is None:
        preserve_target = "B"
    elif gap_b_before is None:
        preserve_target = "A"
    else:
        preserve_target = "A" if gap_a_before <= gap_b_before else "B"

    requested_compensation = 0.0
    if preserve_target == "A" and gap_a_before is not None and gap_a_after is not None:
        requested_compensation = round(gap_a_before - gap_a_after, 4)
    elif preserve_target == "B" and gap_b_before is not None and gap_b_after is not None:
        requested_compensation = round(gap_b_after - gap_b_before, 4)

    block_head_at = _safe_float(events[new_head].get("at", 0.0))
    block_tail_at = _safe_float(events[new_tail].get("at", 0.0))
    min_shift = -10**9
    max_shift = 10**9
    if new_head > 0:
        prev_at = _safe_float(events[new_head - 1].get("at", 0.0))
        min_shift = max(min_shift, prev_at - block_head_at)
    if new_tail < len(events) - 1:
        next_at = _safe_float(events[new_tail + 1].get("at", 0.0))
        max_shift = min(max_shift, next_at - block_tail_at)
    applied_compensation = min(max(requested_compensation, min_shift), max_shift)
    applied_compensation = round(applied_compensation, 4)

    if abs(applied_compensation) > 0.0001:
        for idx in moved_indexes:
            shifted = max(0.0, _safe_float(events[idx].get("at", 0.0)) + applied_compensation)
            events[idx]["at"] = round(shifted, 2)

    gap_a_final = _gap_a(events, new_head)
    gap_b_final = _gap_b(events, new_tail)
    preserved = False
    if preserve_target == "A" and gap_a_before is not None and gap_a_final is not None:
        preserved = abs(gap_a_before - gap_a_final) <= 0.0001
    elif preserve_target == "B" and gap_b_before is not None and gap_b_final is not None:
        preserved = abs(gap_b_before - gap_b_final) <= 0.0001

    move_meta = {
        "direction": direction_text,
        "selected_before": selected_sorted,
        "selected_after": moved_indexes,
        "range_before": [head, tail],
        "range_after": [new_head, new_tail],
        "gap_a_before": gap_a_before,
        "gap_b_before": gap_b_before,
        "gap_a_after_swap": gap_a_after,
        "gap_b_after_swap": gap_b_after,
        "gap_a_final": gap_a_final,
        "gap_b_final": gap_b_final,
        "preserve_target": preserve_target,
        "preserved": bool(preserved),
        "compensation_requested": round(requested_compensation, 4),
        "compensation_applied": round(applied_compensation, 4),
    }
    return events, moved_indexes, move_meta


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
            "at_random_sec": 0.0,
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
        buttons = []
        for ev in grp:
            btn = str(ev.get("button", "")).strip()
            if btn and btn not in buttons:
                buttons.append(btn)
        buttons_text = "+".join(buttons) if buttons else "-"
        summary.append({
            "group": i,
            "group_buttons": buttons,
            "group_label": "group {} ({})".format(i, buttons_text),
            "at_range": [grp[0]["at"], grp[-1]["at"]],
            "events": grp
        })
    return summary


def _safe_float(value, default=0.0):
    try:
        return float(value)
    except (TypeError, ValueError):
        return float(default)


DEFAULT_PR_SEGMENTS = [
    (0, 25),
    (26, 50),
    (51, 70),
    (71, 95),
    (95, 99),
]


def build_extended_pr_segments():
    segments = list(DEFAULT_PR_SEGMENTS)
    for start in range(0, 100, 10):
        segments.append((start, start + 9))
    return segments


def build_pr_gap_pairs(events):
    ordered = sorted(
        [dict(ev) for ev in events],
        key=lambda x: (_safe_float(x.get("at", 0.0)), int(x.get("idx", 0)))
    )
    pairs = []
    for i in range(len(ordered) - 1):
        a = ordered[i]
        b = ordered[i + 1]
        pairs.append({
            "idx_a": int(a.get("idx", 0)),
            "idx_b": int(b.get("idx", 0)),
            "button_a": str(a.get("button_id", "")),
            "button_b": str(b.get("button_id", "")),
            "type_a": str(a.get("type", "")),
            "type_b": str(b.get("type", "")),
            "at_a": round(_safe_float(a.get("at", 0.0)), 4),
            "at_b": round(_safe_float(b.get("at", 0.0)), 4),
            "gap": round(_safe_float(b.get("at", 0.0)) - _safe_float(a.get("at", 0.0)), 4)
        })
    pairs.sort(key=lambda x: (x["gap"], x["idx_a"], x["idx_b"]))
    for rank, row in enumerate(pairs):
        row["pr_rank"] = rank
    return ordered, pairs


def build_pr_segment_summary(pairs, segments=None):
    if segments is None:
        segments = build_extended_pr_segments()
    summary = []
    for start, end in segments:
        section = [p for p in pairs if start <= int(p.get("pr_rank", -1)) <= end]
        avg_gap = 0.0
        if section:
            avg_gap = round(sum(_safe_float(item.get("gap", 0.0)) for item in section) / len(section), 4)
        summary.append({
            "range": "pr{}~pr{}".format(start, end),
            "avg_gap": avg_gap,
            "pairs": [
                {
                    "pr_rank": int(item["pr_rank"]),
                    "idx_a": int(item["idx_a"]),
                    "idx_b": int(item["idx_b"])
                }
                for item in section
            ]
        })
    return summary


def analyze_pr_gap_events(events, top_n=100):
    ordered, pairs = build_pr_gap_pairs(events)
    if top_n is None:
        selected_pairs = list(pairs)
    else:
        selected_pairs = pairs[:max(0, int(top_n))]
    return {
        "min_all_pairs": round(min((_safe_float(p.get("gap", 0.0)) for p in pairs), default=0.0), 4),
        "pr_pairs": selected_pairs,
        "segments": build_pr_segment_summary(pairs),
        "pair_count": len(pairs),
        "events_sorted": ordered,
    }


def apply_minimum_gap_by_pairs(events, pairs_snapshot, minimum_gap):
    min_gap = max(0.0, _safe_float(minimum_gap, 0.0))
    working = [dict(ev) for ev in events]
    logs = []
    for pair in sorted([dict(p) for p in pairs_snapshot], key=lambda x: int(x.get("pr_rank", 0))):
        idx_a = int(pair.get("idx_a", 0))
        idx_b = int(pair.get("idx_b", 0))
        at_a = None
        at_b = None
        for ev in working:
            idx = int(ev.get("idx", -1))
            if idx == idx_a:
                at_a = _safe_float(ev.get("at", 0.0))
            if idx == idx_b:
                at_b = _safe_float(ev.get("at", 0.0))
        if at_a is None or at_b is None:
            continue
        current_gap = round(at_b - at_a, 4)
        delta = round(max(0.0, min_gap - current_gap), 4)
        affected = []
        for ev in working:
            if int(ev.get("idx", -1)) >= idx_b:
                ev["at"] = round(_safe_float(ev.get("at", 0.0)) + delta, 4)
                affected.append(int(ev.get("idx", -1)))
        logs.append({
            "pr_rank": int(pair.get("pr_rank", 0)),
            "idx_a": idx_a,
            "idx_b": idx_b,
            "current_gap": current_gap,
            "minimum_gap": min_gap,
            "delta": delta,
            "affected_idx_range": {
                "from_idx": idx_b,
                "count": len(affected),
                "affected_indices": affected
            }
        })
    return working, logs


def detect_jitter_order_risk_pairs(events):
    ordered = sorted(
        [dict(ev) for ev in events],
        key=lambda x: (_safe_float(x.get("at", 0.0)), int(x.get("idx", 0)))
    )
    risks = []
    for i in range(len(ordered) - 1):
        a = ordered[i]
        b = ordered[i + 1]
        gap = round(_safe_float(b.get("at", 0.0)) - _safe_float(a.get("at", 0.0)), 4)
        jitter_a = max(0.0, _safe_float(a.get("at_jitter", 0.0), 0.0))
        if gap < jitter_a:
            level = "overtake_risk"
        elif gap == jitter_a:
            level = "tie_risk"
        else:
            continue
        risks.append({
            "risk_level": level,
            "idx_a": int(a.get("idx", 0)),
            "idx_b": int(b.get("idx", 0)),
            "type_a": str(a.get("type", "")),
            "type_b": str(b.get("type", "")),
            "button_a": str(a.get("button_id", "")),
            "button_b": str(b.get("button_id", "")),
            "at_a": round(_safe_float(a.get("at", 0.0)), 4),
            "at_b": round(_safe_float(b.get("at", 0.0)), 4),
            "gap": gap,
            "at_jitter_a": round(jitter_a, 4),
        })
    return risks


def apply_positive_jitter(base_value, jitter):
    base = max(0.0, _safe_float(base_value, 0.0))
    j = max(0.0, _safe_float(jitter, 0.0))
    return round(base + random.uniform(0.0, j), 4)


def allocate_randat_blocks(events):
    working = [dict(ev) for ev in events]
    rslot = [
        idx for idx, ev in enumerate(working)
        if str(ev.get("type", "")).strip().lower() == "randat"
    ]

    blocks = {}
    bslot = []
    for idx, ev in enumerate(working):
        group = str(ev.get("buff_group", "")).strip()
        if not group:
            continue
        if group not in blocks:
            blocks[group] = {"first": idx, "indices": [idx]}
            bslot.append(idx)
        else:
            blocks[group]["indices"].append(idx)

    for group in blocks:
        first = blocks[group]["first"]
        indices = list(blocks[group]["indices"])
        blocks[group]["rows_len"] = len(indices)
        blocks[group]["span_len"] = len(indices)

    for idx, ev in enumerate(working):
        group = str(ev.get("buff_group", "")).strip()
        if group:
            ev["__runtime_gid"] = group
            continue
        ev["__runtime_gid"] = "__row_{}".format(idx)

    pool = []
    seen = set()
    for idx in (rslot + bslot):
        if idx in seen:
            continue
        seen.add(idx)
        pool.append(idx)

    block_assignments = {}
    round_traces = []
    randat_executed = bool(rslot)

    if not randat_executed:
        for round_idx, group in enumerate(sorted(blocks.keys()), start=1):
            src_anchor = int(blocks[group]["first"])
            src_range = [src_anchor, src_anchor + blocks[group]["span_len"] - 1]
            round_trace = {
                "round": round_idx,
                "buffGroup": group,
                "group_id": group,
                "anchorIdx": src_anchor,
                "candidateIdxList": [],
                "freeCandidateIdxList": [],
                "pickedCandidatePos": None,
                "diceValue": None,
                "pickedIdx": src_anchor,
                "picked_slot": src_anchor,
                "src_range": src_range,
                "dst_range": list(src_range),
                "randatExecuted": False,
                "result": "kept",
                "pickedReason": "randat_disabled",
                "reason": "randat_disabled",
                "self_picked": True,
                "color": "light_yellow"
            }
            block_assignments[group] = {
                "group": group,
                "anchor_index": src_anchor,
                "landed_index": src_anchor,
                "occupies_original": True,
                "src_range": src_range,
                "dst_range": list(src_range),
                "picked_slot": src_anchor,
                "self_picked": True,
                "color": "light_yellow",
                "round_trace": round_trace
            }
            round_traces.append(round_trace)

        for ev in working:
            group = str(ev.get("buff_group", "")).strip()
            if group in block_assignments:
                assign = block_assignments[group]
                ev["runtime_landed_index"] = assign["landed_index"]
                ev["runtime_anchor_index"] = assign["anchor_index"]
                ev["runtime_occupies_original"] = 1
                ev["runtime_self_picked"] = 1
                ev["runtime_group_color"] = assign["color"]
                ev["runtime_group_src_range"] = list(assign["src_range"])
                ev["runtime_group_dst_range"] = list(assign["dst_range"])
                ev["runtime_picked_slot"] = assign["picked_slot"]
            if "__runtime_gid" in ev:
                del ev["__runtime_gid"]
        return working, block_assignments, round_traces

    def _group_sort_key(group_id):
        text = str(group_id).strip()
        try:
            return (0, int(text))
        except Exception:
            return (1, text)

    group_order = sorted(blocks.keys(), key=_group_sort_key)

    for round_idx, group in enumerate(group_order, start=1):
        src_anchor = int(blocks[group]["first"])
        src_range = [src_anchor, src_anchor + blocks[group]["span_len"] - 1]
        candidates = list(pool)
        picked_candidate_pos = None
        dice_value = None
        if candidates:
            picked_candidate_pos = random.randrange(len(candidates))
            picked_slot = int(candidates[picked_candidate_pos])
            pool.pop(picked_candidate_pos)
            reason = "random_pick"
        else:
            picked_slot = src_anchor
            reason = "fallback_no_free_slot"
        if picked_candidate_pos is not None and candidates:
            dice_value = round((picked_candidate_pos + 1) / float(len(candidates)), 4)

        dst_range = [picked_slot, picked_slot + blocks[group]["span_len"] - 1]
        self_picked = (picked_slot == src_anchor)
        color = "light_yellow" if self_picked else "light_blue"
        result = "kept" if self_picked else "applied"
        round_trace = {
            "round": round_idx,
            "buffGroup": group,
            "group_id": group,
            "anchorIdx": blocks[group]["first"],
            "candidateIdxList": candidates,
            "freeCandidateIdxList": candidates,
            "pickedCandidatePos": picked_candidate_pos,
            "diceValue": dice_value,
            "pickedIdx": picked_slot,
            "picked_slot": picked_slot,
            "src_range": src_range,
            "dst_range": dst_range,
            "randatExecuted": randat_executed,
            "result": result,
            "pickedReason": reason,
            "reason": reason,
            "self_picked": bool(self_picked),
            "color": color
        }
        block_assignments[group] = {
            "group": group,
            "anchor_index": blocks[group]["first"],
            "landed_index": picked_slot,
            "occupies_original": bool(self_picked),
            "src_range": src_range,
            "dst_range": dst_range,
            "picked_slot": picked_slot,
            "self_picked": bool(self_picked),
            "color": color,
            "round_trace": round_trace
        }
        round_traces.append(round_trace)

    for ev in working:
        group = str(ev.get("buff_group", "")).strip()
        if group in block_assignments:
            assign = block_assignments[group]
            ev["runtime_landed_index"] = assign["landed_index"]
            ev["runtime_anchor_index"] = assign["anchor_index"]
            ev["runtime_occupies_original"] = 1 if assign["occupies_original"] else 0
            ev["runtime_self_picked"] = 1 if assign["self_picked"] else 0
            ev["runtime_group_color"] = assign["color"]
            ev["runtime_group_src_range"] = list(assign["src_range"])
            ev["runtime_group_dst_range"] = list(assign["dst_range"])
            ev["runtime_picked_slot"] = assign["picked_slot"]
        if "__runtime_gid" in ev:
            del ev["__runtime_gid"]
    return working, block_assignments, round_traces


def get_buff_cell_visual_state(is_candidate, is_applied, is_running, is_focus):
    tags = []
    if bool(is_applied):
        tags.append("bg_applied")
    elif bool(is_candidate):
        tags.append("bg_candidate")
    if bool(is_running):
        tags.append("ring_running")
    if bool(is_focus):
        tags.append("focus_hint")
    return tags


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
        row["at_random_sec"] = round(abs(float(row.get("at_random_sec", 0.0))), 4)
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
            self.set_status("沒有可執行的 undo")
            return "break"
        before = bucket["undo"].pop()
        bucket["redo"].append(self.copy_events(self.timeline))
        self.timeline = self.copy_events(before)
        self.mark_timeline_dirty()
        self.set_status("已執行 undo")
        return "break"

    def redo_timeline(self, _event=None):
        if not self._ensure_runtime_editable():
            return "break"
        bucket = self._history_bucket()
        if not bucket["redo"]:
            self.set_status("沒有可執行的 redo")
            return "break"
        after = bucket["redo"].pop()
        bucket["undo"].append(self.copy_events(self.timeline))
        self.timeline = self.copy_events(after)
        self.mark_timeline_dirty()
        self.set_status("已執行 redo")
        return "break"

    def restore_original_timeline(self):
        original = self.timeline_meta.get("original_events", [])
        if not original:
            return False
        before = self._begin_timeline_change()
        self.timeline = self.copy_events(original)
        self._finalize_timeline_change(before)
        self.current_loaded_from_saved = False
        self.update_current_labels()
        self.refresh_tree()
        self.refresh_preview()
        return True

    def restore_latest_saved_timeline(self):
        latest_saved = self.timeline_meta.get("latest_saved_events", [])
        if not latest_saved:
            return False
        before = self._begin_timeline_change()
        self.timeline = self.copy_events(latest_saved)
        self._finalize_timeline_change(before)
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

    def validate_buff_group_contiguity(self, events):
        first_last = {}
        for idx, ev in enumerate(events):
            group = str(ev.get("buff_group", "")).strip()
            if not group:
                continue
            if group not in first_last:
                first_last[group] = [idx, idx]
            else:
                first_last[group][1] = idx

        for group, (start, end) in first_last.items():
            for idx in range(start, end + 1):
                middle_group = str(events[idx].get("buff_group", "")).strip()
                if middle_group != group:
                    raise ValueError(
                        "buff_group {} 必須連續，idx {} 到 idx {} 之間不可夾其他列（idx {} 為 {}）。".format(
                            group,
                            start,
                            end,
                            idx,
                            middle_group if middle_group else "空白群組"
                        )
                    )

    def update_error_text(self, widget, message):
        widget.config(state="normal")
        widget.delete("1.0", tk.END)
        widget.insert(tk.END, message or "無")
        widget.config(state="disabled")

    def set_frontend_error(self, message):
        self.frontend_error_main = str(message or "").strip()
        self._render_frontend_error()

    def set_backend_error(self, message):
        self.update_error_text(self.backend_error_text, message)

    def _format_backend_error(self, response):
        if not isinstance(response, dict):
            return "Pi 回傳錯誤"
        message = str(response.get("message", "Pi 回傳錯誤")).strip() or "Pi 回傳錯誤"
        code = str(response.get("code", "")).strip()
        phase = str(response.get("phase", "")).strip()
        server_task_id = str(response.get("server_task_id", "")).strip()
        diag = response.get("diag", {})
        if not isinstance(diag, dict):
            diag = {}
        last_state = str(diag.get("last_state", "")).strip()
        elapsed_ms = diag.get("elapsed_ms")
        threshold_ms = diag.get("threshold_ms")

        lines = [message]
        if code:
            lines.append("code={}".format(code))
        if phase:
            lines.append("phase={}".format(phase))
        if server_task_id:
            lines.append("server_task_id={}".format(server_task_id))
        if last_state:
            lines.append("last_state={}".format(last_state))
        if elapsed_ms is not None:
            lines.append("elapsed_ms={}".format(elapsed_ms))
        if threshold_ms is not None:
            lines.append("threshold_ms={}".format(threshold_ms))
        return "\n".join(lines)

    def clear_errors(self):
        self.frontend_error_main = ""
        self.last_control_error = ""
        self.last_status_error = ""
        self.set_frontend_error("")
        self.set_backend_error("")

    def _render_frontend_error(self):
        lines = []
        if self.frontend_error_main:
            lines.append(self.frontend_error_main)
        if self.last_control_error:
            lines.append("最後 control 錯誤：{}".format(self.last_control_error))
        if self.last_status_error:
            lines.append("最後 status 錯誤：{}".format(self.last_status_error))
        self.update_error_text(self.frontend_error_text, "\n".join(lines))

    def _set_channel_error(self, channel, message):
        text = str(message or "").strip()
        if channel == "status":
            self.last_status_error = text
        else:
            self.last_control_error = text
        self._render_frontend_error()

    def _clear_channel_error(self, channel):
        if channel == "status":
            self.last_status_error = ""
        else:
            self.last_control_error = ""
        self._render_frontend_error()

    def _classify_request_error(self, exc):
        if isinstance(exc, socket.timeout):
            return "timeout"
        if isinstance(exc, json.JSONDecodeError):
            return "json_parse_failed"
        if isinstance(exc, (ConnectionResetError, BrokenPipeError, ConnectionAbortedError)):
            return "connection_reset"
        if isinstance(exc, OSError):
            if getattr(exc, "errno", None) in (errno.ECONNRESET, errno.EPIPE, errno.ECONNABORTED):
                return "connection_reset"
            return "connection_error"
        return "request_error"

    def _get_virtual_desktop_bounds(self):
        try:
            vx = int(self.root.winfo_vrootx())
            vy = int(self.root.winfo_vrooty())
            vw = int(self.root.winfo_vrootwidth())
            vh = int(self.root.winfo_vrootheight())
            if vw > 0 and vh > 0:
                return vx, vy, vx + vw, vy + vh
        except Exception:
            pass
        sw = int(self.root.winfo_screenwidth())
        sh = int(self.root.winfo_screenheight())
        return 0, 0, sw, sh

    def _get_table_anchor_rect(self):
        try:
            self.root.update_idletasks()
            tree = self.tree
            return (
                int(tree.winfo_rootx()),
                int(tree.winfo_rooty()),
                max(1, int(tree.winfo_width())),
                max(1, int(tree.winfo_height()))
            )
        except Exception:
            return (
                int(self.root.winfo_rootx()),
                int(self.root.winfo_rooty()),
                max(1, int(self.root.winfo_width())),
                max(1, int(self.root.winfo_height()))
            )

    def _compute_app_dialog_position(self, dialog_w, dialog_h):
        area_x, area_y, area_w, area_h = self._get_table_anchor_rect()
        x = int(area_x + (area_w - dialog_w) / 2)
        y = int(area_y + area_h * 0.62 - dialog_h / 2)
        min_x, min_y, max_x, max_y = self._get_virtual_desktop_bounds()
        x = max(min_x, min(x, max_x - dialog_w))
        y = max(min_y, min(y, max_y - dialog_h))
        return x, y

    def _position_dialog_to_app_lower_center(self, dialog):
        try:
            dialog.update_idletasks()
            dialog_w = max(1, int(dialog.winfo_width()))
            dialog_h = max(1, int(dialog.winfo_height()))
            x, y = self._compute_app_dialog_position(dialog_w, dialog_h)
            dialog.geometry("{}x{}+{}+{}".format(dialog_w, dialog_h, x, y))
        except Exception:
            pass

    def _show_app_dialog(self, title, message, dialog_type="info", buttons="ok"):
        if not hasattr(self.root, "tk"):
            if buttons == "ok":
                if dialog_type == "warning":
                    messagebox.showwarning(title, message)
                elif dialog_type == "error":
                    messagebox.showerror(title, message)
                else:
                    messagebox.showinfo(title, message)
                return "ok"
            if buttons == "yesno":
                try:
                    return messagebox.askyesno(title, message)
                except Exception:
                    return True
            if buttons == "yesnocancel":
                try:
                    return messagebox.askyesnocancel(title, message)
                except Exception:
                    return None

        result = {"value": None}
        dialog = tk.Toplevel(self.root)
        dialog.title(str(title or "提示"))
        dialog.transient(self.root)
        dialog.resizable(False, False)
        dialog.grab_set()

        container = tk.Frame(dialog, padx=16, pady=14)
        container.pack(fill="both", expand=True)

        message_text = str(message or "")
        tk.Label(
            container,
            text=message_text,
            justify="left",
            anchor="w",
            wraplength=420
        ).grid(row=0, column=0, sticky="w")

        btn_row = tk.Frame(container)
        btn_row.grid(row=1, column=0, sticky="e", pady=(14, 0))

        def choose(value):
            result["value"] = value
            dialog.destroy()

        def on_close():
            if buttons == "yesnocancel":
                choose(None)
            elif buttons == "yesno":
                choose(False)
            else:
                choose("ok")

        focus_btn = None
        if buttons == "ok":
            focus_btn = tk.Button(btn_row, text="確定", width=10, command=lambda: choose("ok"))
            focus_btn.pack(side="right")
            dialog.bind("<Return>", lambda _e: choose("ok"))
            dialog.bind("<Escape>", lambda _e: choose("ok"))
        elif buttons == "yesno":
            btn_no = tk.Button(btn_row, text="否", width=10, command=lambda: choose(False))
            btn_no.pack(side="right")
            focus_btn = tk.Button(btn_row, text="是", width=10, command=lambda: choose(True))
            focus_btn.pack(side="right", padx=(0, 6))
            dialog.bind("<Return>", lambda _e: choose(True))
            dialog.bind("<Escape>", lambda _e: choose(False))
        elif buttons == "yesnocancel":
            btn_cancel = tk.Button(btn_row, text="取消", width=10, command=lambda: choose(None))
            btn_cancel.pack(side="right")
            btn_no = tk.Button(btn_row, text="否", width=10, command=lambda: choose(False))
            btn_no.pack(side="right", padx=(0, 6))
            focus_btn = tk.Button(btn_row, text="是", width=10, command=lambda: choose(True))
            focus_btn.pack(side="right", padx=(0, 6))
            dialog.bind("<Return>", lambda _e: choose(True))
            dialog.bind("<Escape>", lambda _e: choose(None))
        else:
            raise ValueError("不支援的對話窗按鈕模式：{}".format(buttons))

        dialog.protocol("WM_DELETE_WINDOW", on_close)
        self._position_dialog_to_app_lower_center(dialog)
        dialog.focus_force()
        if focus_btn is not None:
            focus_btn.focus_set()
        self.root.wait_window(dialog)
        return result["value"]

    def show_warning(self, title, message, **kwargs):
        self.set_status(str(message or title or "提醒"))
        if not hasattr(self.root, "tk"):
            return self._show_app_dialog(title, message, dialog_type="warning", buttons="ok")
        return "ok"

    def show_info(self, title, message, **kwargs):
        self.set_status(str(message or title or "提示"))
        if not hasattr(self.root, "tk"):
            return self._show_app_dialog(title, message, dialog_type="info", buttons="ok")
        return "ok"

    def show_error(self, title, message, **kwargs):
        text = str(message or title or "錯誤")
        self.set_status(text)
        self.set_frontend_error(text)
        if not hasattr(self.root, "tk"):
            return self._show_app_dialog(title, text, dialog_type="error", buttons="ok")
        return "ok"

    def confirm(self, title, message, **kwargs):
        return self._show_app_dialog(title, message, dialog_type="question", buttons="yesno")

    def confirm_cancel(self, title, message, **kwargs):
        return self._show_app_dialog(title, message, dialog_type="question", buttons="yesnocancel")

    def set_status(self, message):
        self.status_var.set(message)

    def _is_script_switch_for_load(self, target_name):
        target = str(target_name or "").strip()
        current = str(getattr(self, "current_name", "") or "").strip()
        return bool(target) and target != current

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

    def request_pi(self, payload, success_status=None, write_response=True, channel="control", timeout=None):
        payload_action = str(payload.get("action", "")).strip().lower()
        payload_type = str(payload.get("type", "")).strip().lower()
        if self.offline_mode and payload_action != "ping":
            raise ConnectionError("目前為離線模式，請先按「測試連線」")

        action = payload_action or payload_type
        retryable_actions = {"ping", "status", "list_buttons"}
        allow_retry = action in retryable_actions
        last_err = None
        req_timeout = float(timeout if timeout is not None else (
            STATUS_REQUEST_TIMEOUT_SEC if channel == "status" else CONTROL_REQUEST_TIMEOUT_SEC
        ))
        lock = self.status_request_lock if channel == "status" else self.control_request_lock
        with lock:
            max_attempts = 2 if allow_retry else 1
            for retry in range(max_attempts):
                try:
                    self.ensure_connection(channel=channel, timeout=req_timeout)
                    conn = self.status_conn if channel == "status" else self.control_conn
                    conn_file = self.status_conn_file if channel == "status" else self.control_conn_file
                    wire = json.dumps(payload, ensure_ascii=False) + "\n"
                    conn.sendall(wire.encode("utf-8"))
                    line = conn_file.readline()
                    if not line:
                        raise ConnectionError("連線已中斷，未收到回應")
                    res = json.loads(line)
                    break
                except Exception as e:
                    last_err = e
                    self.close_connection(channel=channel, silent=True)
                    if retry < max_attempts - 1:
                        time.sleep(0.06 if channel == "status" else 0.15)
                        continue
                    kind = self._classify_request_error(last_err)
                    err_msg = str(last_err)
                    if kind == "timeout":
                        if channel == "control" and action in {"start_task", "run_macro", "stop"}:
                            err_msg = "控制請求逾時：未在 {:.2f} 秒內收到 ACK 回應（action={}）".format(
                                req_timeout,
                                action or "unknown"
                            )
                        else:
                            err_msg = "{}（未收到回應）".format(err_msg)
                    err = "[{}:{}] {}".format(channel, kind, err_msg)
                    self._set_channel_error(channel, err)
                    if channel == "control":
                        self.set_connected(False)
                    if success_status:
                        self.set_status(success_status + "（前端連線失敗）")
                    if kind == "timeout":
                        raise TimeoutError(err) from last_err
                    raise

        self._clear_channel_error(channel)
        if channel == "control":
            self.set_connected(True)
        status = res.get("status")
        if status in ("error", "busy"):
            self.set_backend_error(self._format_backend_error(res))
        else:
            self.set_backend_error("")

        if write_response:
            self.write_text({
                "current_name": self.current_name,
                "pi_host": self.config["pi_host"],
                "channel": channel,
                "request": payload,
                "response": res
            })
        return res

    def _is_contract_version_mismatch(self, response):
        if not isinstance(response, dict):
            return False
        code = str(response.get("code", "")).strip().upper()
        return code == "CONTRACT_VERSION_MISMATCH"

    def _build_contract_version_mismatch_message(self, response):
        diag = response.get("diag", {}) if isinstance(response, dict) else {}
        if not isinstance(diag, dict):
            diag = {}
        expected = diag.get("expected")
        actual = diag.get("actual")
        expected_text = str(expected).strip() if expected is not None else RUNTIME_CONTRACT_VERSION
        actual_text = "缺失" if actual in (None, "") else str(actual)
        return (
            "前後端執行契約版本不一致，已停止送出流程。\n"
            "請更新前端或後端後重試。\n"
            "預期版本：{}\n"
            "實際版本：{}".format(expected_text, actual_text)
        )

    def ensure_connection(self, channel="control", timeout=None):
        conn = self.status_conn if channel == "status" else self.control_conn
        conn_file = self.status_conn_file if channel == "status" else self.control_conn_file
        if conn is not None and conn_file is not None and (self.connected or channel == "status"):
            return
        self.open_connection(channel=channel, timeout=timeout if timeout is not None else 3.0)

    def open_connection(self, timeout=3.0, channel="control"):
        self.close_connection(channel=channel, silent=True)
        host = self.config["pi_host"]
        sock = socket.create_connection((host, PI_PORT), timeout=timeout)
        sock.settimeout(timeout)
        if channel == "status":
            self.status_conn = sock
            self.status_conn_file = sock.makefile("r", encoding="utf-8")
        else:
            self.control_conn = sock
            self.control_conn_file = sock.makefile("r", encoding="utf-8")
            self.set_connected(True)

    def close_connection(self, channel=None, silent=False):
        channels = ["control", "status"] if channel is None else [channel]
        for ch in channels:
            conn_file = self.status_conn_file if ch == "status" else self.control_conn_file
            conn = self.status_conn if ch == "status" else self.control_conn
            if conn_file is not None:
                try:
                    conn_file.close()
                except Exception:
                    pass
            if conn is not None:
                try:
                    conn.close()
                except Exception:
                    pass
            if ch == "status":
                self.status_conn_file = None
                self.status_conn = None
            else:
                self.control_conn_file = None
                self.control_conn = None
        if channel in (None, "control") and not silent:
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
        self.control_request_lock = threading.Lock()
        self.status_request_lock = threading.Lock()
        self.control_conn = None
        self.control_conn_file = None
        self.status_conn = None
        self.status_conn_file = None
        self.frontend_error_main = ""
        self.last_control_error = ""
        self.last_status_error = ""
        self.timeline_runtime_info = {"events": []}
        self.timeline_runtime_by_index = {}
        self.runtime_round_traces = []
        self.runtime_recent_ok_indices = []
        self.runtime_recent_skipped_indices = []
        self.runtime_latest_index = None
        self.last_runtime_signature = ""
        self.pre_run_timeline_snapshot = []
        self.runtime_working_timeline = []
        self.has_pre_run_snapshot = False
        self.runtime_display_frozen = False
        self.runtime_manual_restore_active = False
        self.runtime_move_gap_logs = []
        self.tree_row_color_tags = set()
        self.last_tree_copy_payload = ""
        self.timeline_histories = {}
        self.front_loop_enabled = False
        self.front_loop_after_id = None
        self.front_loop_round = 0
        self.last_prepared_payload = {}
        self.first_event_progress_watch = {}
        self.task_monitor = {}
        self.json_view_mode_var = tk.StringVar(value="preview")

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

        top_row = tk.Frame(jitter_frame)
        top_row.grid(row=0, column=0, padx=(8, 8), pady=5, sticky="w")
        tk.Label(top_row, text="tip : 多重選取項目後雙擊title可批量設定").pack(
            side="left", padx=(0, 12)
        )
        tk.Button(
            top_row,
            text="UI 保存",
            command=self.save_ui_layout,
            bg="#fff4b3",
            activebackground="#ffe08a"
        ).pack(side="left", padx=(0, 4))
        tk.Button(
            top_row,
            text="提示",
            command=self.open_hint_note_dialog,
            bg="#fff4b3",
            activebackground="#ffe08a"
        ).pack(side="left", padx=(0, 4))
        tk.Button(top_row, text="糾正複製體", command=self.calculate_offsets_only).pack(
            side="left", padx=(0, 4)
        )
        tk.Button(top_row, text="改顏色", command=self.open_row_color_dialog).pack(
            side="left", padx=(0, 4)
        )
        self.restore_pre_run_btn = tk.Button(
            top_row,
            text="恢復執行前狀態",
            command=self.restore_pre_run_state,
            state="disabled"
        )
        self.restore_pre_run_btn.pack(side="left")
        tk.Label(
            jitter_frame,
            text="【註：buff_group 為負值時，請用「糾正複製體」重算；「套用偏移」僅手動平移時間。】"
        ).grid(row=1, column=0, padx=(8, 8), pady=(0, 4), sticky="w")
        action_container = tk.Frame(jitter_frame)
        action_container.grid(row=2, column=0, padx=(8, 8), pady=(0, 6), sticky="w")
        tk.Button(
            action_container,
            text="產生 randat",
            command=self.insert_randat_row,
            width=14
        ).pack(side="left", padx=(0, 6))
        action_row = tk.Frame(action_container)
        action_row.pack(side="left")
        tk.Button(
            action_row,
            text="at 交換",
            command=self.swap_selected_at,
            width=9
        ).pack(side="left", padx=(0, 6))
        tk.Button(
            action_row,
            text="PR 分析",
            command=self.analyze_pr_pairs,
            width=9
        ).pack(side="left", padx=(0, 6))
        tk.Label(action_row, text="minimum gap:").pack(side="left", padx=(2, 3))
        self.minimum_gap_entry = tk.Entry(action_row, width=7)
        self.minimum_gap_entry.insert(0, "0.050")
        self.minimum_gap_entry.pack(side="left", padx=(0, 6))
        tk.Button(
            action_row,
            text="套用 min gap",
            command=self.apply_minimum_gap_for_pairs,
            width=12
        ).pack(side="left", padx=(0, 0))

        right_content_paned = tk.PanedWindow(right_panel, orient=tk.VERTICAL, sashrelief=tk.RAISED)
        right_content_paned.pack(fill="both", expand=True)
        self.right_content_paned = right_content_paned

        table_panel = tk.Frame(right_content_paned)
        json_panel = tk.Frame(right_content_paned)
        right_content_paned.add(table_panel, minsize=220)
        right_content_paned.add(json_panel, minsize=160)

        columns = ("idx", "type", "button", "at", "at_jitter", "buff_group", "buff_cycle_sec", "buff_jitter_sec", "group")
        self.tree_columns = columns
        self.tree = ttk.Treeview(table_panel, columns=columns, show="headings", height=8, selectmode="extended")
        for col in columns:
            self.tree.heading(col, text=col)
            width = 92
            if col in ("buff_group", "buff_cycle_sec", "buff_jitter_sec"):
                width = 100
            self.tree.column(col, width=width, minwidth=40, stretch=False)
        self.tree.pack(fill="both", expand=True, pady=(0, 8))
        self.tree.tag_configure("bg_applied", background="#d8ecff")
        self.tree.tag_configure("bg_candidate", background="#fff4b3")
        self.tree.tag_configure("ring_running")
        self.tree.tag_configure("focus_hint")
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

        edit_row = tk.Frame(table_panel)
        edit_row.pack(fill="x", pady=(0, 8))
        tk.Button(edit_row, text="undo", command=self.undo_timeline, width=9).pack(side="left", padx=2)
        tk.Button(edit_row, text="redo", command=self.redo_timeline, width=9).pack(side="left", padx=2)
        tk.Button(edit_row, text="上移", command=self.move_selected_up, width=9).pack(side="left", padx=2)
        tk.Button(edit_row, text="下移", command=self.move_selected_down, width=9).pack(side="left", padx=2)
        tk.Button(edit_row, text="刪除列", command=self.delete_selected_rows, width=9).pack(side="left", padx=2)
        tk.Label(edit_row, text="自/手動偏移 :").pack(side="left", padx=(14, 5))
        self.offset_sec_entry = tk.Entry(edit_row, width=10)
        self.offset_sec_entry.insert(0, "{:.3f}".format(float(self.config.get("manual_offset_sec", NEGATIVE_GROUP_ANCHOR_GAP_SEC))))
        self.offset_sec_entry.pack(side="left", padx=(0, 5))
        tk.Label(edit_row, text="秒").pack(side="left", padx=(0, 5))
        tk.Button(edit_row, text="套用偏移", command=self.apply_offset_from_selected).pack(side="left", padx=2)

        mode_row = tk.Frame(json_panel)
        mode_row.pack(fill="x", pady=(0, 4))
        tk.Label(mode_row, text="JSON 顯示：").pack(side="left")
        for mode, label in (("preview", "Preview"), ("prepared", "Prepared"), ("runtime", "Runtime")):
            tk.Radiobutton(
                mode_row,
                text=label,
                value=mode,
                variable=self.json_view_mode_var,
                command=self.on_json_mode_change
            ).pack(side="left", padx=(0, 6))
        self.text = tk.Text(json_panel, height=12)
        self.text.pack(fill="both", expand=True)

        self.refresh_saved_list()
        self.restore_last_selected()
        self.on_json_mode_change()
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

    def get_current_right_paned_sash_y(self):
        self.root.update_idletasks()
        panes = self.right_content_paned.panes()
        if len(panes) < 2:
            return None
        _, sash_y = self.right_content_paned.sash_coord(0)
        return max(0, int(sash_y))

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
        self.root.update_idletasks()
        right_total = self.right_content_paned.winfo_height()
        if right_total <= 0:
            return
        right_sash_y = ui_layout.get("right_paned_sash_y")
        if isinstance(right_sash_y, (int, float)):
            target_y = int(right_sash_y)
        else:
            right_ratio = ui_layout.get("right_paned_ratio")
            if isinstance(right_ratio, (int, float)):
                target_y = int(right_total * right_ratio)
            else:
                return
        min_y = 120
        max_y = max(min_y, right_total - 120)
        self.right_content_paned.sash_place(0, 0, min(max_y, max(min_y, target_y)))

    def save_ui_layout(self):
        sash_x = self.get_current_paned_sash_x()
        right_sash_y = self.get_current_right_paned_sash_y()
        if sash_x is None or right_sash_y is None:
            self.show_warning("提醒", "目前無法取得 UI 版面資訊，請稍後再試")
            return

        column_widths = {}
        for col in self.tree_columns:
            width = self.tree.column(col, option="width")
            try:
                column_widths[col] = int(width)
            except Exception:
                pass

        body_width = max(1, int(self.body.winfo_width()))
        right_height = max(1, int(self.right_content_paned.winfo_height()))
        self.config["ui_layout"] = {
            "window_size": [self.root.winfo_width(), self.root.winfo_height()],
            "paned_sash_x": int(sash_x),
            "paned_ratio": round(float(sash_x) / float(body_width), 4),
            "right_paned_sash_y": int(right_sash_y),
            "right_paned_ratio": round(float(right_sash_y) / float(right_height), 4),
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
            self.show_warning("提醒", "請先選取列")
            return False
        if not color:
            self.show_warning("提醒", "請選擇有效顏色")
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
            self.show_warning("提醒", "請先選取列")
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

    def _open_pr_tables_dialog(self, title, sections):
        dialog = tk.Toplevel(self.root)
        dialog.title(title)
        dialog.geometry("1080x620")
        dialog.transient(self.root)
        dialog.grab_set()

        notebook = ttk.Notebook(dialog)
        notebook.pack(fill="both", expand=True, padx=8, pady=8)

        for section in sections:
            tab = ttk.Frame(notebook)
            notebook.add(tab, text=section["title"])
            columns = section.get("columns", [])
            rows = section.get("rows", [])

            tree = ttk.Treeview(tab, columns=columns, show="headings")
            for col in columns:
                tree.heading(col, text=col)
                tree.column(col, width=110, minwidth=60, stretch=True, anchor="w")

            y_scroll = ttk.Scrollbar(tab, orient="vertical", command=tree.yview)
            x_scroll = ttk.Scrollbar(tab, orient="horizontal", command=tree.xview)
            tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)

            tree.grid(row=0, column=0, sticky="nsew")
            y_scroll.grid(row=0, column=1, sticky="ns")
            x_scroll.grid(row=1, column=0, sticky="ew")
            tab.grid_rowconfigure(0, weight=1)
            tab.grid_columnconfigure(0, weight=1)

            for ridx, row in enumerate(rows):
                values = [str(row.get(col, "")) for col in columns]
                tree.insert("", "end", iid=str(ridx), values=values)

        btn_row = tk.Frame(dialog)
        btn_row.pack(fill="x", padx=8, pady=(0, 8))
        tk.Button(btn_row, text="關閉", command=dialog.destroy, width=10).pack(side="right")

    def on_json_mode_change(self):
        mode_var = getattr(self, "json_view_mode_var", None)
        mode = mode_var.get().strip().lower() if mode_var is not None else "preview"
        if mode == "runtime":
            self.render_runtime_analysis(force=True)
        elif mode == "prepared":
            self.render_prepared_payload(force=True)
        else:
            self.refresh_preview(force=True)

    def render_prepared_payload(self, force=False):
        mode_var = getattr(self, "json_view_mode_var", None)
        mode = mode_var.get().strip().lower() if mode_var is not None else "preview"
        if not force and mode != "prepared":
            return
        payload = self.last_prepared_payload if isinstance(self.last_prepared_payload, dict) else {}
        self.write_text(payload or {"prepared": "尚無資料"})

    def _build_preview_payload(self):
        sanitized = self._sanitize_events_for_backend(self.timeline)
        return {
            "current_name": self.current_name,
            "pi_host": self.config["pi_host"],
            "request_preview": self._build_start_task_payload(sanitized),
            "simultaneous_groups": build_overlap_summary(self.timeline)
        }

    def refresh_preview(self, force=False):
        mode_var = getattr(self, "json_view_mode_var", None)
        mode = mode_var.get().strip().lower() if mode_var is not None else "preview"
        if not force and mode != "preview":
            return
        self.write_text(self._build_preview_payload())

    def _round_trace_reason_label(self, reason):
        return ROUND_TRACE_REASON_LABELS.get(str(reason or "").strip(), "未知原因")

    def _format_group_label(self, group, buttons=None):
        group_text = str(group).strip()
        if not isinstance(buttons, list):
            buttons = []
        normalized_buttons = []
        for raw in buttons:
            btn = str(raw).strip()
            if btn and btn not in normalized_buttons:
                normalized_buttons.append(btn)
        if normalized_buttons:
            return "group {} ({})".format(group_text, "+".join(normalized_buttons))
        return "group {}".format(group_text)

    def _runtime_trace_group_meta(self, trace):
        group = str(trace.get("buffGroup", "")).strip()
        source_rows = self.runtime_working_timeline if self.runtime_working_timeline else self.timeline
        buttons = []
        for ev in source_rows:
            if not isinstance(ev, dict):
                continue
            if str(ev.get("buff_group", "")).strip() != group:
                continue
            btn = str(ev.get("button", "")).strip().lower()
            if btn and btn not in buttons:
                buttons.append(btn)
        return {
            "group": group,
            "group_buttons": buttons,
            "group_label": self._format_group_label(group, buttons)
        }

    def render_runtime_analysis(self, force=False):
        mode_var = getattr(self, "json_view_mode_var", None)
        mode = mode_var.get().strip().lower() if mode_var is not None else "preview"
        if not force and mode != "runtime":
            return
        runtime = self.timeline_runtime_info if isinstance(self.timeline_runtime_info, dict) else {}
        traces = self.runtime_round_traces if isinstance(self.runtime_round_traces, list) else []
        payload = {
            "run_id": runtime.get("run_id"),
            "state": runtime.get("state"),
            "processed_count": runtime.get("processed_count"),
            "events_total": runtime.get("events_total"),
            "round_traces": [],
            "move_gap_logs": list(getattr(self, "runtime_move_gap_logs", []) or [])[-20:]
        }
        lines = []
        if payload["move_gap_logs"]:
            lines.append("move | range(before->after) | keep(A/B) | preserved | compensation(req->applied)")
            lines.append("-" * 84)
            for item in payload["move_gap_logs"]:
                if not isinstance(item, dict):
                    continue
                lines.append(
                    "{:>4} | {} -> {} | {:>7} | {:>9} | {:+.4f} -> {:+.4f}".format(
                        str(item.get("direction", "")),
                        item.get("range_before", []),
                        item.get("range_after", []),
                        str(item.get("preserve_target", "-")),
                        str(bool(item.get("preserved", False))).lower(),
                        _safe_float(item.get("compensation_requested", 0.0)),
                        _safe_float(item.get("compensation_applied", 0.0))
                    )
                )
            lines.append("")
        lines.append("round | group_id | src_range -> dst_range | picked_slot | self_picked | color")
        lines.append("-" * 84)
        for trace in traces:
            if not isinstance(trace, dict):
                continue
            round_no = trace.get("round")
            group_meta = self._runtime_trace_group_meta(trace)
            group = group_meta["group"]
            group_label = group_meta["group_label"]
            idx = trace.get("pickedIdx")
            anchor_idx = trace.get("anchorIdx")
            candidate_list = trace.get("candidateIdxList", [])
            free_candidate_list = trace.get("freeCandidateIdxList", [])
            picked_pos = trace.get("pickedCandidatePos")
            dice_value = trace.get("diceValue")
            result = str(trace.get("result", "")).strip().lower()
            reason = trace.get("pickedReason", trace.get("reason"))
            src_range = trace.get("src_range", [])
            dst_range = trace.get("dst_range", [])
            picked_slot = trace.get("picked_slot", trace.get("pickedIdx"))
            self_picked = bool(trace.get("self_picked", False))
            color = str(trace.get("color", "")).strip().lower()
            trace_with_group_meta = {
                "round": round_no,
                "group_id": group,
                "group_label": group_label,
                "src_range": src_range,
                "dst_range": dst_range,
                "picked_slot": picked_slot,
                "self_picked": bool(self_picked),
                "color": color,
                "picked_reason": reason,
                "dice_plan": {
                    "candidate_idx_list": candidate_list,
                    "free_candidate_idx_list": free_candidate_list,
                    "picked_candidate_pos": picked_pos,
                    "dice_value": dice_value
                }
            }
            payload["round_traces"].append(trace_with_group_meta)
            lines.append(
                "{:>5} | {:>8} | {} -> {} | {:>11} | {:>11} | {}".format(
                    str(round_no),
                    group,
                    src_range,
                    dst_range,
                    str(picked_slot),
                    str(self_picked).lower(),
                    color
                )
            )
            lines.append("      {}，原因：{}".format(group_label, self._round_trace_reason_label(reason)))

        self.text.delete("1.0", tk.END)
        if lines:
            summary_text = "\n".join(lines)
        else:
            summary_text = "暫無 runtime 回合資料"
        self.text.insert(tk.END, summary_text)
        self.text.insert(tk.END, "\n\n---\n詳細 JSON\n")
        self.text.insert(tk.END, json.dumps(payload, ensure_ascii=False, indent=2))

    def update_runtime_from_status(self, status_response):
        runtime = status_response.get("timeline_runtime", {})
        if not isinstance(runtime, dict):
            runtime = {}
        events = runtime.get("events", [])
        if not isinstance(events, list):
            events = []
        round_traces = runtime.get("round_traces", [])
        if not isinstance(round_traces, list):
            round_traces = []

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
                "round_trace_count": len(round_traces),
                "round_trace_last": round_traces[-1] if round_traces else None,
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
        self.runtime_round_traces = round_traces
        self.runtime_recent_ok_indices = recent_ok
        self.runtime_recent_skipped_indices = recent_skipped
        self.runtime_latest_index = latest_idx
        self.last_runtime_signature = signature
        return changed

    def poll_runtime_status(self):
        try:
            if not self.offline_mode:
                res = self.request_pi({"action": "status"}, write_response=False, channel="status")
                if isinstance(res, dict):
                    changed = self.update_runtime_from_status(res)
                    state = str(self.timeline_runtime_info.get("state", "")).strip().lower()
                    self._check_first_event_progress_timeout(state)
                    self._monitor_task_status_progress(
                        state,
                        self.timeline_runtime_info.get("processed_count", 0)
                    )
                    if state == "running":
                        self.runtime_manual_restore_active = False
                        self.runtime_display_frozen = False
                        if self.runtime_working_timeline:
                            self._show_runtime_view_timeline()
                    elif state in ("stopped", "idle"):
                        if self.has_pre_run_snapshot and not self.runtime_manual_restore_active:
                            self._restore_pre_run_snapshot_after_runtime()
                        self.runtime_display_frozen = False
                        self._reset_task_monitor()
                    else:
                        self.runtime_display_frozen = False
                    if changed:
                        self.render_runtime_analysis()
                        if state == "running":
                            self.refresh_tree()
                            self.focus_latest_runtime_row()
        except Exception as e:
            kind = self._classify_request_error(e)
            self._set_channel_error("status", "[status:{}] {}".format(kind, str(e)))
        finally:
            self.root.after(RUNTIME_POLL_INTERVAL_MS, self.poll_runtime_status)

    def clear_runtime_highlight(self):
        self.timeline_runtime_info = {"events": []}
        self.timeline_runtime_by_index = {}
        self.runtime_round_traces = []
        self.runtime_recent_ok_indices = []
        self.runtime_recent_skipped_indices = []
        self.runtime_latest_index = None
        self.last_runtime_signature = ""
        self.runtime_move_gap_logs = []
        self._reset_first_event_progress_watch()
        self._reset_task_monitor()

    def _reset_first_event_progress_watch(self):
        self.first_event_progress_watch = {}

    def _reset_task_monitor(self):
        self.task_monitor = {}

    def _arm_task_monitor_wait_ack(self, client_task_id):
        self.task_monitor = {
            "phase": "wait_ack",
            "client_task_id": str(client_task_id or "").strip(),
            "server_task_id": "",
            "phase_started_at": float(time.monotonic()),
            "last_state": "submitted",
            "last_processed_count": 0,
            "last_progress_at": None,
            "failed": False
        }

    def _mark_task_monitor_failed(self, phase, elapsed_ms, threshold_ms, last_state=""):
        monitor_obj = getattr(self, "task_monitor", {})
        monitor = monitor_obj if isinstance(monitor_obj, dict) else {}
        server_task_id = str(monitor.get("server_task_id", "")).strip()
        self.task_monitor = {
            "phase": str(phase or "").strip().lower(),
            "client_task_id": str(monitor.get("client_task_id", "")).strip(),
            "server_task_id": server_task_id,
            "phase_started_at": monitor.get("phase_started_at"),
            "last_state": str(last_state or monitor.get("last_state", "")).strip(),
            "last_processed_count": int(monitor.get("last_processed_count", 0) or 0),
            "last_progress_at": monitor.get("last_progress_at"),
            "failed": True
        }
        message = (
            "任務監控逾時（{}）\nserver_task_id={}\nlast_state={}\nelapsed_ms={}\nthreshold_ms={}".format(
                str(phase or "").upper(),
                server_task_id or "unknown",
                self.task_monitor["last_state"] or "unknown",
                int(max(0, elapsed_ms)),
                int(max(0, threshold_ms))
            )
        )
        self.set_frontend_error(message)

    def _monitor_after_ack(self, ack_response):
        monitor_obj = getattr(self, "task_monitor", {})
        monitor = monitor_obj if isinstance(monitor_obj, dict) else {}
        if not monitor:
            return
        server_task_id = str(ack_response.get("server_task_id", "")).strip() if isinstance(ack_response, dict) else ""
        now = float(time.monotonic())
        monitor["phase"] = "wait_running"
        monitor["server_task_id"] = server_task_id
        monitor["phase_started_at"] = now
        monitor["last_state"] = "accepted"
        monitor["last_processed_count"] = 0
        monitor["last_progress_at"] = None
        monitor["failed"] = False
        self.task_monitor = monitor

    def _monitor_task_status_progress(self, runtime_state, processed_count):
        monitor_obj = getattr(self, "task_monitor", {})
        monitor = monitor_obj if isinstance(monitor_obj, dict) else {}
        if not monitor or monitor.get("failed"):
            return
        phase = str(monitor.get("phase", "")).strip().lower()
        if phase not in {"wait_running", "watch_progress"}:
            return

        now = float(time.monotonic())
        state = str(runtime_state or "").strip().lower()
        if state:
            monitor["last_state"] = state

        if phase == "wait_running":
            started = float(monitor.get("phase_started_at", now))
            if state == "running":
                monitor["phase"] = "watch_progress"
                monitor["phase_started_at"] = now
                monitor["last_progress_at"] = now
                monitor["last_processed_count"] = int(max(0, processed_count))
                self.task_monitor = monitor
                return
            elapsed_ms = int((now - started) * 1000.0)
            if elapsed_ms >= START_TIMEOUT_MS:
                self._mark_task_monitor_failed(
                    phase="START_TIMEOUT",
                    elapsed_ms=elapsed_ms,
                    threshold_ms=START_TIMEOUT_MS,
                    last_state=state or "accepted"
                )
                return
            self.task_monitor = monitor
            return

        if state in {"done", "stopped", "idle", "error"}:
            self._reset_task_monitor()
            return

        last_progress_at = monitor.get("last_progress_at")
        if last_progress_at is None:
            last_progress_at = now
        last_processed = int(monitor.get("last_processed_count", 0) or 0)
        current_processed = int(max(0, processed_count))
        if current_processed > last_processed:
            monitor["last_processed_count"] = current_processed
            monitor["last_progress_at"] = now
            self.task_monitor = monitor
            return

        stall_ms = int((now - float(last_progress_at)) * 1000.0)
        if stall_ms >= PROGRESS_STALL_MS:
            self._mark_task_monitor_failed(
                phase="PROGRESS_STALL",
                elapsed_ms=stall_ms,
                threshold_ms=PROGRESS_STALL_MS,
                last_state=state or "running"
            )
            return
        self.task_monitor = monitor

    def _extract_first_event_timing(self, prepared_events):
        first_event_at = None
        first_event_jitter = 0.0
        for ev in prepared_events or []:
            if not isinstance(ev, dict):
                continue
            try:
                at_value = float(ev.get("at", 0.0))
            except Exception:
                continue
            if first_event_at is None or at_value < first_event_at:
                first_event_at = at_value
                try:
                    first_event_jitter = max(0.0, float(ev.get("at_jitter", 0.0)))
                except Exception:
                    first_event_jitter = 0.0

        if first_event_at is None:
            return None
        return {
            "first_event_at": float(first_event_at),
            "first_event_jitter": float(first_event_jitter)
        }

    def _sanitize_events_for_backend(self, events):
        sanitized = []
        for ev in events or []:
            if not isinstance(ev, dict):
                continue
            row = {}
            for key, value in ev.items():
                text = str(key or "")
                if text.startswith("runtime_") or text.startswith("__"):
                    continue
                if text in {"row_color", "replicatedRow"}:
                    continue
                if text in {"round", "runtime_meta"}:
                    continue
                row[text] = value
            sanitized.append(row)
        return sanitized

    def _next_client_task_id(self):
        return "ct-{}-{}".format(int(time.time() * 1000), uuid.uuid4().hex[:10])

    def _build_start_task_payload(self, backend_events):
        sent_at_ms = int(time.time() * 1000)
        skip_mode = self.config.get("buff_skip_mode", BUFF_SKIP_MODE_COMPRESS)
        timeline = []
        for idx, ev in enumerate(backend_events or []):
            try:
                at_ms = int(round(max(0.0, float(ev.get("at", 0.0))) * 1000.0))
            except Exception:
                at_ms = 0
            timeline.append({
                "idx": int(idx),
                "at_ms": at_ms,
                "action": str(ev.get("type", "")).strip().lower(),
                "btn": str(ev.get("button", "")).strip().lower(),
                "skip_mode": skip_mode
            })
        return {
            "type": "start_task",
            "contract_version": RUNTIME_CONTRACT_VERSION,
            "client_task_id": self._next_client_task_id(),
            "sent_at_ms": sent_at_ms,
            "timeline": timeline
        }

    def _arm_first_event_progress_watch(self, sent_at_monotonic, send_delay_sec, first_event_timing):
        if not first_event_timing:
            self._reset_first_event_progress_watch()
            return
        first_event_at = float(first_event_timing["first_event_at"])
        first_event_jitter = float(first_event_timing["first_event_jitter"])

        grace_sec = FIRST_EVENT_DEADLINE_GRACE_SEC + first_event_jitter
        self.first_event_progress_watch = {
            "sent_at_monotonic": float(sent_at_monotonic),
            "send_delay_sec": float(send_delay_sec or 0.0),
            "first_event_at": float(first_event_at),
            "first_event_jitter": float(first_event_jitter),
            "grace_sec": float(grace_sec),
            "expected_first_event_deadline": float(sent_at_monotonic) + float(first_event_at) + float(grace_sec),
            "timeout_reported": False
        }

    def _check_first_event_progress_timeout(self, runtime_state):
        watch_obj = getattr(self, "first_event_progress_watch", {})
        watch = watch_obj if isinstance(watch_obj, dict) else {}
        if not watch:
            return

        state = str(runtime_state or "").strip().lower()
        if state in {"stopped", "idle", "error"}:
            return

        processed_count_raw = self.timeline_runtime_info.get("processed_count", 0)
        try:
            processed_count = int(processed_count_raw)
        except Exception:
            processed_count = 0
        if processed_count > 0:
            return

        deadline = watch.get("expected_first_event_deadline")
        try:
            deadline = float(deadline)
        except Exception:
            return

        if time.monotonic() < deadline:
            return
        if watch.get("timeout_reported"):
            return

        msg = "已收到執行請求，但超過預期首事件時間仍無執行紀錄"
        self.set_frontend_error(msg)
        self.set_status(msg)
        watch["timeout_reported"] = True
        self.first_event_progress_watch = watch

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

    def _runtime_landing_map(self):
        runtime_events = self.timeline_runtime_info.get("events", [])
        if not isinstance(runtime_events, list):
            runtime_events = []
        landing = {}
        for ev in reversed(runtime_events):
            if not isinstance(ev, dict):
                continue
            group = str(ev.get("buff_group", "")).strip()
            if not group or group in landing:
                continue
            landed_idx = ev.get("runtime_landed_index")
            anchor_idx = ev.get("runtime_anchor_index")
            if landed_idx is None or anchor_idx is None:
                continue
            landing[group] = {
                "landed_index": int(landed_idx),
                "anchor_index": int(anchor_idx),
                "occupies_original": bool(int(ev.get("runtime_occupies_original", 0))),
                "self_picked": bool(int(ev.get("runtime_self_picked", ev.get("runtime_occupies_original", 0)))),
                "color": str(ev.get("runtime_group_color", "")).strip().lower()
            }
        return landing

    def _runtime_cooldown_by_row(self):
        runtime = self.timeline_runtime_info if isinstance(self.timeline_runtime_info, dict) else {}
        cooldowns = runtime.get("cooldowns", [])
        if not isinstance(cooldowns, list):
            cooldowns = []
        by_row = {}
        for item in cooldowns:
            if not isinstance(item, dict):
                continue
            landed_index = item.get("landed_index")
            if landed_index is None:
                continue
            try:
                row_idx = int(landed_index)
            except Exception:
                continue
            by_row[row_idx] = {
                "buff_group": str(item.get("buff_group", "")).strip(),
                "remain_sec": max(0.0, _safe_float(item.get("remain_sec", 0.0)))
            }
        return by_row

    def refresh_tree(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        summary = build_overlap_summary(self.timeline)
        event_to_group = {}

        for grp in summary:
            group_id = grp["group"]
            group_label = grp.get("group_label", self._format_group_label(group_id, grp.get("group_buttons", [])))
            for ev in grp["events"]:
                key = (ev["type"], ev["button"], ev["at"])
                event_to_group[key] = {
                    "group": group_id,
                    "group_label": group_label
                }

        cooldown_by_row = self._runtime_cooldown_by_row()
        runtime_landing = self._runtime_landing_map()
        runtime_state = str(self.timeline_runtime_info.get("state", "")).strip().lower()
        runtime_ok_rank = {
            idx: rank + 1
            for rank, idx in enumerate(self.runtime_recent_ok_indices[:3])
        }
        for i, ev in enumerate(self.timeline):
            self._sync_replicated_row(ev)
            key = (ev["type"], ev["button"], ev["at"])
            grp_meta = event_to_group.get(key, {})
            if isinstance(grp_meta, dict):
                grp = grp_meta.get("group", "")
                grp_display = grp_meta.get("group_label", grp)
            else:
                grp = grp_meta
                grp_display = grp
            original_buff_group = str(ev.get("buff_group", "")).strip()
            buff_group = original_buff_group
            runtime_event = self.timeline_runtime_by_index.get(i, {})
            runtime_status = str(runtime_event.get("status", "")).strip().lower()
            cooldown_info = cooldown_by_row.get(i)
            if cooldown_info:
                remain = cooldown_info["remain_sec"]
                group_text = cooldown_info["buff_group"] or original_buff_group
                buff_group = "{} 冷卻中({:.1f}秒)".format(group_text, remain)
            elif runtime_status == "skipped_by_cooldown":
                buff_group = "{}（冷卻中）".format(original_buff_group) if original_buff_group else "冷卻中"
            at_value = "{:.2f}".format(float(ev["at"]))
            buff_cycle_value = ev.get("buff_cycle_sec", 0.0)
            buff_cycle_display = buff_cycle_value
            if runtime_state == "running":
                base_cycle = _safe_float(ev.get("runtime_buff_cycle_base", buff_cycle_value))
                applied_cycle = _safe_float(ev.get("runtime_buff_cycle_applied", buff_cycle_value))
                if base_cycle > 0.0 and abs(applied_cycle - base_cycle) > 1e-9:
                    buff_cycle_display = "{:.4f} ({:.4f})".format(base_cycle, applied_cycle)
                else:
                    buff_cycle_display = "{:.4f}".format(applied_cycle)
            tags = []
            row_type = str(ev.get("type", "")).strip().lower()
            if row_type == "randat":
                at_value = "-"
            is_candidate = False
            is_applied = False
            if runtime_state != "running":
                is_candidate = bool(original_buff_group)
            else:
                if original_buff_group:
                    landing = runtime_landing.get(original_buff_group, {})
                    self_picked = bool(landing.get("self_picked"))
                    color = str(landing.get("color", "")).strip().lower()
                    if color == "light_yellow":
                        is_candidate = True
                        is_applied = False
                    elif color == "light_blue":
                        is_candidate = False
                        is_applied = True
                    else:
                        is_candidate = self_picked
                        is_applied = bool(landing) and not self_picked
            is_running = runtime_state == "running" and i == self.runtime_latest_index
            is_focus = i == self.runtime_latest_index
            tags.extend(get_buff_cell_visual_state(
                is_candidate=is_candidate,
                is_applied=is_applied,
                is_running=is_running,
                is_focus=is_focus
            ))
            ok_rank = runtime_ok_rank.get(i)
            if ok_rank and "bg_applied" not in tags and "bg_candidate" not in tags:
                tags.append("runtime_ok_{}".format(min(3, int(ok_rank))))
            row_color_tag = self._ensure_tree_color_tag(ev.get("row_color", ""))
            if row_color_tag and "bg_applied" not in tags and "bg_candidate" not in tags:
                tags.append(row_color_tag)
            if row_type != "randat" and str(ev.get("button", "")).strip().lower() not in SUPPORTED_BUTTONS:
                tags.append("unsupported_button")
            self.tree.insert("", "end", iid=str(i), values=(
                i,
                ev["type"],
                ev["button"],
                at_value,
                ev["at_jitter"],
                buff_group,
                buff_cycle_display,
                ev.get("buff_jitter_sec", 0.0),
                grp_display
            ), tags=tuple(tags))

    def _is_runtime_readonly(self):
        state = str(self.timeline_runtime_info.get("state", "")).strip().lower()
        return state == "running" or bool(self.runtime_display_frozen)

    def _ensure_runtime_editable(self):
        if not self.has_pre_run_snapshot and self.runtime_display_frozen:
            self.runtime_display_frozen = False
        if not self._is_runtime_readonly():
            return True
        self.show_warning("提醒", "執行中或停止後凍結中，請先「恢復執行前狀態」再編輯")
        return False

    def _set_restore_pre_run_button_state(self):
        btn = getattr(self, "restore_pre_run_btn", None)
        if btn is None:
            return
        btn.config(state=("normal" if self.has_pre_run_snapshot else "disabled"))

    def restore_pre_run_state(self):
        if not self.has_pre_run_snapshot:
            self.show_warning("提醒", "沒有可恢復內容")
            return
        before = self._begin_timeline_change()
        self.timeline = self.copy_events(self.pre_run_timeline_snapshot)
        self._finalize_timeline_change(before)
        self.runtime_manual_restore_active = True
        self.runtime_display_frozen = False
        self.clear_runtime_highlight()
        self.mark_timeline_dirty()
        self.set_status("已恢復執行前狀態，可使用 Ctrl+Z undo")

    def _show_runtime_view_timeline(self):
        if not self.runtime_working_timeline:
            return
        self.timeline = self.copy_events(self.runtime_working_timeline)
        self.refresh_tree()
        self.refresh_preview()

    def _restore_pre_run_snapshot_after_runtime(self):
        if not self.has_pre_run_snapshot:
            return
        self.timeline = self.copy_events(self.pre_run_timeline_snapshot)
        self.runtime_display_frozen = False
        self.runtime_manual_restore_active = False
        self.has_pre_run_snapshot = False
        self.pre_run_timeline_snapshot = []
        self.runtime_working_timeline = []
        self._set_restore_pre_run_button_state()
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
            self.show_warning("提醒", "Pi IP 不可空白")
            return
        try:
            delay = self.parse_send_delay_sec()
        except ValueError as e:
            self.set_frontend_error(str(e))
            self.show_warning("提醒", str(e))
            return

        self.config["pi_host"] = ip
        self.config["send_delay_sec"] = delay
        save_config(self.config)
        self.update_current_labels()
        self.set_frontend_error("")
        self.set_status("已保存 Pi IP：{}，送出延遲 {} 秒".format(ip, delay))
        self.set_connected(False, "設定已更新，請重新測試連線")

    def apply_send_delay_if_needed(self, on_ready, delay_override=None):
        delay = float(delay_override) if delay_override is not None else self.parse_send_delay_sec()
        self.config["send_delay_sec"] = delay
        save_config(self.config)

        now_monotonic = time.monotonic()
        scheduled_send_time_monotonic = now_monotonic + max(0.0, delay)
        if delay <= 0:
            actual_send_time_monotonic = now_monotonic
            send_error_ms = (actual_send_time_monotonic - scheduled_send_time_monotonic) * 1000.0
            on_ready({
                "send_delay_sec": 0.0,
                "scheduled_send_time_monotonic": scheduled_send_time_monotonic,
                "actual_send_time_monotonic": actual_send_time_monotonic,
                "send_error_ms": send_error_ms
            })
            return

        self.set_status("延遲 {} 秒後送出...".format(delay))

        def poll_delay_ready():
            now = time.monotonic()
            if now >= scheduled_send_time_monotonic:
                actual_send_time_monotonic = now
                send_error_ms = (actual_send_time_monotonic - scheduled_send_time_monotonic) * 1000.0
                on_ready({
                    "send_delay_sec": delay,
                    "scheduled_send_time_monotonic": scheduled_send_time_monotonic,
                    "actual_send_time_monotonic": actual_send_time_monotonic,
                    "send_error_ms": send_error_ms
                })
                return
            remaining_ms = int(max(10.0, min(50.0, (scheduled_send_time_monotonic - now) * 1000.0)))
            self.root.after(remaining_ms, poll_delay_ready)

        self.root.after(10, poll_delay_ready)

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
                ans = self.confirm_cancel(
                    "重新分析來源",
                    "目前無新錄製資料。\n是=還原初版 Timeline；否=還原上次保存新版 Timeline；取消=不變更。"
                )
                if ans is None:
                    self.set_status("已取消重新分析")
                    return
                if ans:
                    self.restore_original_timeline()
                    self.set_status("目前無新錄製資料，已還原初版 Timeline（重新分析）")
                else:
                    restored = self.restore_latest_saved_timeline()
                    if restored:
                        self.set_status("目前無新錄製資料，已還原上次保存新版 Timeline（重新分析）")
                return
            if has_original:
                self.restore_original_timeline()
                self.set_status("目前無新錄製資料，已還原初版 Timeline（重新分析）")
                return
            if has_latest_saved:
                self.restore_latest_saved_timeline()
                self.show_info("重新分析", "目前無新錄製資料，已還原上次保存新版 Timeline")
                self.set_status("目前無新錄製資料，已還原上次保存新版 Timeline（重新分析）")
                return
            self.set_status("目前沒有錄到資料，也沒有可還原的初版/新版 Timeline")
            return

        before = self._begin_timeline_change()
        self.timeline = build_timeline(events, recording_start)
        self._finalize_timeline_change(before)
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
        if not hasattr(self.root, "tk"):
            ans = messagebox.askyesnocancel(
                "保存同名 Timeline",
                "名稱 '{}' 已存在。\n是=完全取代，否=存為新版，取消=放棄。".format(name)
            )
            if ans is None:
                return None
            return "full_replace" if ans else "new_version"

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
        self._position_dialog_to_app_lower_center(dialog)
        self.root.wait_window(dialog)
        return result["mode"]

    def save_current_timeline(self):
        if not self.timeline:
            self.show_warning("提醒", "目前沒有 timeline 可保存")
            return

        name = self.name_entry.get().strip()
        if not name:
            self.show_warning("提醒", "請先輸入保存名稱")
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
            self.show_error("保存失敗", str(e))
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
            self.show_warning("提醒", "請先從清單選一個已保存項目")
            return

        if self._is_script_switch_for_load(name):
            confirmed = self.confirm(
                "切換腳本確認",
                "將清空目前 SESSION undo/redo，是否繼續？"
            )
            if not confirmed:
                self.set_status("已取消載入")
                return

        try:
            data = load_named_timeline(name)
        except Exception as e:
            self.show_error("載入失敗", str(e))
            return
        target_name = data.get("name", name)
        is_script_switch = self._is_script_switch_for_load(target_name)
        before = self._begin_timeline_change()

        self.timeline = [self.normalize_event_schema(ev) for ev in data.get("events", [])]
        self.timeline_meta = self.normalize_meta(data.get("_meta", {}))
        if not self.timeline_meta.get("original_events") and self.timeline:
            self.timeline_meta["original_events"] = self.copy_events(self.timeline)
        self.current_name = target_name
        if is_script_switch:
            self._reset_timeline_history()
        else:
            self._finalize_timeline_change(before)
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
            self.show_warning("提醒", "請先選取要刪除的項目")
            return

        yes = self.confirm("確認刪除", "確定要刪除 '{}' 嗎？".format(name))
        if not yes:
            return

        try:
            delete_named_timeline(name)
        except Exception as e:
            self.show_error("刪除失敗", str(e))
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
            self.show_warning("提醒", "請先從清單選一個已保存項目")
            return

        new_name = sanitize_filename(self.name_entry.get().strip())
        if not new_name:
            self.show_warning("提醒", "請先輸入新的腳本名稱")
            return

        if new_name == old_name:
            self.set_status("名稱未變更：{}".format(old_name))
            return

        overwrite = False
        if os.path.exists(timeline_file_path(new_name)):
            overwrite = self.confirm("確認覆蓋", "名稱 '{}' 已存在，是否覆蓋？".format(new_name))
            if not overwrite:
                return

        try:
            rename_named_timeline(old_name, new_name, overwrite=overwrite)
        except Exception as e:
            self.show_error("重新命名失敗", str(e))
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
            self.show_warning("提醒", "目前沒有 timeline 資料")
            return

        selected = self.get_selected_indexes()
        if not selected:
            self.show_warning("提醒", "請先選取一列或多列")
            return

        try:
            jitter = float(self.at_jitter_entry.get().strip())
        except ValueError:
            self.show_error("錯誤", "at jitter 必須是數字")
            return

        before = self._begin_timeline_change()
        for idx in selected:
            self.timeline[idx]["at_jitter"] = jitter

        self._finalize_timeline_change(before)
        self.mark_timeline_dirty()
        for idx in selected:
            self.tree.selection_add(str(idx))
        self.set_status("已套用 jitter 到 {} 筆選取列".format(len(selected)))
        self._show_jitter_order_risk_reminder("套用 jitter 到選取列")

    def apply_jitter_to_all(self):
        if not self.timeline:
            self.show_warning("提醒", "目前沒有 timeline 資料")
            return

        try:
            jitter = float(self.at_jitter_entry.get().strip())
        except ValueError:
            self.show_error("錯誤", "at jitter 必須是數字")
            return

        before = self._begin_timeline_change()
        for ev in self.timeline:
            ev["at_jitter"] = jitter

        self._finalize_timeline_change(before)
        self.mark_timeline_dirty()
        self.set_status("已套用 jitter 到全部 event")
        self._show_jitter_order_risk_reminder("套用 jitter 到全部 event")

    def clear_jitter_selected(self):
        if not self.timeline:
            self.show_warning("提醒", "目前沒有 timeline 資料")
            return

        selected = self.get_selected_indexes()
        if not selected:
            self.show_warning("提醒", "請先選取一列或多列")
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
            self.show_warning("提醒", "目前沒有 timeline 資料")
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
            self.show_warning("提醒", "目前沒有 timeline 資料")
            return

        selected = sorted(self.get_selected_indexes())
        if not selected:
            self.show_warning("提醒", "請先選取一列或多列")
            return
        if any(self._is_negative_buff_group(self.timeline[idx].get("buff_group", "")) for idx in selected):
            self.show_warning("提醒", "選到複製體負群（buff_group < 0），請改用「糾正複製體」。")
            return

        try:
            offset_sec = self.get_manual_offset_sec()
        except Exception as e:
            self.show_error("錯誤", str(e))
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
                self.show_error("Pi 連線失敗", str(e))

    def stop_pi(self):
        try:
            self.front_loop_enabled = False
            after_id = getattr(self, "front_loop_after_id", None)
            if after_id:
                try:
                    self.root.after_cancel(after_id)
                except Exception:
                    pass
                self.front_loop_after_id = None
            self.config["pi_host"] = self.pi_ip_entry.get().strip() or DEFAULT_PI_HOST
            save_config(self.config)
            self.update_current_labels()

            self.set_frontend_error("")
            self._reset_first_event_progress_watch()
            self.request_pi({"action": "stop"})
            self._restore_pre_run_snapshot_after_runtime()
            self.set_status("已停止 Pi：{}".format(self.config["pi_host"]))
        except Exception as e:
            self.set_frontend_error(str(e))
            self.show_error("停止失敗", str(e))

    def send_timeline(self):
        if not self.timeline:
            self.show_warning("提醒", "請先錄製並分析，或載入已保存項目")
            return
        self.front_loop_enabled = False
        if self.front_loop_after_id:
            try:
                self.root.after_cancel(self.front_loop_after_id)
            except Exception:
                pass
            self.front_loop_after_id = None
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
            self.show_warning("提醒", str(e))
            return
        save_config(self.config)
        self.update_current_labels()

        prepared_events, resolve_note = self.prepare_events_for_send(action_reason="before_send")
        if prepared_events is None:
            return
        self.runtime_working_timeline = self.copy_events(prepared_events)
        self._show_runtime_view_timeline()

        backend_events = self._sanitize_events_for_backend(prepared_events)
        payload = self._build_start_task_payload(backend_events)
        self._arm_task_monitor_wait_ack(payload.get("client_task_id"))

        display_name = self.current_name if self.current_name else "未命名資料"

        try:
            self.set_frontend_error("")
            self.clear_runtime_highlight()
            self.refresh_tree()
            def do_send(delay_meta):
                try:
                    res = self.request_pi(
                        payload,
                        write_response=False,
                        timeout=max(0.2, ACK_TIMEOUT_MS / 1000.0)
                    )
                    sent_at_monotonic = delay_meta["actual_send_time_monotonic"]
                    first_event_timing = self._extract_first_event_timing(backend_events)
                    self.write_text({
                        "sending_name": display_name,
                        "pi_host": self.config["pi_host"],
                        "send_delay_sec": delay_meta["send_delay_sec"],
                        "scheduled_send_time_monotonic": delay_meta["scheduled_send_time_monotonic"],
                        "actual_send_time_monotonic": delay_meta["actual_send_time_monotonic"],
                        "send_error_ms": delay_meta["send_error_ms"],
                        "sent_at_monotonic": sent_at_monotonic,
                        "first_event_at": first_event_timing["first_event_at"] if first_event_timing else None,
                        "request": payload,
                        "response": res
                    })

                    if res.get("status") == "stopped":
                        self._reset_task_monitor()
                        self._reset_first_event_progress_watch()
                        self._restore_pre_run_snapshot_after_runtime()
                        self.set_status("Pi 已停止執行：{} -> {}".format(display_name, self.config["pi_host"]))
                    elif self._is_contract_version_mismatch(res):
                        self._reset_task_monitor()
                        self._reset_first_event_progress_watch()
                        self._restore_pre_run_snapshot_after_runtime()
                        self.show_error("版本不相容", self._build_contract_version_mismatch_message(res))
                    elif res.get("status") in ("error", "busy"):
                        self._reset_task_monitor()
                        self._reset_first_event_progress_watch()
                        self._restore_pre_run_snapshot_after_runtime()
                        self.set_status("送出失敗：{} -> {}".format(display_name, self.config["pi_host"]))
                    else:
                        self._monitor_after_ack(res)
                        self._arm_first_event_progress_watch(sent_at_monotonic, delay_meta["send_delay_sec"], first_event_timing)
                        suffix = ""
                        if resolve_note:
                            suffix = "（{}）".format(resolve_note)
                        self.set_status("已送出：{} -> {}{}".format(display_name, self.config["pi_host"], suffix))
                except Exception as e:
                    if isinstance(e, TimeoutError):
                        self._mark_task_monitor_failed(
                            phase="ACK_TIMEOUT",
                            elapsed_ms=ACK_TIMEOUT_MS,
                            threshold_ms=ACK_TIMEOUT_MS,
                            last_state="submitted"
                        )
                    else:
                        self._reset_task_monitor()
                        self.set_frontend_error(str(e))
                    self._restore_pre_run_snapshot_after_runtime()
                    self.show_error("傳送失敗", str(e))

            self.apply_send_delay_if_needed(do_send)
        except Exception as e:
            self.set_frontend_error(str(e))
            self._restore_pre_run_snapshot_after_runtime()
            self.show_error("傳送失敗", str(e))

    def send_timeline_loop(self):
        if not self.timeline:
            self.show_warning("提醒", "請先錄製並分析，或載入已保存項目")
            return
        if self.front_loop_enabled:
            self.show_warning("提醒", "前端重複送出已在執行中")
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
            self.show_warning("提醒", str(e))
            return
        save_config(self.config)
        self.update_current_labels()

        display_name = self.current_name if self.current_name else "未命名資料"
        self.front_loop_enabled = True
        self.front_loop_round = 0
        self.set_status("已啟用前端重複送出：{}".format(display_name))
        self._dispatch_front_loop_once(display_name)

    def _dispatch_front_loop_once(self, display_name):
        if not self.front_loop_enabled:
            return

        prepared_events, resolve_note = self.prepare_events_for_send(action_reason="before_send_loop")
        if prepared_events is None:
            self.front_loop_enabled = False
            return
        self.runtime_working_timeline = self.copy_events(prepared_events)
        self._show_runtime_view_timeline()

        backend_events = self._sanitize_events_for_backend(prepared_events)
        payload = self._build_start_task_payload(backend_events)
        self._arm_task_monitor_wait_ack(payload.get("client_task_id"))

        try:
            self.set_frontend_error("")
            self.clear_runtime_highlight()
            self.refresh_tree()
            def do_send(delay_meta):
                if not self.front_loop_enabled:
                    return
                try:
                    res = self.request_pi(
                        payload,
                        write_response=False,
                        timeout=max(0.2, ACK_TIMEOUT_MS / 1000.0)
                    )
                    sent_at_monotonic = delay_meta["actual_send_time_monotonic"]
                    first_event_timing = self._extract_first_event_timing(backend_events)
                    self.write_text({
                        "sending_name": display_name,
                        "pi_host": self.config["pi_host"],
                        "send_delay_sec": delay_meta["send_delay_sec"],
                        "scheduled_send_time_monotonic": delay_meta["scheduled_send_time_monotonic"],
                        "actual_send_time_monotonic": delay_meta["actual_send_time_monotonic"],
                        "send_error_ms": delay_meta["send_error_ms"],
                        "sent_at_monotonic": sent_at_monotonic,
                        "first_event_at": first_event_timing["first_event_at"] if first_event_timing else None,
                        "request": payload,
                        "response": res
                    })
                    if self._is_contract_version_mismatch(res):
                        self._reset_task_monitor()
                        self._reset_first_event_progress_watch()
                        self._restore_pre_run_snapshot_after_runtime()
                        self.front_loop_enabled = False
                        self.show_error("版本不相容", self._build_contract_version_mismatch_message(res))
                        return
                    if res.get("status") in ("error", "busy"):
                        self._reset_task_monitor()
                        self._reset_first_event_progress_watch()
                        self._restore_pre_run_snapshot_after_runtime()
                        self.front_loop_after_id = self.root.after(
                            300,
                            lambda: self._dispatch_front_loop_once(display_name)
                        )
                        return
                    self._arm_first_event_progress_watch(sent_at_monotonic, delay_meta["send_delay_sec"], first_event_timing)
                    self._monitor_after_ack(res)

                    self.front_loop_round += 1
                    suffix = ""
                    if resolve_note:
                        suffix = "（{}）".format(resolve_note)
                    self.set_status(
                        "前端重複送出中：{} 第 {} 輪{}".format(display_name, self.front_loop_round, suffix)
                    )
                    self._poll_front_loop_round_done(display_name)
                except Exception as e:
                    if isinstance(e, TimeoutError):
                        self._mark_task_monitor_failed(
                            phase="ACK_TIMEOUT",
                            elapsed_ms=ACK_TIMEOUT_MS,
                            threshold_ms=ACK_TIMEOUT_MS,
                            last_state="submitted"
                        )
                    else:
                        self._reset_task_monitor()
                        self.set_frontend_error(str(e))
                    self.front_loop_enabled = False
                    self._restore_pre_run_snapshot_after_runtime()
                    self.show_error("重複傳送失敗", str(e))

            delay_override = None if self.front_loop_round == 0 else 0.0
            self.apply_send_delay_if_needed(do_send, delay_override=delay_override)
        except Exception as e:
            self.front_loop_enabled = False
            self.set_frontend_error(str(e))
            self._restore_pre_run_snapshot_after_runtime()
            self.show_error("重複傳送失敗", str(e))

    def _poll_front_loop_round_done(self, display_name):
        if not self.front_loop_enabled:
            return
        try:
            status_response = self.request_pi({"action": "status"}, write_response=False, channel="status")
            runtime = status_response.get("timeline_runtime", {})
            state = str(runtime.get("state", "")).strip().lower()
            if state == "running":
                self.front_loop_after_id = self.root.after(
                    250,
                    lambda: self._poll_front_loop_round_done(display_name)
                )
                return
            self.front_loop_after_id = self.root.after(
                80,
                lambda: self._dispatch_front_loop_once(display_name)
            )
        except Exception:
            self.front_loop_after_id = self.root.after(
                400,
                lambda: self._poll_front_loop_round_done(display_name)
            )

    def prepare_events_for_send(self, action_reason="before_send"):
        try:
            offset_sec = self.get_manual_offset_sec()
            base_events = [self.normalize_event_schema(ev) for ev in self.timeline]
            rslot_count = sum(
                1 for ev in base_events
                if str(ev.get("type", "")).strip().lower() == "randat"
            )
            if rslot_count > 0:
                self.validate_buff_group_contiguity(base_events)
            self.validate_negative_group_monotonic_by_index(base_events)
            events = recalculate_runtime_events_by_index(base_events, offset_sec)
        except Exception as e:
            self.show_error("時間重算失敗", str(e))
            return None, ""
        unsupported = self.get_unsupported_buttons(events)
        if unsupported:
            msg = (
                "前端偵測到不支援的 button：{}\n"
                "已用紅底標記對應列。請先修改 button 欄位後再送出。".format(", ".join(unsupported))
            )
            self.set_frontend_error(msg)
            self.show_warning("送出前檢查", msg)
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
                    self.show_error("輸入錯誤", "選項不存在：{}".format(ans))
                    return None, ""
                chosen_cycle, chosen_jitter = option_map[ans]
                resolved_groups.append(str(group_name))

            for ev in events:
                if ev.get("buff_group", "").strip() == group_name:
                    ev["buff_cycle_sec"] = chosen_cycle
                    ev["buff_jitter_sec"] = chosen_jitter

        events, block_assignments, round_traces = allocate_randat_blocks(events)
        trace_by_group = {
            str(item.get("buffGroup", "")).strip(): dict(item)
            for item in round_traces
            if isinstance(item, dict) and str(item.get("buffGroup", "")).strip()
        }
        group_cycle_roll = {}
        for ev in events:
            ev["at"] = apply_positive_jitter(ev.get("at", 0.0), ev.get("at_jitter", 0.0))
            cycle = max(0.0, _safe_float(ev.get("buff_cycle_sec", 0.0)))
            jitter = max(0.0, _safe_float(ev.get("buff_jitter_sec", 0.0)))
            group = str(ev.get("buff_group", "")).strip()
            cycle_key = group if group else "__nogroup_{}".format(id(ev))
            if cycle > 0.0:
                if cycle_key not in group_cycle_roll:
                    group_cycle_roll[cycle_key] = round(cycle + random.uniform(0.0, jitter), 4)
                rolled_cycle = group_cycle_roll[cycle_key]
            else:
                rolled_cycle = round(cycle, 4)
            ev["runtime_buff_cycle_base"] = round(cycle, 4)
            ev["runtime_buff_cycle_applied"] = round(rolled_cycle, 4)
            ev["buff_cycle_sec"] = round(rolled_cycle, 4)
            ev["buff_jitter_sec"] = round(jitter, 4)
            ev["at_random_sec"] = 0.0
        for idx, ev in enumerate(events):
            ev["runtime_source_index"] = idx
        events = [ev for ev in events if str(ev.get("type", "")).strip().lower() in ("press", "release")]
        for ev in events:
            group = str(ev.get("buff_group", "")).strip()
            if group in trace_by_group:
                trace = trace_by_group[group]
                ev["runtime_round"] = int(trace.get("round", 0) or 0)
                ev["runtime_picked_slot"] = int(trace.get("picked_slot", trace.get("pickedIdx", ev.get("runtime_landed_index", 0))) or 0)
                ev["runtime_self_picked"] = 1 if bool(trace.get("self_picked", False)) else 0
                ev["runtime_group_color"] = str(trace.get("color", ev.get("runtime_group_color", ""))).strip().lower()

        resolve_note = ""
        if resolved_groups:
            resolve_note = "已解決衝突群組: {}".format("、".join(resolved_groups))
        random_note = "jitter 已前端重算（+0~j，僅延後）"
        resolve_note = "{}；{}".format(resolve_note, random_note) if resolve_note else random_note
        self.last_prepared_payload = {
            "action_reason": action_reason,
            "rslot_count": int(rslot_count),
            "prepared_events": self.copy_events(events),
            "block_assignments": copy.deepcopy(block_assignments),
            "round_traces": copy.deepcopy(round_traces),
            "resolve_note": resolve_note
        }
        self.render_prepared_payload()
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
                self.show_warning("提醒", "請先選取一列或多列，再雙擊欄位標題")
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
                self.show_error("修改失敗", str(e))
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
            self.show_error("修改失敗", str(e))
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
            self.set_status("剪貼簿沒有可貼上的內容")
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
            self.show_warning("提醒", "剪貼簿只有標題列，沒有可貼上的資料")
            return "break"

        try:
            parsed_rows = [self._normalize_paste_row(values) for values in parsed_rows]
        except Exception as e:
            self.show_error("貼上失敗", str(e))
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
                self.show_error("貼上失敗", str(e))
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
            self.show_warning("提醒", "貼上範圍超出目前列數，未更新任何資料")
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

        btn_above = tk.Button(btn_row, text="上方", width=10, command=lambda: choose("above"))
        btn_above.pack(side="left")
        tk.Button(btn_row, text="下方", width=10, command=lambda: choose("below")).pack(side="left", padx=(6, 0))
        tk.Button(btn_row, text="取消", width=10, command=lambda: choose(None)).pack(side="right")
        dialog.bind("<Up>", lambda _e: choose("above"))
        dialog.bind("<Down>", lambda _e: choose("below"))
        dialog.bind("<Escape>", lambda _e: choose(None))

        dialog.update_idletasks()
        dialog_w = dialog.winfo_width()
        dialog_h = dialog.winfo_height()
        x, y = self._compute_app_dialog_position(dialog_w, dialog_h)
        dialog.geometry("{}x{}+{}+{}".format(dialog_w, dialog_h, x, y))
        dialog.focus_force()
        btn_above.focus_set()

        self.root.wait_window(dialog)
        return choice["value"]

    def _normalize_paste_row(self, values):
        if not values:
            raise ValueError("貼上資料有空白列")
        if len(values) < 7:
            raise ValueError("每列至少需要 7 欄（type 到 buff_jitter_sec）")
        # 優先支援 7 欄格式；若是完整表格 9 欄（idx + 7 欄 + group），則忽略 idx/group。
        if len(values) >= 9 and values[1].strip().lower() in ("press", "release", "randat"):
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
        row_type = str(event.get("type", "")).strip().lower()
        if field == "type" and value not in ("press", "release", "randat"):
            raise ValueError("type 只能是 press / release / randat")
        if field == "button" and row_type != "randat" and not value:
            raise ValueError("button 不可空白")
        if field == "buff_group":
            value = raw_value.strip()
        event[field] = value

    def insert_randat_row(self):
        if not self._ensure_runtime_editable():
            return
        selected = sorted(self.get_selected_indexes())
        insert_at = 0
        anchor_idx = None
        if selected:
            anchor_idx = selected[0]
            position = self.ask_paste_position(anchor_idx, 1)
            if position is None:
                self.set_status("已取消插入 randat")
                return
            insert_at = anchor_idx if position == "above" else anchor_idx + 1

        at_value = 0.0
        if anchor_idx is not None and 0 <= anchor_idx < len(self.timeline):
            try:
                at_value = round(max(0.0, float(self.timeline[anchor_idx].get("at", 0.0))), 2)
            except Exception:
                at_value = 0.0

        event = {
            "type": "randat",
            "button": "",
            "at": at_value,
            "at_jitter": 0.0,
            "buff_group": "",
            "buff_cycle_sec": 0.0,
            "buff_jitter_sec": 0.0,
            "row_color": "#d8ecff",
            "replicatedRow": 0
        }
        before = self._begin_timeline_change()
        self.timeline.insert(insert_at, event)
        self._finalize_timeline_change(before)
        self.mark_timeline_dirty()
        self.tree.selection_set([str(insert_at)])
        self.set_status("已插入 randat 列於 idx {}".format(insert_at))

    def move_selected_up(self):
        if not self._ensure_runtime_editable():
            return
        selected = sorted(self.get_selected_indexes())
        if not selected:
            self.show_warning("提醒", "請先選取列")
            return
        if selected[0] == 0:
            return
        before = self._begin_timeline_change()
        self.timeline, moved_indexes, move_meta = move_rows_with_ab_gap_compensation(
            self.timeline,
            selected,
            "up"
        )
        self._finalize_timeline_change(before)
        self.mark_timeline_dirty()
        self.tree.selection_set([str(i) for i in moved_indexes])
        if move_meta:
            if not isinstance(self.runtime_move_gap_logs, list):
                self.runtime_move_gap_logs = []
            self.runtime_move_gap_logs.append(move_meta)
            self.runtime_move_gap_logs = self.runtime_move_gap_logs[-100:]
        self.render_runtime_analysis(force=True)

    def move_selected_down(self):
        if not self._ensure_runtime_editable():
            return
        selected = sorted(self.get_selected_indexes())
        if not selected:
            self.show_warning("提醒", "請先選取列")
            return
        if selected[-1] == len(self.timeline) - 1:
            return
        before = self._begin_timeline_change()
        self.timeline, moved_indexes, move_meta = move_rows_with_ab_gap_compensation(
            self.timeline,
            selected,
            "down"
        )
        self._finalize_timeline_change(before)
        self.mark_timeline_dirty()
        self.tree.selection_set([str(i) for i in moved_indexes])
        if move_meta:
            if not isinstance(self.runtime_move_gap_logs, list):
                self.runtime_move_gap_logs = []
            self.runtime_move_gap_logs.append(move_meta)
            self.runtime_move_gap_logs = self.runtime_move_gap_logs[-100:]
        self.render_runtime_analysis(force=True)

    def swap_selected_at(self):
        if not self._ensure_runtime_editable():
            return
        selected = sorted(self.get_selected_indexes())
        if len(selected) != 2:
            self.show_warning("提醒", "at 交換只支援恰好選取 2 列")
            return
        idx_a, idx_b = selected
        before = self._begin_timeline_change()
        at_a = float(self.timeline[idx_a].get("at", 0.0))
        at_b = float(self.timeline[idx_b].get("at", 0.0))
        self.timeline[idx_a]["at"] = round(at_b, 2)
        self.timeline[idx_b]["at"] = round(at_a, 2)
        self._finalize_timeline_change(before)
        self.mark_timeline_dirty()
        self.tree.selection_set([str(idx_a), str(idx_b)])
        self.set_status("已交換 idx {} 與 idx {} 的 at".format(idx_a, idx_b))

    def _collect_press_release_for_pr(self):
        rows = []
        for idx, ev in enumerate(self.timeline):
            ev_type = str(ev.get("type", "")).strip().lower()
            if ev_type not in ("press", "release"):
                continue
            rows.append({
                "idx": idx,
                "at": round(max(0.0, _safe_float(ev.get("at", 0.0))), 4),
                "type": ev_type,
                "button_id": str(ev.get("button", "")).strip(),
                "at_jitter": round(max(0.0, _safe_float(ev.get("at_jitter", 0.0))), 4),
            })
        return rows

    def _show_jitter_order_risk_reminder(self, source_label):
        events = self._collect_press_release_for_pr()
        if len(events) < 2:
            return
        risks = detect_jitter_order_risk_pairs(events)
        if not risks:
            return
        overtakes = [r for r in risks if r.get("risk_level") == "overtake_risk"]
        ties = [r for r in risks if r.get("risk_level") == "tie_risk"]
        preview_rows = risks[:20]
        lines = [
            "來源：{}".format(source_label),
            "偵測到順序風險 {} 組（超車風險 {}、同時刻風險 {}）".format(
                len(risks), len(overtakes), len(ties)
            ),
            "（僅列前 20 組）"
        ]
        for row in preview_rows:
            lines.append(
                "idx{idx_a}->{idx_b} gap={gap:.4f} jitter_a={at_jitter_a:.4f} {risk_level}".format(**row)
            )
        self.show_warning("jitter 順序提醒", "\n".join(lines))

    def analyze_pr_pairs(self):
        events = self._collect_press_release_for_pr()
        if len(events) < 2:
            self.show_warning("提醒", "至少需要 2 筆 press/release 才能做 PR 分析")
            return
        analysis = analyze_pr_gap_events(events, top_n=None)
        payload = {
            "task": "pr_gap_analysis",
            "total_events": len(events),
            "min_all_pairs": analysis["min_all_pairs"],
            "pr_pairs": analysis["pr_pairs"][:100],
            "segments": analysis["segments"]
        }
        self.write_text(payload)
        self._open_pr_tables_dialog(
            "PR 分析結果",
            sections=[
                {
                    "title": "PR Pairs",
                    "columns": ["pr_rank", "idx_a", "idx_b", "button_a", "button_b", "type_a", "type_b", "at_a", "at_b", "gap"],
                    "rows": analysis["pr_pairs"]
                },
                {
                    "title": "Segments",
                    "columns": ["range", "avg_gap", "pair_count"],
                    "rows": [
                        {
                            "range": item.get("range", ""),
                            "avg_gap": item.get("avg_gap", 0.0),
                            "pair_count": len(item.get("pairs", []))
                        }
                        for item in analysis.get("segments", [])
                    ]
                }
            ]
        )
        self.set_status("PR 分析完成：共 {} 組相鄰 pair（表格已開啟）".format(len(analysis["pr_pairs"])))

    def apply_minimum_gap_for_pairs(self):
        if not self._ensure_runtime_editable():
            return
        events = self._collect_press_release_for_pr()
        if len(events) < 2:
            self.show_warning("提醒", "至少需要 2 筆 press/release 才能套用 minimum gap")
            return
        try:
            min_gap = max(0.0, float(self.minimum_gap_entry.get().strip()))
        except Exception:
            self.show_error("錯誤", "minimum gap 必須是數字")
            return

        before_analysis = analyze_pr_gap_events(events, top_n=None)
        target_pairs = before_analysis.get("pr_pairs", [])
        adjusted, adjust_logs = apply_minimum_gap_by_pairs(
            events=events,
            pairs_snapshot=target_pairs,
            minimum_gap=min_gap
        )
        after_analysis = analyze_pr_gap_events(adjusted, top_n=None)

        idx_to_at = {int(ev["idx"]): round(_safe_float(ev.get("at", 0.0)), 4) for ev in adjusted}
        before = self._begin_timeline_change()
        for idx, ev in enumerate(self.timeline):
            if idx in idx_to_at:
                ev["at"] = idx_to_at[idx]
        self._finalize_timeline_change(before)
        self.mark_timeline_dirty()
        self.refresh_tree()

        self.write_text({
            "task": "minimum_gap_adjustment",
            "minimum_gap": round(min_gap, 4),
            "analysis_before": {
                "min_all_pairs": before_analysis["min_all_pairs"],
                "pr_pairs": before_analysis["pr_pairs"][:100],
                "segments": before_analysis["segments"],
            },
            "manual_adjustment_log": adjust_logs,
            "analysis_after": {
                "min_all_pairs": after_analysis["min_all_pairs"],
                "pr_pairs": after_analysis["pr_pairs"][:100],
                "segments": after_analysis["segments"],
            }
        })
        self._open_pr_tables_dialog(
            "minimum gap 套用結果",
            sections=[
                {
                    "title": "Before PR",
                    "columns": ["pr_rank", "idx_a", "idx_b", "button_a", "button_b", "type_a", "type_b", "at_a", "at_b", "gap"],
                    "rows": before_analysis["pr_pairs"]
                },
                {
                    "title": "After PR",
                    "columns": ["pr_rank", "idx_a", "idx_b", "button_a", "button_b", "type_a", "type_b", "at_a", "at_b", "gap"],
                    "rows": after_analysis["pr_pairs"]
                },
                {
                    "title": "Adjustment Log",
                    "columns": ["pr_rank", "idx_a", "idx_b", "current_gap", "minimum_gap", "delta", "affected_count"],
                    "rows": [
                        {
                            "pr_rank": row.get("pr_rank", ""),
                            "idx_a": row.get("idx_a", ""),
                            "idx_b": row.get("idx_b", ""),
                            "current_gap": row.get("current_gap", ""),
                            "minimum_gap": row.get("minimum_gap", ""),
                            "delta": row.get("delta", ""),
                            "affected_count": row.get("affected_idx_range", {}).get("count", 0),
                        }
                        for row in adjust_logs
                    ]
                }
            ]
        )
        self.set_status(
            "minimum gap 套用完成：共處理 {} 組相鄰 pair（表格已開啟）".format(len(adjust_logs))
        )

    def delete_selected_rows(self):
        if not self._ensure_runtime_editable():
            return
        selected = sorted(self.get_selected_indexes(), reverse=True)
        if not selected:
            self.show_warning("提醒", "請先選取列")
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
