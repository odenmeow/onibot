# -*- coding: utf-8 -*-
"""
executor_only 邊界說明：
- 此模組僅負責「執行」已排程完成的 timeline 事件（press/release + at_ms）。
- start_task 只接受可執行欄位：idx/at_ms/action/btn/skip_mode。
- 任何抽籤/排程語意欄位（例如 randat 指令級資料）一律拒收，避免後端參與排程決策。
"""
import json
import random
import socket
import threading
import time
import uuid
import RPi.GPIO as GPIO

HOST = "0.0.0.0"
PORT = 5000

ACTIVE_LOW = True
DEFAULT_PRESS_TIME = 0.25
BUFF_SKIP_MODE_WALK = "walk"          # 走過：不按，但保留原時間軸（照等）
BUFF_SKIP_MODE_PASS = "pass"          # 略過：不按，並壓縮時間軸（不等）
BUFF_SKIP_MODE_NONE = "none"          # 相容：等同正常按放
BUFF_SKIP_MODE_COMPRESS = "compress"  # deprecated: 等同 pass
RUNTIME_CONTRACT_VERSION = "v1"

BUTTONS = {
    "fn": 26,
    "g": 2,
    "shift": 3,
    "f": 6,
    "c": 17,
    "v": 27,
    "d": 22,
    "alt": 10,
    "ctrl": 9,
    "left": 11,
    "up": 21,
    "down": 20,
    "right": 16,
    "x": 12,
    "space": 19,
    "6": 13
}

stop_event = threading.Event()
pause_event = threading.Event()
run_lock = threading.Lock()
runtime_lock = threading.Lock()

current_run_thread = None
current_server_task_id = ""
current_run_status = {
    "state": "idle",   # idle / running / stopping / stopped / error
    "mode": "",
    "message": "",
    "server_task_id": ""
}
timeline_runtime = {
    "run_id": 0,
    "server_task_id": "",
    "state": "idle",   # idle / running / paused / resumed / stopped / finished / error
    "mode": "",
    "loop_count": 0,
    "events_total": 0,
    "processed_count": 0,
    "events": [],
    "last_event": None,
    "round_traces": [],
    "runtime_diag": {},
    "progress": {
        "current_idx": -1,
        "event_time_ms": 0
    }
}
timeline_cooldown_runtime = {
    "next_ready_at": {},
    "landed_index_by_group": {}
}

# 狀態輪詢防爆：避免 timeline_runtime.events / round_traces 無上限回傳造成卡頓
STATUS_EVENTS_LIMIT = 120
STATUS_ROUND_TRACES_LIMIT = 40
START_TASK_ALLOWED_TIMELINE_FIELDS = {"idx", "at_ms", "action", "btn", "skip_mode"}


def _now_ms():
    return int(time.time() * 1000)


def _new_server_task_id():
    return "srv-{}-{}".format(_now_ms(), uuid.uuid4().hex[:8])


def make_error_response(code, message, phase, diag=None, status="error", server_task_id=None):
    return {
        "type": "error",
        "status": str(status or "error"),
        "code": str(code or "unknown_error"),
        "phase": str(phase or "request"),
        "message": str(message or "error"),
        "server_task_id": str(server_task_id or current_server_task_id or "").strip(),
        "diag": diag if isinstance(diag, dict) else {},
        "event_time_ms": _now_ms()
    }


def set_timeline_runtime(mode, state, events_total=0, loop_count=0, server_task_id=""):
    with runtime_lock:
        timeline_runtime["run_id"] += 1
        timeline_runtime["server_task_id"] = str(server_task_id or "")
        timeline_runtime["state"] = state
        timeline_runtime["mode"] = mode
        timeline_runtime["loop_count"] = int(loop_count)
        timeline_runtime["events_total"] = int(events_total)
        timeline_runtime["processed_count"] = 0
        timeline_runtime["events"] = []
        timeline_runtime["last_event"] = None
        timeline_runtime["round_traces"] = []
        timeline_runtime["runtime_diag"] = {}
        timeline_runtime["progress"] = {
            "current_idx": -1,
            "event_time_ms": _now_ms()
        }
        timeline_cooldown_runtime["next_ready_at"] = {}
        timeline_cooldown_runtime["landed_index_by_group"] = {}


def append_timeline_runtime_event(event_payload):
    try:
        current_idx = int(event_payload.get("original_index"))
    except Exception:
        current_idx = -1
    event_time_ms = _now_ms()
    with runtime_lock:
        timeline_runtime["events"].append(event_payload)
        timeline_runtime["processed_count"] = len(timeline_runtime["events"])
        timeline_runtime["last_event"] = event_payload
        timeline_runtime["progress"] = {
            "current_idx": current_idx,
            "event_time_ms": event_time_ms
        }


def patch_timeline_runtime(**kwargs):
    with runtime_lock:
        for key, value in kwargs.items():
            timeline_runtime[key] = value


def append_timeline_round_trace(trace_payload):
    with runtime_lock:
        timeline_runtime["round_traces"].append(trace_payload)


def get_timeline_runtime_snapshot():
    with runtime_lock:
        now_abs = time.monotonic()
        cooldowns = []
        for group, ready_at in timeline_cooldown_runtime.get("next_ready_at", {}).items():
            remain = max(0.0, float(ready_at) - now_abs)
            if remain <= 0:
                continue
            cooldowns.append({
                "buff_group": str(group),
                "remain_sec": round(remain, 3),
                "landed_index": timeline_cooldown_runtime.get("landed_index_by_group", {}).get(group)
            })
        events = timeline_runtime.get("events", [])
        if len(events) > STATUS_EVENTS_LIMIT:
            events = events[-STATUS_EVENTS_LIMIT:]
        else:
            events = list(events)

        round_traces = timeline_runtime.get("round_traces", [])
        if len(round_traces) > STATUS_ROUND_TRACES_LIMIT:
            round_traces = round_traces[-STATUS_ROUND_TRACES_LIMIT:]
        else:
            round_traces = list(round_traces)

        return {
            "run_id": int(timeline_runtime.get("run_id", 0)),
            "server_task_id": str(timeline_runtime.get("server_task_id", "")),
            "state": timeline_runtime.get("state", "idle"),
            "mode": timeline_runtime.get("mode", ""),
            "loop_count": int(timeline_runtime.get("loop_count", 0)),
            "events_total": int(timeline_runtime.get("events_total", 0)),
            "processed_count": int(timeline_runtime.get("processed_count", 0)),
            "progress": dict(timeline_runtime.get("progress", {})),
            "events": events,
            "last_event": timeline_runtime.get("last_event"),
            "round_traces": round_traces,
            "runtime_diag": dict(timeline_runtime.get("runtime_diag", {})),
            "cooldowns": cooldowns
        }


def _normalize_round_traces_payload(raw):
    if not isinstance(raw, list):
        return []
    return [item for item in raw if isinstance(item, dict)]


def _extract_runtime_diag_from_start_task(data):
    runtime_meta = data.get("runtime_meta")
    round_traces = data.get("round_traces")
    diag = {}
    normalized_round_traces = []
    if isinstance(runtime_meta, dict):
        for key in ("rslot_count", "randat_executed", "picked_reason", "draw_result"):
            if key in runtime_meta:
                diag[key] = runtime_meta.get(key)
        normalized_round_traces = _normalize_round_traces_payload(runtime_meta.get("round_traces"))
    if not normalized_round_traces:
        normalized_round_traces = _normalize_round_traces_payload(round_traces)
    return normalized_round_traces, diag


def get_press_level():
    return GPIO.LOW if ACTIVE_LOW else GPIO.HIGH


def get_release_level():
    return GPIO.HIGH if ACTIVE_LOW else GPIO.LOW


PRESS_LEVEL = get_press_level()
RELEASE_LEVEL = get_release_level()


def setup_gpio():
    GPIO.setmode(GPIO.BCM)
    GPIO.setwarnings(False)
    for pin in BUTTONS.values():
        GPIO.setup(pin, GPIO.OUT)
        GPIO.output(pin, RELEASE_LEVEL)


def release_all():
    for pin in BUTTONS.values():
        GPIO.output(pin, RELEASE_LEVEL)


def cleanup_gpio():
    release_all()
    GPIO.cleanup()


def press_only(name):
    if name not in BUTTONS:
        raise ValueError("找不到按鍵: {}".format(name))
    pin = BUTTONS[name]
    GPIO.output(pin, PRESS_LEVEL)
    print("[press] {} GPIO{}".format(name, pin))


def release_only(name):
    if name not in BUTTONS:
        raise ValueError("找不到按鍵: {}".format(name))
    pin = BUTTONS[name]
    GPIO.output(pin, RELEASE_LEVEL)
    print("[release] {} GPIO{}".format(name, pin))


def _wait_if_paused():
    paused_once = False
    while pause_event.is_set():
        if stop_event.is_set():
            raise InterruptedError("執行已被停止")
        paused_once = True
        patch_timeline_runtime(state="paused")
        time.sleep(0.05)
    if stop_event.is_set():
        raise InterruptedError("執行已被停止")
    if paused_once:
        patch_timeline_runtime(state="resumed")
    return paused_once


def safe_sleep(seconds, check_interval=0.01):
    remaining = max(0.0, float(seconds))
    while remaining > 0.0:
        _wait_if_paused()
        if stop_event.is_set():
            raise InterruptedError("執行已被停止")
        chunk = min(check_interval, remaining)
        tick = time.monotonic()
        time.sleep(chunk)
        elapsed = max(0.0, time.monotonic() - tick)
        if not pause_event.is_set():
            remaining -= elapsed


def press_button(name, duration=DEFAULT_PRESS_TIME):
    if stop_event.is_set():
        raise InterruptedError("執行已被停止")
    press_only(name)
    try:
        safe_sleep(duration)
    finally:
        release_only(name)
    safe_sleep(0.03)


def apply_jitter(base_value, jitter):
    if jitter is None:
        jitter = 0.0
    jitter = abs(float(jitter))
    value = float(base_value) + random.uniform(0.0, jitter)
    return max(0.0, value)


def run_macro(steps):
    results = []
    stop_event.clear()

    with run_lock:
        for i, step in enumerate(steps):
            if stop_event.is_set():
                raise InterruptedError("執行已被停止")

            button = step["button"]
            delay = apply_jitter(step.get("delay", 0.0), step.get("delay_jitter", 0.0))
            duration = apply_jitter(step.get("duration", DEFAULT_PRESS_TIME), step.get("duration_jitter", 0.0))

            print("[macro step {}] button={} delay={:.4f} duration={:.4f}".format(
                i, button, delay, duration
            ))

            safe_sleep(delay)
            press_button(button, duration)

            results.append({
                "index": i,
                "button": button,
                "actual_delay": round(delay, 4),
                "actual_duration": round(duration, 4),
                "status": "ok"
            })

    return results


def run_timeline(events, reset_stop_event=True, buff_runtime=None, buff_skip_mode=BUFF_SKIP_MODE_WALK):
    if not isinstance(events, list):
        raise ValueError("events 必須是 list")
    if buff_skip_mode not in (BUFF_SKIP_MODE_WALK, BUFF_SKIP_MODE_PASS, BUFF_SKIP_MODE_NONE):
        raise ValueError("buff_skip_mode 錯誤: {}".format(buff_skip_mode))

    if reset_stop_event:
        stop_event.clear()
    normalized = []

    group_config = {}
    for i, ev in enumerate(events):
        ev_type = ev.get("type")
        button = ev.get("button")
        at = max(0.0, float(ev.get("at", 0.0)))
        buff_group = str(ev.get("buff_group", "")).strip()
        buff_cycle_sec = max(0.0, float(ev.get("buff_cycle_sec", 0.0)))
        buff_jitter_sec = abs(float(ev.get("buff_jitter_sec", 0.0)))

        if ev_type not in ("press", "release"):
            raise ValueError("第 {} 筆 event type 錯誤".format(i))
        if button not in BUTTONS:
            raise ValueError("第 {} 筆找不到按鍵: {}".format(i, button))

        if buff_group and buff_cycle_sec > 0:
            cfg = group_config.get(buff_group)
            if cfg is None:
                group_config[buff_group] = {
                    "cycle_sec": buff_cycle_sec,
                    "jitter_sec": buff_jitter_sec
                }
        try:
            source_index = int(ev.get("runtime_source_index", i))
        except Exception:
            source_index = i
        normalized.append({
            "original_index": source_index,
            "type": ev_type,
            "button": button,
            "at": at,
            "buff_group": buff_group,
            "runtime_landed_index": ev.get("runtime_landed_index"),
            "runtime_anchor_index": ev.get("runtime_anchor_index"),
            "runtime_occupies_original": ev.get("runtime_occupies_original", 0),
            "runtime_round_trace": ev.get("runtime_round_trace")
        })

    normalized.sort(key=lambda x: x["at"])
    results = []

    with run_lock:
        start = time.monotonic()
        timeline_shift = 0.0
        emitted_round_trace_groups = set()

        for i, ev in enumerate(normalized):
            resumed = _wait_if_paused()
            if resumed:
                patch_timeline_runtime(state="running")
            if stop_event.is_set():
                raise InterruptedError("執行已被停止")

            source_target = ev["at"]
            target = max(0.0, source_target - timeline_shift)
            now = time.monotonic() - start
            wait_time = target - now
            skip_by_cooldown = False

            buff_group = ev.get("buff_group", "")
            round_trace = ev.get("runtime_round_trace")
            if (
                buff_group
                and buff_group not in emitted_round_trace_groups
                and isinstance(round_trace, dict)
            ):
                append_timeline_round_trace(round_trace)
                emitted_round_trace_groups.add(buff_group)
            if buff_runtime is not None and buff_group in group_config:
                now_abs = time.monotonic()
                next_ready = buff_runtime["next_ready_at"].get(buff_group)
                if next_ready is not None and now_abs < next_ready:
                    skip_by_cooldown = True
                else:
                    cfg = group_config[buff_group]
                    cd = max(0.0, cfg["cycle_sec"])
                    buff_runtime["next_ready_at"][buff_group] = now_abs + cd
                    landed_index = ev.get("runtime_landed_index")
                    with runtime_lock:
                        timeline_cooldown_runtime["next_ready_at"][buff_group] = now_abs + cd
                        timeline_cooldown_runtime["landed_index_by_group"][buff_group] = landed_index

            if skip_by_cooldown and buff_skip_mode != BUFF_SKIP_MODE_NONE:
                compressed_sec = 0.0
                if buff_skip_mode == BUFF_SKIP_MODE_PASS:
                    compressed_sec = max(0.0, wait_time)
                    timeline_shift += compressed_sec
                elif buff_skip_mode == BUFF_SKIP_MODE_WALK and wait_time > 0:
                    safe_sleep(wait_time)
                results.append({
                    "index": i,
                    "original_index": ev["original_index"],
                    "type": ev["type"],
                    "button": ev["button"],
                    "source_target_at": round(source_target, 4),
                    "target_at": round(target, 4),
                    "actual_at": round(time.monotonic() - start, 4),
                    "status": "skipped_by_cooldown",
                    "buff_skip_mode": buff_skip_mode,
                    "compressed_sec": round(compressed_sec, 4)
                })
                append_timeline_runtime_event({
                    "index": i,
                    "original_index": ev["original_index"],
                    "type": ev["type"],
                    "button": ev["button"],
                    "status": "skipped_by_cooldown",
                    "source_target_at": round(source_target, 4),
                    "target_at": round(target, 4),
                    "actual_at": round(time.monotonic() - start, 4),
                    "buff_skip_mode": buff_skip_mode,
                    "buff_group": buff_group,
                    "runtime_landed_index": ev.get("runtime_landed_index"),
                    "runtime_anchor_index": ev.get("runtime_anchor_index"),
                    "runtime_occupies_original": ev.get("runtime_occupies_original", 0)
                })
                continue

            if wait_time > 0:
                safe_sleep(wait_time)

            if ev["type"] == "press":
                press_only(ev["button"])
            else:
                release_only(ev["button"])

            actual_now = time.monotonic() - start
            results.append({
                "index": i,
                "original_index": ev["original_index"],
                "type": ev["type"],
                "button": ev["button"],
                "source_target_at": round(source_target, 4),
                "target_at": round(target, 4),
                "actual_at": round(actual_now, 4),
                "status": "ok"
            })
            append_timeline_runtime_event({
                "index": i,
                "original_index": ev["original_index"],
                "type": ev["type"],
                "button": ev["button"],
                "status": "ok",
                "source_target_at": round(source_target, 4),
                "target_at": round(target, 4),
                "actual_at": round(actual_now, 4),
                "buff_group": buff_group,
                "runtime_landed_index": ev.get("runtime_landed_index"),
                "runtime_anchor_index": ev.get("runtime_anchor_index"),
                "runtime_occupies_original": ev.get("runtime_occupies_original", 0)
            })

    return results


def run_timeline_background(events, buff_skip_mode=BUFF_SKIP_MODE_WALK):
    global current_run_status
    try:
        pause_event.clear()
        set_timeline_runtime(
            "timeline",
            "running",
            events_total=len(events),
            loop_count=1,
            server_task_id=current_server_task_id
        )
        current_run_status = {
            "state": "running",
            "mode": "timeline",
            "message": "正在執行 timeline",
            "server_task_id": current_server_task_id
        }
        run_timeline(events, buff_skip_mode=buff_skip_mode)
        release_all()
        patch_timeline_runtime(state="finished")
        current_run_status = {
            "state": "idle",
            "mode": "timeline",
            "message": "timeline 執行完成",
            "server_task_id": current_server_task_id
        }
    except InterruptedError as e:
        release_all()
        patch_timeline_runtime(state="stopped")
        current_run_status = {
            "state": "stopped",
            "mode": "timeline",
            "message": str(e),
            "server_task_id": current_server_task_id
        }
    except Exception as e:
        release_all()
        patch_timeline_runtime(state="error")
        current_run_status = {
            "state": "error",
            "mode": "timeline",
            "message": str(e),
            "server_task_id": current_server_task_id
        }


def run_timeline_loop_background(events, buff_skip_mode=BUFF_SKIP_MODE_WALK):
    global current_run_status
    loop_count = 0
    buff_runtime = {"next_ready_at": {}}
    try:
        stop_event.clear()
        pause_event.clear()
        while not stop_event.is_set():
            loop_count += 1
            set_timeline_runtime(
                "timeline_loop",
                "running",
                events_total=len(events),
                loop_count=loop_count,
                server_task_id=current_server_task_id
            )
            current_run_status = {
                "state": "running",
                "mode": "timeline_loop",
                "message": "正在第 {} 次循環".format(loop_count),
                "server_task_id": current_server_task_id
            }
            run_timeline(
                events,
                reset_stop_event=False,
                buff_runtime=buff_runtime,
                buff_skip_mode=buff_skip_mode
            )
            patch_timeline_runtime(state="finished", loop_count=loop_count)
            release_all()

        patch_timeline_runtime(state="stopped", loop_count=loop_count)
        current_run_status = {
            "state": "stopped",
            "mode": "timeline_loop",
            "message": "已停止循環，共執行 {} 次".format(loop_count),
            "server_task_id": current_server_task_id
        }
    except InterruptedError as e:
        release_all()
        patch_timeline_runtime(state="stopped", loop_count=loop_count)
        current_run_status = {
            "state": "stopped",
            "mode": "timeline_loop",
            "message": "{}（共執行 {} 次）".format(str(e), loop_count),
            "server_task_id": current_server_task_id
        }
    except Exception as e:
        release_all()
        patch_timeline_runtime(state="error", loop_count=loop_count)
        current_run_status = {
            "state": "error",
            "mode": "timeline_loop",
            "message": str(e),
            "server_task_id": current_server_task_id
        }


def parse_start_task_timeline(raw_timeline):
    if not isinstance(raw_timeline, list):
        raise ValueError("timeline 必須是 list")
    events = []
    for i, row in enumerate(raw_timeline):
        if not isinstance(row, dict):
            raise ValueError("timeline 第 {} 列格式錯誤".format(i))
        # executor_only：只接受可執行欄位，拒收任何排程語意欄位（如 randat）。
        extra_keys = sorted([k for k in row.keys() if k not in START_TASK_ALLOWED_TIMELINE_FIELDS])
        if extra_keys:
            raise ValueError(
                "timeline 第 {} 列包含未允許欄位: {}（僅接受: idx/at_ms/action/btn/skip_mode）".format(
                    i, ",".join(extra_keys)
                )
            )
        action = str(row.get("action", "")).strip().lower()
        button = str(row.get("btn", "")).strip().lower()
        if action not in ("press", "release"):
            raise ValueError("timeline 第 {} 列 action 錯誤".format(i))
        if button not in BUTTONS:
            raise ValueError("timeline 第 {} 列 btn 不存在: {}".format(i, button))

        try:
            at_ms = int(row.get("at_ms", 0))
        except Exception:
            raise ValueError("timeline 第 {} 列 at_ms 錯誤".format(i))
        events.append({
            "type": action,
            "button": button,
            "at": max(0, at_ms) / 1000.0
        })
    return events


def run_macro_background(steps):
    global current_run_status
    try:
        current_run_status = {
            "state": "running",
            "mode": "macro",
            "message": "正在執行 macro",
            "server_task_id": current_server_task_id
        }
        run_macro(steps)
        release_all()
        current_run_status = {
            "state": "idle",
            "mode": "macro",
            "message": "macro 執行完成",
            "server_task_id": current_server_task_id
        }
    except InterruptedError as e:
        release_all()
        current_run_status = {
            "state": "stopped",
            "mode": "macro",
            "message": str(e),
            "server_task_id": current_server_task_id
        }
    except Exception as e:
        release_all()
        current_run_status = {
            "state": "error",
            "mode": "macro",
            "message": str(e),
            "server_task_id": current_server_task_id
        }


def _release_all_background():
    global current_run_status
    try:
        release_all()
    finally:
        state = str(current_run_status.get("state", "")).strip().lower()
        if state == "stopping":
            current_run_status = {
                "state": "stopped",
                "mode": current_run_status.get("mode", ""),
                "message": "已停止，GPIO 全部釋放",
                "server_task_id": current_server_task_id
            }
            patch_timeline_runtime(state="stopped")


def stop_current_run():
    global current_run_status
    stop_event.set()
    pause_event.clear()
    mode = current_run_status.get("mode", "")
    current_run_status = {
        "state": "stopping",
        "mode": mode,
        "message": "已收到停止指令，正在停止並釋放 GPIO",
        "server_task_id": current_server_task_id
    }
    patch_timeline_runtime(state="stopped")
    threading.Thread(target=_release_all_background, daemon=True).start()
    return {
        "status": "ok",
        "message": "已收到停止指令"
    }


def pause_current_run():
    global current_run_status
    if current_run_thread is None or not current_run_thread.is_alive():
        return {"status": "error", "message": "目前沒有執行中的工作"}
    pause_event.set()
    mode = current_run_status.get("mode", "")
    current_run_status = {
        "state": "paused",
        "mode": mode,
        "message": "已暫停",
        "server_task_id": current_server_task_id
    }
    patch_timeline_runtime(state="paused")
    return {"status": "ok", "state": "paused", "message": "已暫停"}


def resume_current_run():
    global current_run_status
    if current_run_thread is None or not current_run_thread.is_alive():
        return {"status": "error", "message": "目前沒有可繼續的工作"}
    pause_event.clear()
    mode = current_run_status.get("mode", "")
    current_run_status = {
        "state": "resumed",
        "mode": mode,
        "message": "已繼續",
        "server_task_id": current_server_task_id
    }
    patch_timeline_runtime(state="resumed")
    return {"status": "ok", "state": "resumed", "message": "已繼續"}


def handle_request(data):
    global current_run_thread, current_run_status, current_server_task_id

    action = data.get("action")

    def normalize_buff_skip_mode(raw):
        mode = str(raw or BUFF_SKIP_MODE_WALK).strip().lower()
        deprecated_alias = ""
        if mode == BUFF_SKIP_MODE_COMPRESS:
            deprecated_alias = BUFF_SKIP_MODE_COMPRESS
            mode = BUFF_SKIP_MODE_PASS
        if mode not in (BUFF_SKIP_MODE_WALK, BUFF_SKIP_MODE_PASS, BUFF_SKIP_MODE_NONE):
            raise ValueError("buff_skip_mode 只能是 pass、walk（或相容 none/compress）")
        return mode, deprecated_alias

    service_meta = {
        "service_name": "onibot-backend",
        "contract_version": RUNTIME_CONTRACT_VERSION,
        "supports_start_task": True,
        "supported_actions": ["ping", "list_buttons", "status", "hello", "stop", "pause", "resume", "run_macro", "start_task"]
    }

    if action == "ping":
        return {"status": "ok", "message": "pong"}

    if action == "list_buttons":
        return {"status": "ok", "buttons": BUTTONS}

    if action == "status":
        payload = {
            "status": "ok",
            "run_status": current_run_status,
            "timeline_runtime": get_timeline_runtime_snapshot()
        }
        payload.update(service_meta)
        return payload

    if action == "hello":
        return {"status": "ok", **service_meta}

    if action == "stop":
        return stop_current_run()
    if action == "pause":
        return pause_current_run()
    if action == "resume":
        return resume_current_run()

    if action == "run_macro":
        if current_run_thread is not None and current_run_thread.is_alive():
            return make_error_response(
                code="busy",
                message="Pi 目前已有執行中的工作",
                phase="submit",
                status="busy",
                server_task_id=current_server_task_id,
                diag={"last_state": str(current_run_status.get("state", "")), "elapsed_ms": 0, "threshold_ms": 0}
            )

        steps = data.get("steps", [])
        stop_event.clear()
        pause_event.clear()
        current_run_thread = threading.Thread(
            target=run_macro_background,
            args=(steps,),
            daemon=True
        )
        current_run_thread.start()
        return {"status": "ok", "mode": "macro", "message": "已收到 macro，開始背景執行"}

    if str(data.get("type", "")).strip().lower() == "start_task":
        # executor_only：只執行前端/上游已決定好的 timeline，不處理排程語意。
        if current_run_thread is not None and current_run_thread.is_alive():
            return make_error_response(
                code="busy",
                message="Pi 目前已有執行中的工作",
                phase="submit",
                status="busy",
                server_task_id=current_server_task_id,
                diag={"last_state": str(current_run_status.get("state", "")), "elapsed_ms": 0, "threshold_ms": 0}
            )

        raw_contract_version = data.get("contract_version")
        contract_version = str(raw_contract_version).strip() if raw_contract_version is not None else ""
        if contract_version != RUNTIME_CONTRACT_VERSION:
            return make_error_response(
                code="CONTRACT_VERSION_MISMATCH",
                message="contract_version 不相容",
                phase="submit",
                diag={
                    "expected": RUNTIME_CONTRACT_VERSION,
                    "actual": raw_contract_version,
                    "last_state": str(current_run_status.get("state", "")),
                    "elapsed_ms": 0,
                    "threshold_ms": 0
                }
            )
        client_task_id = str(data.get("client_task_id", "")).strip()
        if not client_task_id:
            return make_error_response(
                code="invalid_client_task_id",
                message="client_task_id 不可空白",
                phase="submit",
                diag={"last_state": str(current_run_status.get("state", "")), "elapsed_ms": 0, "threshold_ms": 0}
            )

        events = parse_start_task_timeline(data.get("timeline", []))
        incoming_round_traces, incoming_runtime_diag = _extract_runtime_diag_from_start_task(data)
        skip_mode_values = [str(row.get("skip_mode", "")).strip().lower() for row in data.get("timeline", []) if isinstance(row, dict)]
        chosen_skip_mode = next((mode for mode in skip_mode_values if mode), BUFF_SKIP_MODE_WALK)
        buff_skip_mode, deprecated_alias = normalize_buff_skip_mode(chosen_skip_mode)
        current_server_task_id = _new_server_task_id()
        stop_event.clear()
        pause_event.clear()
        current_run_thread = threading.Thread(
            target=run_timeline_background,
            args=(events, buff_skip_mode),
            daemon=True
        )
        current_run_thread.start()
        if incoming_round_traces:
            patch_timeline_runtime(round_traces=incoming_round_traces)
        if incoming_runtime_diag:
            patch_timeline_runtime(runtime_diag=incoming_runtime_diag)
        return {
            "type": "ack",
            "status": "ok",
            "mode": "timeline",
            "client_task_id": client_task_id,
            "server_task_id": current_server_task_id,
            "state": {
                "status": "running",
                "mode": "timeline"
            },
            "progress": {
                "processed_count": 0,
                "events_total": len(events)
            },
            "buff_skip_mode": buff_skip_mode,
            "buff_skip_mode_deprecated_alias": deprecated_alias,
            "message": "已收到 start_task，開始背景執行"
        }

    return make_error_response(
        code="unknown_action",
        message="未知 action/type",
        phase="request",
        diag={"last_state": str(current_run_status.get("state", "")), "elapsed_ms": 0, "threshold_ms": 0}
    )


def client_thread(conn, addr):
    print("[連線] {}".format(addr))
    buffer = b""

    def handle_line(raw_line):
        try:
            text = raw_line.decode("utf-8")
            print("[收到] {}".format(text))
            data = json.loads(text)
            response = handle_request(data)
        except Exception as e:
            response = make_error_response(
                code="exception",
                message=str(e),
                phase="request",
                diag={"last_state": str(current_run_status.get("state", "")), "elapsed_ms": 0, "threshold_ms": 0}
            )

        try:
            wire = json.dumps(response, ensure_ascii=False) + "\n"
            conn.sendall(wire.encode("utf-8"))
        except Exception as e:
            print("[回傳失敗] {}".format(e))
            raise

    try:
        while True:
            raw = conn.recv(1024 * 1024)
            if not raw:
                break

            buffer += raw
            while b"\n" in buffer:
                line, buffer = buffer.split(b"\n", 1)
                line = line.strip()
                if not line:
                    continue
                handle_line(line)

        rest = buffer.strip()
        if rest:
            handle_line(rest)
    finally:
        conn.close()
        print("[斷線] {}".format(addr))


def start_server():
    setup_gpio()

    server = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    server.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    server.bind((HOST, PORT))
    server.listen(5)

    print("Pi backend v4 listening on {}:{}".format(HOST, PORT))

    try:
        while True:
            conn, addr = server.accept()
            t = threading.Thread(target=client_thread, args=(conn, addr))
            t.daemon = True
            t.start()
    except KeyboardInterrupt:
        print("\n[中止]")
    finally:
        server.close()
        cleanup_gpio()


if __name__ == "__main__":
    start_server()
