# -*- coding: utf-8 -*-
import json
import random
import socket
import threading
import time
import RPi.GPIO as GPIO

HOST = "0.0.0.0"
PORT = 5000

ACTIVE_LOW = True
DEFAULT_PRESS_TIME = 0.25

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
run_lock = threading.Lock()

current_run_thread = None
current_run_status = {
    "state": "idle",   # idle / running / stopped / error
    "mode": "",
    "message": ""
}


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


def safe_sleep(seconds, check_interval=0.01):
    seconds = max(0.0, float(seconds))
    end_time = time.monotonic() + seconds
    while time.monotonic() < end_time:
        if stop_event.is_set():
            raise InterruptedError("執行已被停止")
        remain = end_time - time.monotonic()
        time.sleep(min(check_interval, max(0.0, remain)))


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
    value = float(base_value) + random.uniform(-jitter, jitter)
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


def run_timeline(events, reset_stop_event=True):
    if not isinstance(events, list):
        raise ValueError("events 必須是 list")

    if reset_stop_event:
        stop_event.clear()
    normalized = []

    for i, ev in enumerate(events):
        ev_type = ev.get("type")
        button = ev.get("button")
        at = float(ev.get("at", 0.0))
        at_jitter = abs(float(ev.get("at_jitter", 0.0)))

        if ev_type not in ("press", "release"):
            raise ValueError("第 {} 筆 event type 錯誤".format(i))
        if button not in BUTTONS:
            raise ValueError("第 {} 筆找不到按鍵: {}".format(i, button))

        actual_at = max(0.0, at + random.uniform(-at_jitter, at_jitter))
        normalized.append({
            "type": ev_type,
            "button": button,
            "at": actual_at
        })

    normalized.sort(key=lambda x: x["at"])
    results = []

    with run_lock:
        start = time.monotonic()

        for i, ev in enumerate(normalized):
            if stop_event.is_set():
                raise InterruptedError("執行已被停止")

            target = ev["at"]
            now = time.monotonic() - start
            wait_time = target - now

            if wait_time > 0:
                safe_sleep(wait_time)

            if ev["type"] == "press":
                press_only(ev["button"])
            else:
                release_only(ev["button"])

            actual_now = time.monotonic() - start
            results.append({
                "index": i,
                "type": ev["type"],
                "button": ev["button"],
                "target_at": round(target, 4),
                "actual_at": round(actual_now, 4),
                "status": "ok"
            })

    return results


def run_timeline_background(events):
    global current_run_status
    try:
        current_run_status = {
            "state": "running",
            "mode": "timeline",
            "message": "正在執行 timeline"
        }
        run_timeline(events)
        release_all()
        current_run_status = {
            "state": "idle",
            "mode": "timeline",
            "message": "timeline 執行完成"
        }
    except InterruptedError as e:
        release_all()
        current_run_status = {
            "state": "stopped",
            "mode": "timeline",
            "message": str(e)
        }
    except Exception as e:
        release_all()
        current_run_status = {
            "state": "error",
            "mode": "timeline",
            "message": str(e)
        }


def run_timeline_loop_background(events):
    global current_run_status
    loop_count = 0
    try:
        stop_event.clear()
        while not stop_event.is_set():
            loop_count += 1
            current_run_status = {
                "state": "running",
                "mode": "timeline_loop",
                "message": "正在第 {} 次循環".format(loop_count)
            }
            run_timeline(events, reset_stop_event=False)
            release_all()

        current_run_status = {
            "state": "stopped",
            "mode": "timeline_loop",
            "message": "已停止循環，共執行 {} 次".format(loop_count)
        }
    except InterruptedError as e:
        release_all()
        current_run_status = {
            "state": "stopped",
            "mode": "timeline_loop",
            "message": "{}（共執行 {} 次）".format(str(e), loop_count)
        }
    except Exception as e:
        release_all()
        current_run_status = {
            "state": "error",
            "mode": "timeline_loop",
            "message": str(e)
        }


def run_macro_background(steps):
    global current_run_status
    try:
        current_run_status = {
            "state": "running",
            "mode": "macro",
            "message": "正在執行 macro"
        }
        run_macro(steps)
        release_all()
        current_run_status = {
            "state": "idle",
            "mode": "macro",
            "message": "macro 執行完成"
        }
    except InterruptedError as e:
        release_all()
        current_run_status = {
            "state": "stopped",
            "mode": "macro",
            "message": str(e)
        }
    except Exception as e:
        release_all()
        current_run_status = {
            "state": "error",
            "mode": "macro",
            "message": str(e)
        }


def stop_current_run():
    global current_run_status
    stop_event.set()
    release_all()
    current_run_status = {
        "state": "stopped",
        "mode": current_run_status.get("mode", ""),
        "message": "已送出停止指令，並釋放所有 GPIO"
    }
    return {
        "status": "ok",
        "message": "已送出停止指令，並釋放所有 GPIO"
    }


def handle_request(data):
    global current_run_thread, current_run_status

    action = data.get("action")

    if action == "ping":
        return {"status": "ok", "message": "pong"}

    if action == "list_buttons":
        return {"status": "ok", "buttons": BUTTONS}

    if action == "status":
        return {"status": "ok", "run_status": current_run_status}

    if action == "stop":
        return stop_current_run()

    if action == "run_macro":
        if current_run_thread is not None and current_run_thread.is_alive():
            return {"status": "busy", "message": "Pi 目前已有執行中的工作"}

        steps = data.get("steps", [])
        stop_event.clear()
        current_run_thread = threading.Thread(
            target=run_macro_background,
            args=(steps,),
            daemon=True
        )
        current_run_thread.start()
        return {"status": "ok", "mode": "macro", "message": "已收到 macro，開始背景執行"}

    if action == "run_timeline":
        if current_run_thread is not None and current_run_thread.is_alive():
            return {"status": "busy", "message": "Pi 目前已有執行中的工作"}

        events = data.get("events", [])
        stop_event.clear()
        current_run_thread = threading.Thread(
            target=run_timeline_background,
            args=(events,),
            daemon=True
        )
        current_run_thread.start()
        return {"status": "ok", "mode": "timeline", "message": "已收到 timeline，開始背景執行"}

    if action == "run_timeline_loop":
        if current_run_thread is not None and current_run_thread.is_alive():
            return {"status": "busy", "message": "Pi 目前已有執行中的工作"}

        events = data.get("events", [])
        stop_event.clear()
        current_run_thread = threading.Thread(
            target=run_timeline_loop_background,
            args=(events,),
            daemon=True
        )
        current_run_thread.start()
        return {"status": "ok", "mode": "timeline_loop", "message": "已收到 timeline，開始重複執行"}

    return {"status": "error", "message": "未知 action"}


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
            response = {"status": "error", "message": str(e)}

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
