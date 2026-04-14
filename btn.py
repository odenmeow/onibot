# -*- coding: utf-8 -*-
import time
import RPi.GPIO as GPIO

# ========== 設定區 ==========
ACTIVE_LOW = True   # 如果繼電器相反就改 False
PRESS_TIME = 0.25

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
# ===========================


def get_press_level():
    return GPIO.LOW if ACTIVE_LOW else GPIO.HIGH


def get_release_level():
    return GPIO.HIGH if ACTIVE_LOW else GPIO.LOW


PRESS_LEVEL = get_press_level()
RELEASE_LEVEL = get_release_level()


def setup():
    GPIO.setmode(GPIO.BCM)
    GPIO.setwarnings(False)

    for pin in BUTTONS.values():
        GPIO.setup(pin, GPIO.OUT)
        GPIO.output(pin, RELEASE_LEVEL)


def cleanup():
    for pin in BUTTONS.values():
        GPIO.output(pin, RELEASE_LEVEL)
    GPIO.cleanup()


def press_button(name, duration=PRESS_TIME):
    if name not in BUTTONS:
        print("[錯誤] 找不到按鍵: {}".format(name))
        return

    pin = BUTTONS[name]
    print("[測試] {} (GPIO {}) 按下 {:.2f} 秒".format(name, pin, duration))
    GPIO.output(pin, PRESS_LEVEL)
    time.sleep(duration)
    GPIO.output(pin, RELEASE_LEVEL)
    time.sleep(0.1)


def delayed_press_button(name, delay, duration):
    if name not in BUTTONS:
        print("[錯誤] 找不到按鍵: {}".format(name))
        return

    delay = abs(float(delay))
    duration = abs(float(duration))

    print("[排程] {} 將在 {:.2f} 秒後按下，持續 {:.2f} 秒".format(name, delay, duration))
    time.sleep(delay)
    press_button(name, duration)


def press_and_hold(name):
    if name not in BUTTONS:
        print("[錯誤] 找不到按鍵: {}".format(name))
        return

    pin = BUTTONS[name]
    print("[保持按下] {} (GPIO {})".format(name, pin))
    GPIO.output(pin, PRESS_LEVEL)


def release_button(name):
    if name not in BUTTONS:
        print("[錯誤] 找不到按鍵: {}".format(name))
        return

    pin = BUTTONS[name]
    print("[放開] {} (GPIO {})".format(name, pin))
    GPIO.output(pin, RELEASE_LEVEL)


def test_all():
    print("\n=== 全部輪流測試開始 ===")
    for name in BUTTONS:
        press_button(name)
    print("=== 全部輪流測試結束 ===\n")


def show_buttons():
    print("\n可用按鍵：")
    for name, pin in BUTTONS.items():
        print("  {:<6} -> GPIO {}".format(name, pin))
    print()


def show_help():
    print("指令說明：")
    print("  all                    -> 全部輪流測試")
    print("  <按鍵名稱>             -> 單次按一下")
    print("  hold <按鍵名稱>        -> 持續按住")
    print("  release <按鍵名稱>     -> 放開")
    print("  <按鍵> -延遲 -持續     -> 延遲後按住一段時間")
    print("                           例如：down -2 -2")
    print("                           表示 2 秒後按 down，持續 2 秒")
    print("  list                   -> 顯示按鍵清單")
    print("  q                      -> 離開")
    print()


def main():
    setup()
    show_buttons()
    show_help()

    try:
        while True:
            cmd = input("請輸入指令: ").strip().lower()

            if cmd == "q":
                break

            elif cmd == "all":
                test_all()

            elif cmd == "list":
                show_buttons()

            elif cmd.startswith("hold "):
                parts = cmd.split(None, 1)
                if len(parts) == 2:
                    name = parts[1]
                    press_and_hold(name)
                else:
                    print("格式錯誤，例：hold space")

            elif cmd.startswith("release "):
                parts = cmd.split(None, 1)
                if len(parts) == 2:
                    name = parts[1]
                    release_button(name)
                else:
                    print("格式錯誤，例：release space")

            else:
                parts = cmd.split()

                # 單一按鍵
                if len(parts) == 1 and parts[0] in BUTTONS:
                    press_button(parts[0])

                # 延遲 + 持續按壓，例如：down -2 -2
                elif len(parts) == 3:
                    name = parts[0]
                    delay_str = parts[1]
                    duration_str = parts[2]

                    if name not in BUTTONS:
                        print("[錯誤] 找不到按鍵: {}".format(name))
                        continue

                    try:
                        delay = abs(float(delay_str))
                        duration = abs(float(duration_str))
                    except ValueError:
                        print("格式錯誤，例：down -2 -2")
                        continue

                    delayed_press_button(name, delay, duration)

                else:
                    print("未知指令，輸入 list 查看按鍵名稱。")

    except KeyboardInterrupt:
        print("\n[中止] 使用者中斷")

    finally:
        cleanup()
        print("[結束] GPIO 已清理")


if __name__ == "__main__":
    main()