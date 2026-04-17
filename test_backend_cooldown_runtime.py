import sys
import types
import unittest


if "RPi" not in sys.modules:
    rpi_pkg = types.ModuleType("RPi")
    gpio_stub = types.SimpleNamespace(
        BCM=11,
        OUT=0,
        LOW=0,
        HIGH=1,
        setmode=lambda *_args, **_kwargs: None,
        setwarnings=lambda *_args, **_kwargs: None,
        setup=lambda *_args, **_kwargs: None,
        output=lambda *_args, **_kwargs: None,
        cleanup=lambda *_args, **_kwargs: None,
    )
    rpi_pkg.GPIO = gpio_stub
    sys.modules["RPi"] = rpi_pkg
    sys.modules["RPi.GPIO"] = gpio_stub

import backend


class BackendCooldownRuntimeTests(unittest.TestCase):
    def test_same_group_second_event_can_be_skipped_in_same_round(self):
        runtime = {"next_ready_at": {}}
        events = [
            {"type": "press", "button": "f", "at": 0.0, "at_jitter": 0.0, "buff_group": "1", "buff_cycle_sec": 15.0, "buff_jitter_sec": 0.0},
            {"type": "press", "button": "f", "at": 0.0, "at_jitter": 0.0, "buff_group": "1", "buff_cycle_sec": 15.0, "buff_jitter_sec": 0.0},
        ]

        orig_monotonic = backend.time.monotonic
        orig_uniform = backend.random.uniform
        orig_sleep = backend.safe_sleep
        orig_press = backend.press_only
        orig_release = backend.release_only
        try:
            tick = {"t": 0.0}

            def _mono():
                tick["t"] += 0.001
                return tick["t"]

            backend.time.monotonic = _mono
            backend.random.uniform = lambda _a, _b: 0.0
            backend.safe_sleep = lambda *_args, **_kwargs: None
            backend.press_only = lambda *_args, **_kwargs: None
            backend.release_only = lambda *_args, **_kwargs: None

            results = backend.run_timeline(
                events,
                buff_runtime=runtime,
                buff_skip_mode=backend.BUFF_SKIP_MODE_WALK
            )
        finally:
            backend.time.monotonic = orig_monotonic
            backend.random.uniform = orig_uniform
            backend.safe_sleep = orig_sleep
            backend.press_only = orig_press
            backend.release_only = orig_release

        self.assertEqual(results[0]["status"], "ok")
        self.assertEqual(results[1]["status"], "skipped_by_cooldown")

    def test_at_random_sec_is_limited_by_neighbor_gap(self):
        events = [
            {"type": "press", "button": "f", "at": 1.00, "at_jitter": 0.0, "at_random_sec": 1.0},
            {"type": "release", "button": "f", "at": 1.10, "at_jitter": 0.0, "at_random_sec": 1.0},
            {"type": "press", "button": "g", "at": 1.30, "at_jitter": 0.0, "at_random_sec": 0.01},
        ]

        orig_monotonic = backend.time.monotonic
        orig_uniform = backend.random.uniform
        orig_sleep = backend.safe_sleep
        orig_press = backend.press_only
        orig_release = backend.release_only
        try:
            tick = {"t": 0.0}
            uniform_calls = []

            def _mono():
                tick["t"] += 0.001
                return tick["t"]

            def _uniform(a, b):
                uniform_calls.append((round(float(a), 6), round(float(b), 6)))
                return 0.0

            backend.time.monotonic = _mono
            backend.random.uniform = _uniform
            backend.safe_sleep = lambda *_args, **_kwargs: None
            backend.press_only = lambda *_args, **_kwargs: None
            backend.release_only = lambda *_args, **_kwargs: None

            backend.run_timeline(events, buff_runtime={"next_ready_at": {}})
        finally:
            backend.time.monotonic = orig_monotonic
            backend.random.uniform = orig_uniform
            backend.safe_sleep = orig_sleep
            backend.press_only = orig_press
            backend.release_only = orig_release

        # 第一筆：最近間距 0.1，auto_jitter 上限 0.035
        self.assertIn((-0.035, 0.035), uniform_calls)
        # 第二筆：最近間距也是 0.1，同樣限制為 0.035
        self.assertGreaterEqual(uniform_calls.count((-0.035, 0.035)), 2)
        # 第三筆：at_random_sec 很小，使用 0.01
        self.assertIn((-0.01, 0.01), uniform_calls)


if __name__ == "__main__":
    unittest.main()
