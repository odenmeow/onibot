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

    def test_at_random_sec_adds_runtime_randomization_with_gap_cap(self):
        events = [
            {"type": "press", "button": "f", "at": 1.0, "at_jitter": 0.0, "at_random_sec": 1.0, "buff_group": "", "buff_cycle_sec": 0.0, "buff_jitter_sec": 0.0},
            {"type": "release", "button": "f", "at": 1.4, "at_jitter": 0.0, "at_random_sec": 0.0, "buff_group": "", "buff_cycle_sec": 0.0, "buff_jitter_sec": 0.0},
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
            backend.random.uniform = lambda _a, b: b
            backend.safe_sleep = lambda *_args, **_kwargs: None
            backend.press_only = lambda *_args, **_kwargs: None
            backend.release_only = lambda *_args, **_kwargs: None

            results = backend.run_timeline(events, buff_runtime=None, buff_skip_mode=backend.BUFF_SKIP_MODE_WALK)
        finally:
            backend.time.monotonic = orig_monotonic
            backend.random.uniform = orig_uniform
            backend.safe_sleep = orig_sleep
            backend.press_only = orig_press
            backend.release_only = orig_release

        # 間距 0.4 秒，auto jitter 上限為 0.4 * 0.35 = 0.14 秒
        self.assertAlmostEqual(results[0]["source_target_at"], 1.14, places=4)


if __name__ == "__main__":
    unittest.main()
