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
    def test_runtime_snapshot_progress_contains_prediction_fields(self):
        backend.set_timeline_runtime("timeline", "running", events_total=2, loop_count=1, server_task_id="srv-1")
        backend.update_timeline_runtime_progress_prediction(
            next_expected_idx=1,
            next_expected_at_ms=1234567890,
            hint_slow_segment=True
        )
        snap = backend.get_timeline_runtime_snapshot()
        progress = snap["progress"]
        self.assertEqual(progress["next_expected_idx"], 1)
        self.assertEqual(progress["next_expected_at_ms"], 1234567890)
        self.assertTrue(progress["hint_slow_segment"])
        self.assertGreater(progress["heartbeat_interval_ms"], 0)

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

    def test_backend_uses_frontend_computed_at_without_runtime_randomization(self):
        events = [
            {"type": "press", "button": "f", "at": 1.00, "at_jitter": 0.2, "at_random_sec": 1.0},
            {"type": "release", "button": "f", "at": 1.10, "at_jitter": 0.2, "at_random_sec": 1.0},
        ]

        orig_monotonic = backend.time.monotonic
        orig_sleep = backend.safe_sleep
        orig_press = backend.press_only
        orig_release = backend.release_only
        try:
            tick = {"t": 0.0}

            def _mono():
                tick["t"] += 0.001
                return tick["t"]

            backend.time.monotonic = _mono
            backend.safe_sleep = lambda *_args, **_kwargs: None
            backend.press_only = lambda *_args, **_kwargs: None
            backend.release_only = lambda *_args, **_kwargs: None

            results = backend.run_timeline(events, buff_runtime={"next_ready_at": {}})
        finally:
            backend.time.monotonic = orig_monotonic
            backend.safe_sleep = orig_sleep
            backend.press_only = orig_press
            backend.release_only = orig_release

        self.assertEqual(results[0]["source_target_at"], 1.0)
        self.assertEqual(results[1]["source_target_at"], 1.1)

    def test_backend_respects_runtime_source_index_for_frontend_row_mapping(self):
        events = [
            {"type": "press", "button": "f", "at": 0.1, "runtime_source_index": 7},
        ]
        orig_monotonic = backend.time.monotonic
        orig_sleep = backend.safe_sleep
        orig_press = backend.press_only
        orig_release = backend.release_only
        try:
            tick = {"t": 0.0}

            def _mono():
                tick["t"] += 0.001
                return tick["t"]

            backend.time.monotonic = _mono
            backend.safe_sleep = lambda *_args, **_kwargs: None
            backend.press_only = lambda *_args, **_kwargs: None
            backend.release_only = lambda *_args, **_kwargs: None
            results = backend.run_timeline(events)
        finally:
            backend.time.monotonic = orig_monotonic
            backend.safe_sleep = orig_sleep
            backend.press_only = orig_press
            backend.release_only = orig_release

        self.assertEqual(results[0]["original_index"], 7)

    def test_skip_mode_pass_compresses_timeline_without_wait(self):
        runtime = {"next_ready_at": {}}
        events = [
            {"type": "press", "button": "f", "at": 0.0, "buff_group": "1", "buff_cycle_sec": 9.0, "buff_jitter_sec": 0.0},
            {"type": "press", "button": "f", "at": 0.5, "buff_group": "1", "buff_cycle_sec": 9.0, "buff_jitter_sec": 0.0},
        ]
        orig_monotonic = backend.time.monotonic
        orig_sleep = backend.safe_sleep
        orig_press = backend.press_only
        orig_release = backend.release_only
        try:
            tick = {"t": 0.0}
            waits = []
            backend.time.monotonic = lambda: tick.__setitem__("t", tick["t"] + 0.001) or tick["t"]
            backend.safe_sleep = lambda sec: waits.append(sec)
            backend.press_only = lambda *_args, **_kwargs: None
            backend.release_only = lambda *_args, **_kwargs: None
            results = backend.run_timeline(events, buff_runtime=runtime, buff_skip_mode=backend.BUFF_SKIP_MODE_PASS)
        finally:
            backend.time.monotonic = orig_monotonic
            backend.safe_sleep = orig_sleep
            backend.press_only = orig_press
            backend.release_only = orig_release

        self.assertEqual(results[1]["status"], "skipped_by_cooldown")
        self.assertEqual(len(waits), 0)

    def test_skip_mode_walk_keeps_wait_for_skipped_row(self):
        runtime = {"next_ready_at": {}}
        events = [
            {"type": "press", "button": "f", "at": 0.0, "buff_group": "1", "buff_cycle_sec": 9.0, "buff_jitter_sec": 0.0},
            {"type": "press", "button": "f", "at": 0.5, "buff_group": "1", "buff_cycle_sec": 9.0, "buff_jitter_sec": 0.0},
        ]
        orig_monotonic = backend.time.monotonic
        orig_sleep = backend.safe_sleep
        orig_press = backend.press_only
        orig_release = backend.release_only
        try:
            tick = {"t": 0.0}
            waits = []
            backend.time.monotonic = lambda: tick.__setitem__("t", tick["t"] + 0.001) or tick["t"]
            backend.safe_sleep = lambda sec: waits.append(sec)
            backend.press_only = lambda *_args, **_kwargs: None
            backend.release_only = lambda *_args, **_kwargs: None
            results = backend.run_timeline(events, buff_runtime=runtime, buff_skip_mode=backend.BUFF_SKIP_MODE_WALK)
        finally:
            backend.time.monotonic = orig_monotonic
            backend.safe_sleep = orig_sleep
            backend.press_only = orig_press
            backend.release_only = orig_release

        self.assertEqual(results[1]["status"], "skipped_by_cooldown")
        self.assertGreater(waits[0], 0.0)

    def test_handle_request_maps_legacy_compress_to_pass(self):
        orig_thread = backend.threading.Thread
        orig_current_run_thread = backend.current_run_thread
        try:
            backend.current_run_thread = None

            class _FakeThread:
                def __init__(self, target=None, args=(), daemon=None):
                    self.target = target
                    self.args = args
                    self.daemon = daemon

                def start(self):
                    return None

                def is_alive(self):
                    return False

            backend.threading.Thread = _FakeThread
            res = backend.handle_request({
                "type": "start_task",
                "contract_version": backend.RUNTIME_CONTRACT_VERSION,
                "client_task_id": "ct-1",
                "timeline": [{"idx": 0, "at_ms": 0, "action": "press", "btn": "f", "skip_mode": "compress"}]
            })
        finally:
            backend.threading.Thread = orig_thread
            backend.current_run_thread = orig_current_run_thread

        self.assertEqual(res["status"], "ok")
        self.assertEqual(res["buff_skip_mode"], backend.BUFF_SKIP_MODE_PASS)
        self.assertEqual(res["buff_skip_mode_deprecated_alias"], "compress")


if __name__ == "__main__":
    unittest.main()
