import unittest
import sys
import types

if "pynput" not in sys.modules:
    pynput_stub = types.ModuleType("pynput")
    pynput_stub.keyboard = types.SimpleNamespace(Listener=object)
    sys.modules["pynput"] = pynput_stub

from front import App, recalculate_runtime_events_by_index


def _event(at, buff_group=""):
    return {"at": at, "buff_group": buff_group}


class RecalculateRuntimeTimelineTests(unittest.TestCase):
    def test_duplicate_segment_with_same_at_does_not_flatten_following_events(self):
        events = [
            _event(5.00, "-1"),
            _event(5.08, "-1"),
            _event(5.27, "-1"),
            _event(5.00, ""),
            _event(5.08, ""),
            _event(5.27, ""),
            _event(5.30, ""),
            _event(5.80, ""),
            _event(6.10, ""),
        ]

        recalculated = recalculate_runtime_events_by_index(events, anchor_gap_sec=0.2)
        recalculated_at = [row["at"] for row in recalculated]
        follow_gaps = [round(recalculated_at[i] - recalculated_at[i - 1], 2) for i in range(4, len(recalculated_at))]

        self.assertNotEqual(recalculated_at[3:], [recalculated_at[3]] * len(recalculated_at[3:]))
        self.assertEqual(follow_gaps, [0.08, 0.19, 0.03, 0.5, 0.3])

    def test_multiple_negative_groups_follow_idx_order_chain(self):
        events = [
            _event(1.00, ""),
            _event(2.00, "-1"),
            _event(2.20, "-1"),
            _event(2.50, ""),
            _event(3.00, "-2"),
            _event(3.30, "-2"),
            _event(3.60, ""),
        ]

        recalculated = recalculate_runtime_events_by_index(events, anchor_gap_sec=0.2)
        recalculated_at = [row["at"] for row in recalculated]

        self.assertEqual(recalculated_at, [1.0, 1.2, 1.4, 2.9, 3.1, 3.4, 4.5])

    def test_follow_gap_after_negative_copy_uses_previous_normal_anchor(self):
        events = [_event(idx * 0.5, "") for idx in range(10)]
        events.extend([
            _event(3.5, "-1"),
            _event(4.0, "-1"),
            _event(4.5, "-1"),
            _event(5.0, "-1"),
            _event(5.0, ""),
        ])

        recalculated = recalculate_runtime_events_by_index(events, anchor_gap_sec=0.2)
        self.assertEqual(recalculated[14]["at"], round(recalculated[13]["at"] + 0.5, 2))

    def test_follow_gap_skips_whole_negative_prefix_before_event(self):
        events = [_event(idx * 0.5, "") for idx in range(10)]
        events[8]["buff_group"] = "-3"
        events[9]["buff_group"] = "-3"
        events.extend([
            _event(3.5, "-1"),
            _event(4.0, "-1"),
            _event(4.5, "-1"),
            _event(5.0, "-1"),
            _event(5.0, ""),
        ])

        recalculated = recalculate_runtime_events_by_index(events, anchor_gap_sec=0.2)
        self.assertEqual(recalculated[14]["at"], round(recalculated[13]["at"] + 1.5, 2))


class _FakeTree:
    def __init__(self):
        self._children = []
        self.deleted = []
        self.rows = {}
        self.seen = None

    def get_children(self):
        return tuple(self._children)

    def delete(self, item):
        self.deleted.append(str(item))
        self._children = [iid for iid in self._children if iid != str(item)]
        self.rows.pop(str(item), None)

    def insert(self, _parent, _where, iid, values, tags=()):
        row_id = str(iid)
        if row_id not in self._children:
            self._children.append(row_id)
        self.rows[row_id] = {"values": values, "tags": tuple(tags)}

    def item(self, iid, option=None):
        row = self.rows[str(iid)]
        if option is None:
            return row
        return row[option]

    def exists(self, iid):
        return str(iid) in self.rows

    def see(self, iid):
        self.seen = str(iid)


class _FakeRoot:
    def after(self, _delay, _callback):
        return "after-id"


class RuntimeDisplayTests(unittest.TestCase):
    def setUp(self):
        self._orig_showwarning = sys.modules["front"].messagebox.showwarning
        self.warning_calls = []
        sys.modules["front"].messagebox.showwarning = (
            lambda title, msg: self.warning_calls.append((title, msg))
        )

    def tearDown(self):
        sys.modules["front"].messagebox.showwarning = self._orig_showwarning

    def _new_app(self):
        app = App.__new__(App)
        app.tree = _FakeTree()
        app.root = _FakeRoot()
        app.timeline = []
        app.runtime_recent_ok_indices = []
        app.runtime_recent_skipped_indices = []
        app.timeline_runtime_by_index = {}
        app.timeline_runtime_info = {"events": []}
        app.runtime_latest_index = None
        app.last_runtime_signature = ""
        app.runtime_display_frozen = False
        app.pre_run_timeline_snapshot = []
        app.has_pre_run_snapshot = False
        app.copy_events = App.copy_events.__get__(app, App)
        app._normalize_replicated_row_flag = App._normalize_replicated_row_flag.__get__(app, App)
        app._sync_replicated_row = App._sync_replicated_row.__get__(app, App)
        app.update_runtime_from_status = App.update_runtime_from_status.__get__(app, App)
        app.refresh_tree = App.refresh_tree.__get__(app, App)
        app.poll_runtime_status = App.poll_runtime_status.__get__(app, App)
        app.restore_pre_run_state = App.restore_pre_run_state.__get__(app, App)
        app._is_runtime_readonly = App._is_runtime_readonly.__get__(app, App)
        app.clear_runtime_highlight = App.clear_runtime_highlight.__get__(app, App)
        app.refresh_preview = lambda: None
        app.focus_latest_runtime_row = lambda: None
        app.offline_mode = False
        app.conn = object()
        app.connected = True
        return app

    def test_running_shows_cooldown_text_in_buff_group_column(self):
        app = self._new_app()
        app.timeline = [
            {"type": "press", "button": "a", "at": 0.0, "at_jitter": 0.0, "buff_group": "A", "buff_cycle_sec": 0.0, "buff_jitter_sec": 0.0, "replicatedRow": 0}
        ]
        app.timeline_runtime_by_index = {0: {"status": "skipped_by_cooldown"}}
        app.refresh_tree()

        values = app.tree.item("0", "values")
        self.assertEqual(values[5], "A（冷卻中）")

    def test_poll_stopped_freezes_and_repeated_poll_does_not_refresh_table(self):
        app = self._new_app()
        app.timeline = [
            {"type": "press", "button": "a", "at": 0.0, "at_jitter": 0.0, "buff_group": "", "buff_cycle_sec": 0.0, "buff_jitter_sec": 0.0, "replicatedRow": 0}
        ]
        refresh_count = {"count": 0}

        def _count_refresh():
            refresh_count["count"] += 1
        app.refresh_tree = _count_refresh

        stopped_payload = {"timeline_runtime": {"state": "stopped", "events": []}}
        app.request_pi = lambda *_args, **_kwargs: stopped_payload

        app.poll_runtime_status()
        app.poll_runtime_status()

        self.assertTrue(app.runtime_display_frozen)
        self.assertEqual(refresh_count["count"], 0)

    def test_restore_pre_run_state_restores_snapshot(self):
        app = self._new_app()
        app.timeline = [{"type": "press", "button": "z", "at": 9.0, "at_jitter": 0.0, "buff_group": "", "buff_cycle_sec": 0.0, "buff_jitter_sec": 0.0, "replicatedRow": 0}]
        app.pre_run_timeline_snapshot = [{"type": "press", "button": "a", "at": 1.0, "at_jitter": 0.0, "buff_group": "", "buff_cycle_sec": 0.0, "buff_jitter_sec": 0.0, "replicatedRow": 0}]
        app.has_pre_run_snapshot = True
        app.runtime_display_frozen = True

        app.restore_pre_run_state()

        self.assertEqual(app.timeline[0]["button"], "a")
        self.assertFalse(app.runtime_display_frozen)

    def test_restore_without_snapshot_does_not_change_data(self):
        app = self._new_app()
        app.timeline = [{"type": "press", "button": "z", "at": 9.0, "at_jitter": 0.0, "buff_group": "", "buff_cycle_sec": 0.0, "buff_jitter_sec": 0.0, "replicatedRow": 0}]
        app.pre_run_timeline_snapshot = []
        app.has_pre_run_snapshot = False

        app.restore_pre_run_state()

        self.assertEqual(app.timeline[0]["button"], "z")
        self.assertTrue(any("沒有可恢復內容" in msg for _, msg in self.warning_calls))


if __name__ == "__main__":
    unittest.main()
