import unittest
import sys
import types
from unittest import mock

if "pynput" not in sys.modules:
    pynput_stub = types.ModuleType("pynput")
    pynput_stub.keyboard = types.SimpleNamespace(Listener=object)
    sys.modules["pynput"] = pynput_stub

from front import App, recalculate_runtime_events_by_index, allocate_randat_blocks, get_buff_cell_visual_state


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
        app.runtime_round_traces = []
        app.runtime_latest_index = None
        app.last_runtime_signature = ""
        app.runtime_display_frozen = False
        app.runtime_manual_restore_active = False
        app.pre_run_timeline_snapshot = []
        app.has_pre_run_snapshot = False
        app.copy_events = App.copy_events.__get__(app, App)
        app._normalize_replicated_row_flag = App._normalize_replicated_row_flag.__get__(app, App)
        app._sync_replicated_row = App._sync_replicated_row.__get__(app, App)
        app.update_runtime_from_status = App.update_runtime_from_status.__get__(app, App)
        app.render_runtime_analysis = App.render_runtime_analysis.__get__(app, App)
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
        app._set_restore_pre_run_button_state = App._set_restore_pre_run_button_state.__get__(app, App)
        app._ensure_runtime_editable = App._ensure_runtime_editable.__get__(app, App)
        app.stop_pi = App.stop_pi.__get__(app, App)
        app.update_current_labels = lambda: None
        app.set_frontend_error = lambda *_args, **_kwargs: None
        app.set_status = lambda *_args, **_kwargs: None
        app.config = {"pi_host": "127.0.0.1"}
        app.pi_ip_entry = types.SimpleNamespace(get=lambda: "127.0.0.1")
        app.text = types.SimpleNamespace(
            delete=lambda *_args, **_kwargs: None,
            insert=lambda *_args, **_kwargs: None
        )
        return app

    def test_visual_state_priority_keeps_background_semantics(self):
        candidate_tags = get_buff_cell_visual_state(
            is_candidate=True,
            is_applied=False,
            is_running=True,
            is_focus=True
        )
        applied_tags = get_buff_cell_visual_state(
            is_candidate=False,
            is_applied=True,
            is_running=True,
            is_focus=True
        )
        self.assertIn("bg_candidate", candidate_tags)
        self.assertNotIn("bg_applied", candidate_tags)
        self.assertIn("ring_running", candidate_tags)
        self.assertIn("focus_hint", candidate_tags)
        self.assertIn("bg_applied", applied_tags)
        self.assertNotIn("bg_candidate", applied_tags)

    def test_allocate_randat_blocks_generates_round_trace_reason(self):
        events = [
            {"type": "press", "button": "space", "at": 0.0, "buff_group": "G1"},
            {"type": "randat", "button": "", "at": 1.0, "buff_group": ""},
            {"type": "press", "button": "x", "at": 2.0, "buff_group": "G2"},
        ]
        _working, _assignments, traces = allocate_randat_blocks(events)
        self.assertEqual(len(traces), 2)
        self.assertEqual(traces[0]["result"], "kept")
        self.assertEqual(traces[0]["reason"], "already_applied")
        self.assertEqual(traces[1]["result"], "kept")
        self.assertEqual(traces[1]["reason"], "already_applied")

    def test_running_shows_cooldown_text_in_buff_group_column(self):
        app = self._new_app()
        app.timeline = [
            {"type": "press", "button": "space", "at": 0.0, "at_jitter": 0.0, "buff_group": "A", "buff_cycle_sec": 0.0, "buff_jitter_sec": 0.0, "replicatedRow": 0}
        ]
        app.timeline_runtime_by_index = {0: {"status": "skipped_by_cooldown"}}
        app.refresh_tree()

        values = app.tree.item("0", "values")
        self.assertEqual(values[5], "A（冷卻中）")

    def test_poll_stopped_freezes_and_repeated_poll_does_not_refresh_table(self):
        app = self._new_app()
        app.timeline = [
            {"type": "press", "button": "space", "at": 0.0, "at_jitter": 0.0, "buff_group": "", "buff_cycle_sec": 0.0, "buff_jitter_sec": 0.0, "replicatedRow": 0}
        ]
        app.has_pre_run_snapshot = True
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

    def test_poll_stopped_after_manual_restore_does_not_refreeze(self):
        app = self._new_app()
        app.has_pre_run_snapshot = True
        app.runtime_manual_restore_active = True
        app.runtime_display_frozen = False
        app.request_pi = lambda *_args, **_kwargs: {"timeline_runtime": {"state": "stopped", "events": []}}

        app.poll_runtime_status()

        self.assertFalse(app.runtime_display_frozen)

    def test_restore_pre_run_state_restores_snapshot(self):
        app = self._new_app()
        app.timeline = [{"type": "press", "button": "z", "at": 9.0, "at_jitter": 0.0, "buff_group": "", "buff_cycle_sec": 0.0, "buff_jitter_sec": 0.0, "replicatedRow": 0}]
        app.pre_run_timeline_snapshot = [{"type": "press", "button": "space", "at": 1.0, "at_jitter": 0.0, "buff_group": "", "buff_cycle_sec": 0.0, "buff_jitter_sec": 0.0, "replicatedRow": 0}]
        app.has_pre_run_snapshot = True
        app.runtime_display_frozen = True

        app.restore_pre_run_state()

        self.assertEqual(app.timeline[0]["button"], "space")
        self.assertFalse(app.runtime_display_frozen)

    def test_restore_without_snapshot_does_not_change_data(self):
        app = self._new_app()
        app.timeline = [{"type": "press", "button": "z", "at": 9.0, "at_jitter": 0.0, "buff_group": "", "buff_cycle_sec": 0.0, "buff_jitter_sec": 0.0, "replicatedRow": 0}]
        app.pre_run_timeline_snapshot = []
        app.has_pre_run_snapshot = False

        app.restore_pre_run_state()

        self.assertEqual(app.timeline[0]["button"], "z")
        self.assertTrue(any("沒有可恢復內容" in msg for _, msg in self.warning_calls))

    def test_stop_without_snapshot_keeps_runtime_editable(self):
        app = self._new_app()
        app.has_pre_run_snapshot = False
        app.runtime_display_frozen = True
        app.request_pi = lambda *_args, **_kwargs: {"status": "ok"}
        app.restore_pre_run_btn = types.SimpleNamespace(config=lambda **_kwargs: None)

        app.stop_pi()

        self.assertFalse(app.runtime_display_frozen)
        self.assertTrue(app._ensure_runtime_editable())

    def test_stop_with_snapshot_freezes_runtime_display(self):
        app = self._new_app()
        app.has_pre_run_snapshot = True
        app.runtime_display_frozen = False
        app.request_pi = lambda *_args, **_kwargs: {"status": "ok"}
        states = []
        app.restore_pre_run_btn = types.SimpleNamespace(config=lambda **kwargs: states.append(kwargs.get("state")))

        app.stop_pi()

        self.assertTrue(app.runtime_display_frozen)
        self.assertIn("normal", states)


class TimelineWorkflowTests(unittest.TestCase):
    def _new_workflow_app(self):
        app = App.__new__(App)
        app.timeline = []
        app.timeline_meta = {"original_events": [], "latest_saved_events": []}
        app.current_name = "demo"
        app.current_loaded_from_saved = True
        app.runtime_display_frozen = False
        app.has_pre_run_snapshot = False
        app.root = types.SimpleNamespace()
        app.name_entry = types.SimpleNamespace(get=lambda: "demo", delete=lambda *_: None, insert=lambda *_: None)
        app.saved_listbox = types.SimpleNamespace(selection_clear=lambda *_: None, selection_set=lambda *_: None, activate=lambda *_: None)
        app.normalize_event_schema = App.normalize_event_schema.__get__(app, App)
        app.copy_events = App.copy_events.__get__(app, App)
        app.validate_negative_group_monotonic_by_index = App.validate_negative_group_monotonic_by_index.__get__(app, App)
        app.get_unsupported_buttons = App.get_unsupported_buttons.__get__(app, App)
        app.prepare_events_for_send = App.prepare_events_for_send.__get__(app, App)
        app.calculate_offsets_only = App.calculate_offsets_only.__get__(app, App)
        app._allocate_auto_negative_group = App._allocate_auto_negative_group.__get__(app, App)
        app._history_key = App._history_key.__get__(app, App)
        app._history_bucket = App._history_bucket.__get__(app, App)
        app._begin_timeline_change = App._begin_timeline_change.__get__(app, App)
        app._finalize_timeline_change = App._finalize_timeline_change.__get__(app, App)
        app._reset_timeline_history = App._reset_timeline_history.__get__(app, App)
        app.undo_timeline = App.undo_timeline.__get__(app, App)
        app.redo_timeline = App.redo_timeline.__get__(app, App)
        app._ensure_runtime_editable = lambda: True
        app.save_current_timeline = App.save_current_timeline.__get__(app, App)
        app.analyze = App.analyze.__get__(app, App)
        app.restore_original_timeline = App.restore_original_timeline.__get__(app, App)
        app.restore_latest_saved_timeline = App.restore_latest_saved_timeline.__get__(app, App)
        app.get_manual_offset_sec = lambda: 0.2
        app.refresh_tree = lambda: None
        app.refresh_preview = lambda: None
        app.update_current_labels = lambda: None
        app.mark_timeline_dirty = lambda: None
        app.refresh_saved_list = lambda: None
        app.select_saved_name = lambda _name: None
        app.set_status = lambda *_args, **_kwargs: None
        app.set_frontend_error = lambda *_args, **_kwargs: None
        app.validate_negative_group_monotonic_by_index = App.validate_negative_group_monotonic_by_index.__get__(app, App)
        app.recalculate_timeline_for_runtime = App.recalculate_timeline_for_runtime.__get__(app, App)
        app.config = {}
        return app

    def test_auto_negative_group_reuses_minus_one_after_delete(self):
        app = self._new_workflow_app()
        app.timeline = [{"buff_group": "-1"}, {"buff_group": "-2"}]
        app.timeline = [{"buff_group": "-2"}]
        self.assertEqual(app._allocate_auto_negative_group(), "-1")

    def test_auto_negative_group_assigns_minus_two_when_minus_one_exists(self):
        app = self._new_workflow_app()
        app.timeline = [{"buff_group": "-1"}, {"buff_group": "A"}]
        self.assertEqual(app._allocate_auto_negative_group(), "-2")

    def test_prepare_events_for_send_is_pure_and_does_not_persist(self):
        app = self._new_workflow_app()
        app.timeline = [
            {"type": "press", "button": "space", "at": 1.0, "at_jitter": 0.0, "buff_group": "", "buff_cycle_sec": 0.0, "buff_jitter_sec": 0.0, "replicatedRow": 0}
        ]
        app.timeline_meta["original_events"] = app.copy_events(app.timeline)
        timeline_before = [dict(ev) for ev in app.timeline]
        original_before = app.copy_events(app.timeline_meta["original_events"])

        with mock.patch("front.save_named_timeline") as save_mock:
            prepared, _ = app.prepare_events_for_send("before_send")

        self.assertEqual(app.timeline, timeline_before)
        self.assertEqual(app.timeline_meta["original_events"], original_before)
        self.assertIsNotNone(prepared)
        save_mock.assert_not_called()

    def test_save_updates_latest_saved_only(self):
        app = self._new_workflow_app()
        app.timeline = [
            {"type": "press", "button": "space", "at": 1.0, "at_jitter": 0.0, "buff_group": "", "buff_cycle_sec": 0.0, "buff_jitter_sec": 0.0, "replicatedRow": 0}
        ]
        app.timeline_meta["original_events"] = [
            {"type": "press", "button": "space", "at": 0.5, "at_jitter": 0.0, "buff_group": "", "buff_cycle_sec": 0.0, "buff_jitter_sec": 0.0, "replicatedRow": 0}
        ]

        with mock.patch("front.save_named_timeline") as save_mock, \
             mock.patch("front.save_config"), \
             mock.patch("front.os.path.exists", return_value=False):
            app.save_current_timeline()

        self.assertEqual(app.timeline_meta["original_events"][0]["at"], 0.5)
        self.assertTrue(app.timeline_meta["latest_saved_events"])
        save_mock.assert_called_once()

    def test_save_same_name_as_new_version_keeps_original_events(self):
        app = self._new_workflow_app()
        app.timeline = [
            {"type": "press", "button": "space", "at": 3.0, "at_jitter": 0.0, "buff_group": "", "buff_cycle_sec": 0.0, "buff_jitter_sec": 0.0, "replicatedRow": 0}
        ]
        app.timeline_meta["original_events"] = [
            {"type": "press", "button": "space", "at": 1.0, "at_jitter": 0.0, "buff_group": "", "buff_cycle_sec": 0.0, "buff_jitter_sec": 0.0, "replicatedRow": 0}
        ]
        original_before = [dict(ev) for ev in app.timeline_meta["original_events"]]

        with mock.patch("front.save_named_timeline") as save_mock, \
             mock.patch("front.save_config"), \
             mock.patch("front.os.path.exists", return_value=True), \
             mock.patch("front.messagebox.askyesnocancel", return_value=False):
            app.save_current_timeline()

        self.assertEqual(app.timeline_meta["original_events"], original_before)
        self.assertNotEqual(app.timeline_meta["latest_saved_events"][0]["at"], original_before[0]["at"])
        save_mock.assert_called_once()

    def test_save_same_name_full_replace_rebuilds_original_events(self):
        app = self._new_workflow_app()
        app.timeline = [
            {"type": "press", "button": "space", "at": 4.0, "at_jitter": 0.0, "buff_group": "", "buff_cycle_sec": 0.0, "buff_jitter_sec": 0.0, "replicatedRow": 0}
        ]
        app.timeline_meta["original_events"] = [
            {"type": "press", "button": "space", "at": 1.0, "at_jitter": 0.0, "buff_group": "", "buff_cycle_sec": 0.0, "buff_jitter_sec": 0.0, "replicatedRow": 0}
        ]

        with mock.patch("front.save_named_timeline") as save_mock, \
             mock.patch("front.save_config"), \
             mock.patch("front.os.path.exists", return_value=True), \
             mock.patch("front.messagebox.askyesnocancel", return_value=True):
            app.save_current_timeline()

        self.assertEqual(app.timeline_meta["original_events"][0]["at"], 4.0)
        save_mock.assert_called_once()

    def test_analyze_without_new_events_chooses_original_or_latest_saved(self):
        app = self._new_workflow_app()
        front_mod = sys.modules["front"]
        original_events_backup = list(front_mod.events)
        original_recording_start = front_mod.recording_start
        try:
            front_mod.events = []
            front_mod.recording_start = 0.0
            app.timeline_meta = {
                "original_events": [{"type": "press", "button": "space", "at": 1.0, "at_jitter": 0.0, "buff_group": "", "buff_cycle_sec": 0.0, "buff_jitter_sec": 0.0, "replicatedRow": 0}],
                "latest_saved_events": [{"type": "press", "button": "space", "at": 2.0, "at_jitter": 0.0, "buff_group": "", "buff_cycle_sec": 0.0, "buff_jitter_sec": 0.0, "replicatedRow": 0}],
            }

            with mock.patch("front.messagebox.askyesnocancel", return_value=True):
                app.analyze()
                self.assertEqual(app.timeline[0]["at"], 1.0)

            with mock.patch("front.messagebox.askyesnocancel", return_value=False):
                app.analyze()
                self.assertEqual(app.timeline[0]["at"], 2.0)
        finally:
            front_mod.events = original_events_backup
            front_mod.recording_start = original_recording_start

    def test_load_legacy_timeline_backfills_original_events(self):
        app = self._new_workflow_app()
        app.get_selected_saved_name = lambda: "legacy"
        app.load_selected_timeline = App.load_selected_timeline.__get__(app, App)
        app.normalize_meta = App.normalize_meta.__get__(app, App)
        app.new_meta = App.new_meta.__get__(app, App)

        legacy_event = {"type": "press", "button": "space", "at": 9.0, "at_jitter": 0.0, "buff_group": "", "buff_cycle_sec": 0.0, "buff_jitter_sec": 0.0, "replicatedRow": 0}
        with mock.patch("front.load_named_timeline", return_value={"name": "legacy", "events": [legacy_event], "_meta": {}}), \
             mock.patch("front.save_config"):
            app.load_selected_timeline()

        self.assertEqual(app.timeline_meta["original_events"][0]["at"], 9.0)

    def test_analyze_without_new_events_only_latest_saved_shows_clear_status(self):
        app = self._new_workflow_app()
        front_mod = sys.modules["front"]
        original_events_backup = list(front_mod.events)
        original_recording_start = front_mod.recording_start
        try:
            front_mod.events = []
            front_mod.recording_start = 0.0
            app.timeline_meta = {
                "original_events": [],
                "latest_saved_events": [{"type": "press", "button": "space", "at": 2.0, "at_jitter": 0.0, "buff_group": "", "buff_cycle_sec": 0.0, "buff_jitter_sec": 0.0, "replicatedRow": 0}],
            }
            with mock.patch("front.messagebox.showinfo") as info_mock:
                app.analyze()
            self.assertEqual(app.timeline[0]["at"], 2.0)
            info_mock.assert_called_once()
        finally:
            front_mod.events = original_events_backup
            front_mod.recording_start = original_recording_start

    def test_calculate_offsets_only_refreshes_table_without_touching_original_slot(self):
        app = self._new_workflow_app()
        app.timeline = [
            {"type": "press", "button": "space", "at": 6.12, "at_jitter": 0.0, "buff_group": "-1", "buff_cycle_sec": 0.0, "buff_jitter_sec": 0.0, "replicatedRow": 0}
        ]
        app.timeline_meta["original_events"] = app.copy_events(app.timeline)
        refreshed = {"tree": 0, "preview": 0}
        app.refresh_tree = lambda: refreshed.__setitem__("tree", refreshed["tree"] + 1)
        app.refresh_preview = lambda: refreshed.__setitem__("preview", refreshed["preview"] + 1)
        app.prepare_events_for_send = lambda action_reason="calculate_offset_only": ([
            {"type": "press", "button": "space", "at": 1.23, "at_jitter": 0.0, "buff_group": "-1", "buff_cycle_sec": 0.0, "buff_jitter_sec": 0.0, "replicatedRow": 0, "row_color": ""}
        ], "")

        app.calculate_offsets_only()

        self.assertEqual(app.timeline[0]["at"], 1.23)
        self.assertEqual(app.timeline[0]["buff_group"], "")
        self.assertEqual(app.timeline_meta["original_events"][0]["at"], 6.12)
        self.assertEqual(refreshed["tree"], 1)
        self.assertEqual(refreshed["preview"], 1)

    def test_undo_redo_groups_multi_row_change_in_one_step(self):
        app = self._new_workflow_app()
        app.timeline = [
            {"type": "press", "button": "space", "at": 1.0, "at_jitter": 0.0, "buff_group": "", "buff_cycle_sec": 0.0, "buff_jitter_sec": 0.0, "replicatedRow": 0},
            {"type": "release", "button": "space", "at": 1.2, "at_jitter": 0.0, "buff_group": "", "buff_cycle_sec": 0.0, "buff_jitter_sec": 0.0, "replicatedRow": 0},
        ]
        before = app.copy_events(app.timeline)
        snapshot = app._begin_timeline_change()
        app.timeline.extend([
            {"type": "press", "button": "x", "at": 2.0, "at_jitter": 0.0, "buff_group": "", "buff_cycle_sec": 0.0, "buff_jitter_sec": 0.0, "replicatedRow": 0},
            {"type": "release", "button": "x", "at": 2.2, "at_jitter": 0.0, "buff_group": "", "buff_cycle_sec": 0.0, "buff_jitter_sec": 0.0, "replicatedRow": 0},
        ])
        app._finalize_timeline_change(snapshot)
        self.assertEqual(len(app.timeline), 4)

        app.undo_timeline()
        self.assertEqual(app.timeline, before)

        app.redo_timeline()
        self.assertEqual(len(app.timeline), 4)


if __name__ == "__main__":
    unittest.main()
