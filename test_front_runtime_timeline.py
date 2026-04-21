import unittest
import sys
import types
import random
from unittest import mock

if "pynput" not in sys.modules:
    pynput_stub = types.ModuleType("pynput")
    pynput_stub.keyboard = types.SimpleNamespace(Listener=object)
    sys.modules["pynput"] = pynput_stub

from front import (
    App,
    PiRequestError,
    ROUND_DONE_FINISH_GRACE_MS,
    recalculate_runtime_events_by_index,
    allocate_randat_blocks,
    get_buff_cell_visual_state,
    is_slot_excluded_buff_group,
    move_rows_with_ab_gap_compensation,
)


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

    def test_move_rows_with_ab_gap_compensation_keeps_tighter_side_gap(self):
        events = [
            {"at": 0.00},
            {"at": 1.00},
            {"at": 1.10},
            {"at": 2.50},
        ]
        moved, new_indexes, meta = move_rows_with_ab_gap_compensation(events, [2], "up")
        self.assertEqual(new_indexes, [1])
        self.assertEqual(moved[1]["at"], 0.1)
        self.assertEqual(meta["preserve_target"], "A")
        self.assertTrue(meta["preserved"])
        self.assertLess(meta["compensation_applied"], 0.0)

    def test_move_rows_with_ab_gap_compensation_handles_down_direction(self):
        events = [
            {"at": 0.00},
            {"at": 0.80},
            {"at": 1.20},
            {"at": 1.25},
        ]
        moved, new_indexes, meta = move_rows_with_ab_gap_compensation(events, [1], "down")
        self.assertEqual(new_indexes, [2])
        self.assertEqual(meta["direction"], "down")
        self.assertIn(meta["preserve_target"], {"A", "B", None})
        self.assertGreaterEqual(moved[2]["at"], moved[1]["at"])


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


class _FakeStringVar:
    def __init__(self, value=""):
        self._value = value

    def get(self):
        return self._value

    def set(self, value):
        self._value = value


class RuntimeDisplayTests(unittest.TestCase):
    def setUp(self):
        self._orig_showwarning = sys.modules["front"].messagebox.showwarning
        self._orig_askyesno = sys.modules["front"].messagebox.askyesno
        self.warning_calls = []
        self.ask_calls = []
        sys.modules["front"].messagebox.showwarning = (
            lambda title, msg: self.warning_calls.append((title, msg))
        )
        sys.modules["front"].messagebox.askyesno = (
            lambda title, msg: self.ask_calls.append((title, msg)) or False
        )

    def tearDown(self):
        sys.modules["front"].messagebox.showwarning = self._orig_showwarning
        sys.modules["front"].messagebox.askyesno = self._orig_askyesno

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
        app.runtime_trace_owner = {"run_id": 0, "server_task_id": ""}
        app.runtime_trace_status_note = ""
        app.runtime_latest_index = None
        app.last_runtime_signature = ""
        app.runtime_display_frozen = False
        app.runtime_manual_restore_active = False
        app.runtime_wait_ack_active = False
        app.runtime_working_timeline = []
        app.wait_ack_snapshot_timeline = []
        app.user_stop_requested = False
        app.stop_restore_prompted_task_keys = set()
        app.runtime_seen_active_task_keys = set()
        app.pre_run_timeline_snapshot = []
        app.has_pre_run_snapshot = False
        app.origin_snapshot_before_loop = []
        app.origin_version = 0
        app.runtime_version = 0
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
        app.frontend_error_main = ""
        app.last_control_error = ""
        app.last_status_error = ""
        app.last_handshake_info = {}
        app._set_restore_pre_run_button_state = App._set_restore_pre_run_button_state.__get__(app, App)
        app._ensure_runtime_editable = App._ensure_runtime_editable.__get__(app, App)
        app._freeze_runtime_for_wait_ack = App._freeze_runtime_for_wait_ack.__get__(app, App)
        app._show_runtime_view_timeline = App._show_runtime_view_timeline.__get__(app, App)
        app._monitor_after_ack = App._monitor_after_ack.__get__(app, App)
        app._preflight_start_task = App._preflight_start_task.__get__(app, App)
        app._get_ack_timeout_ms = App._get_ack_timeout_ms.__get__(app, App)
        app._build_request_diag_message = App._build_request_diag_message.__get__(app, App)
        app._get_handshake_contract_version = App._get_handshake_contract_version.__get__(app, App)
        app.stop_pi = App.stop_pi.__get__(app, App)
        app.update_current_labels = lambda: None
        app.set_frontend_error = lambda *_args, **_kwargs: None
        app.set_status = lambda *_args, **_kwargs: None
        app.config = {"pi_host": "127.0.0.1", "ack_timeout_ms": 1234}
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
        random.seed(7)
        events = [
            {"type": "press", "button": "space", "at": 0.0, "buff_group": "G1"},
            {"type": "randat", "button": "", "at": 1.0, "buff_group": ""},
            {"type": "press", "button": "x", "at": 2.0, "buff_group": "G2"},
        ]
        _working, assignments, traces, _debug = allocate_randat_blocks(events)
        self.assertEqual(len(traces), 2)
        used_slots = set()
        for trace in traces:
            self.assertIn("anchorIdx", trace)
            self.assertIn("candidateIdxList", trace)
            self.assertIn("freeCandidateIdxList", trace)
            self.assertIn("pickedIdx", trace)
            self.assertIn("pickedReason", trace)
            self.assertIn("diceValue", trace)
            self.assertIn("pickedCandidatePos", trace)
            free_candidates = trace["freeCandidateIdxList"]
            picked_idx = trace["pickedIdx"]
            if free_candidates:
                self.assertIn(picked_idx, free_candidates)
                self.assertEqual(trace["pickedReason"], "random_pick")
            else:
                self.assertEqual(trace["pickedReason"], "fallback_no_free_slot")
            self.assertNotIn(picked_idx, used_slots)
            used_slots.add(picked_idx)
            self.assertEqual(trace["reason"], trace["pickedReason"])
            group = trace["buffGroup"]
            self.assertEqual(assignments[group]["landed_index"], picked_idx)

    def test_allocate_randat_blocks_reorders_runtime_queue_with_gap_preserved(self):
        events = [
            {"type": "press", "button": "a", "at": 0.0, "buff_group": "1"},
            {"type": "release", "button": "a", "at": 0.2, "buff_group": "1"},
            {"type": "press", "button": "x", "at": 1.0, "buff_group": ""},
            {"type": "randat", "button": "", "at": 2.0, "buff_group": ""},
            {"type": "press", "button": "b", "at": 3.0, "buff_group": "2"},
            {"type": "release", "button": "b", "at": 3.3, "buff_group": "2"},
            {"type": "press", "button": "y", "at": 4.0, "buff_group": ""},
        ]

        with mock.patch("front.random.randrange", side_effect=[2, 0]):
            working, assignments, _traces, debug = allocate_randat_blocks(events)

        self.assertEqual(working[0].get("buff_group", ""), "")
        self.assertEqual(working[1].get("buff_group", ""), "1")
        self.assertEqual(working[3].get("buff_group", ""), "2")
        at_values = [float(row.get("at", 0.0)) for row in working]
        self.assertEqual(at_values, sorted(at_values))
        self.assertNotEqual(assignments["1"]["landed_index"], assignments["1"]["anchor_index"])
        self.assertEqual(debug.get("apply_order"), ["2", "1"])
        self.assertEqual([item.get("group_id") for item in debug.get("placement_ledger", [])], ["2", "1"])
        self.assertIn("1", debug.get("group_final_positions", {}))
        self.assertIn("2", debug.get("group_final_positions", {}))

    def test_allocate_randat_blocks_excludes_numeric_groups_greater_than_100(self):
        random.seed(11)
        events = [
            {"type": "press", "button": "a", "at": 0.0, "buff_group": "1"},
            {"type": "press", "button": "b", "at": 1.0, "buff_group": "101"},
            {"type": "randat", "button": "", "at": 2.0, "buff_group": ""},
            {"type": "press", "button": "c", "at": 3.0, "buff_group": "2"},
            {"type": "press", "button": "d", "at": 4.0, "buff_group": "150"},
        ]
        _working, assignments, traces, _debug = allocate_randat_blocks(events)
        self.assertIn("1", assignments)
        self.assertIn("2", assignments)
        self.assertNotIn("101", assignments)
        self.assertNotIn("150", assignments)
        traced_groups = {str(item.get("buffGroup", "")).strip() for item in traces}
        self.assertEqual(traced_groups, {"1", "2"})

    def test_running_shows_cooldown_text_in_buff_group_column(self):
        app = self._new_app()
        app.timeline = [
            {"type": "press", "button": "space", "at": 0.0, "at_jitter": 0.0, "buff_group": "A", "buff_cycle_sec": 0.0, "buff_jitter_sec": 0.0, "replicatedRow": 0}
        ]
        app.timeline_runtime_by_index = {0: {"status": "skipped_by_cooldown"}}
        app.refresh_tree()

        values = app.tree.item("0", "values")
        self.assertEqual(values[5], "A（冷卻中）")

    def test_recent_ok_green_tag_is_restored_for_non_buff_row(self):
        app = self._new_app()
        app.timeline = [
            {"type": "press", "button": "space", "at": 0.0, "at_jitter": 0.0, "buff_group": "", "buff_cycle_sec": 0.0, "buff_jitter_sec": 0.0, "replicatedRow": 0}
        ]
        app.runtime_recent_ok_indices = [0]
        app.refresh_tree()
        tags = app.tree.item("0", "tags")
        self.assertIn("runtime_ok_1", tags)

    def test_runtime_index_mapping_uses_source_index_when_randat_exists(self):
        app = self._new_app()
        app.timeline = [
            {"type": "press", "button": "space", "at": 0.0, "at_jitter": 0.0, "buff_group": "", "buff_cycle_sec": 0.0, "buff_jitter_sec": 0.0, "replicatedRow": 0},
            {"type": "randat", "button": "", "at": 0.1, "at_jitter": 0.0, "buff_group": "", "buff_cycle_sec": 0.0, "buff_jitter_sec": 0.0, "replicatedRow": 0},
            {"type": "release", "button": "space", "at": 0.2, "at_jitter": 0.0, "buff_group": "", "buff_cycle_sec": 0.0, "buff_jitter_sec": 0.0, "replicatedRow": 0},
        ]
        changed = app.update_runtime_from_status({
            "timeline_runtime": {
                "state": "running",
                "events": [{"original_index": 2, "status": "ok"}]
            }
        })
        self.assertTrue(changed)
        self.assertEqual(app.runtime_latest_index, 2)

    def test_update_runtime_accepts_next_round_traces_when_owner_run_id_unknown(self):
        app = self._new_app()
        app.runtime_round_traces = [
            {"execution_round": 1, "draw_order": 1, "buffGroup": "A", "pickedReason": "random_pick"}
        ]
        app.runtime_trace_owner = {"run_id": 0, "server_task_id": "srv-round2"}
        app.runtime_trace_status_note = "已抽籤但送出失敗"

        changed = app.update_runtime_from_status({
            "timeline_runtime": {
                "state": "running",
                "run_id": 202,
                "server_task_id": "srv-round2",
                "round_traces": [
                    {"execution_round": 2, "draw_order": 1, "buffGroup": "B", "pickedReason": "random_pick"}
                ],
                "events": [{"original_index": 0, "status": "ok"}]
            }
        })

        self.assertTrue(changed)
        self.assertEqual(app.runtime_round_traces[0]["execution_round"], 2)
        self.assertEqual(app.runtime_round_traces[0]["buffGroup"], "B")
        self.assertEqual(app.runtime_trace_owner, {"run_id": 202, "server_task_id": "srv-round2"})
        self.assertEqual(app.runtime_trace_status_note, "")

    def test_render_runtime_analysis_shows_distinct_execution_rounds(self):
        app = self._new_app()
        captured = {"text": ""}

        app.text = types.SimpleNamespace(
            delete=lambda *_args, **_kwargs: captured.__setitem__("text", ""),
            insert=lambda *_args, **_kwargs: captured.__setitem__("text", captured["text"] + (_args[1] if len(_args) > 1 else ""))
        )
        app.json_view_mode_var = types.SimpleNamespace(get=lambda: "runtime")
        app.timeline = [
            {"type": "press", "button": "a", "buff_group": "A"},
            {"type": "press", "button": "b", "buff_group": "B"},
        ]
        app.runtime_round_traces = [
            {"execution_round": 1, "draw_order": 1, "buffGroup": "A", "pickedReason": "random_pick", "picked_slot": 0},
            {"execution_round": 2, "draw_order": 1, "buffGroup": "B", "pickedReason": "random_pick", "picked_slot": 1},
        ]
        app.runtime_version = 2

        app.render_runtime_analysis(force=True)

        self.assertIn("Execution Round #1", captured["text"])
        self.assertIn("Execution Round #2", captured["text"])
        self.assertIn("picked A-slot=", captured["text"])

    def test_recent_ok_green_tag_does_not_override_buff_background(self):
        app = self._new_app()
        app.timeline = [
            {"type": "press", "button": "space", "at": 0.0, "at_jitter": 0.0, "buff_group": "A", "buff_cycle_sec": 0.0, "buff_jitter_sec": 0.0, "replicatedRow": 0}
        ]
        app.runtime_recent_ok_indices = [0]
        app.refresh_tree()
        tags = app.tree.item("0", "tags")
        self.assertIn("bg_candidate", tags)
        self.assertNotIn("runtime_ok_1", tags)

    def test_paused_uses_runtime_working_timeline_color_when_backend_events_empty(self):
        app = self._new_app()
        app.timeline = [
            {"type": "press", "button": "a", "at": 0.0, "at_jitter": 0.0, "buff_group": "1", "buff_cycle_sec": 0.0, "buff_jitter_sec": 0.0, "replicatedRow": 0},
            {"type": "press", "button": "b", "at": 1.0, "at_jitter": 0.0, "buff_group": "2", "buff_cycle_sec": 0.0, "buff_jitter_sec": 0.0, "replicatedRow": 0},
        ]
        app.timeline_runtime_info = {"state": "paused", "events": [], "processed_count": 0}
        app.runtime_working_timeline = [
            {"type": "press", "button": "a", "buff_group": "1", "runtime_landed_index": 0, "runtime_anchor_index": 0, "runtime_occupies_original": 1, "runtime_self_picked": 1, "runtime_group_color": "light_yellow"},
            {"type": "press", "button": "b", "buff_group": "2", "runtime_landed_index": 0, "runtime_anchor_index": 1, "runtime_occupies_original": 0, "runtime_self_picked": 0, "runtime_group_color": "light_blue"},
        ]
        app.refresh_tree()
        self.assertIn("bg_candidate", app.tree.item("0", "tags"))
        self.assertIn("bg_applied", app.tree.item("1", "tags"))

    def test_is_slot_excluded_buff_group(self):
        self.assertFalse(is_slot_excluded_buff_group(""))
        self.assertFalse(is_slot_excluded_buff_group("A"))
        self.assertFalse(is_slot_excluded_buff_group("100"))
        self.assertTrue(is_slot_excluded_buff_group("101"))

    def test_randat_row_does_not_auto_show_applied_blue_when_idle(self):
        app = self._new_app()
        app.timeline = [
            {"type": "randat", "button": "", "at": 1.23, "at_jitter": 0.0, "buff_group": "", "buff_cycle_sec": 0.0, "buff_jitter_sec": 0.0, "replicatedRow": 0}
        ]
        app.timeline_runtime_info = {"state": "idle", "events": []}
        app.refresh_tree()
        tags = app.tree.item("0", "tags")
        values = app.tree.item("0", "values")
        self.assertNotIn("bg_applied", tags)
        self.assertEqual(values[3], "-")

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

    def test_poll_stopped_without_user_stop_does_not_prompt_restore(self):
        app = self._new_app()
        app.has_pre_run_snapshot = True
        app.user_stop_requested = False
        payloads = [
            {"timeline_runtime": {"state": "running", "run_id": 8, "server_task_id": "srv-a", "events": []}},
            {"timeline_runtime": {"state": "stopped", "run_id": 8, "server_task_id": "srv-a", "events": []}},
        ]
        app.request_pi = lambda *_args, **_kwargs: payloads.pop(0)

        app.poll_runtime_status()
        app.poll_runtime_status()

        self.assertEqual(self.ask_calls, [])

    def test_poll_stopped_after_user_stop_prompts_restore_once(self):
        app = self._new_app()
        app.has_pre_run_snapshot = True
        app.user_stop_requested = True
        payloads = [
            {"timeline_runtime": {"state": "running", "run_id": 12, "server_task_id": "srv-z", "events": []}},
            {"timeline_runtime": {"state": "stopped", "run_id": 12, "server_task_id": "srv-z", "events": []}},
            {"timeline_runtime": {"state": "stopped", "run_id": 12, "server_task_id": "srv-z", "events": []}},
        ]
        app.request_pi = lambda *_args, **_kwargs: payloads.pop(0)

        app.poll_runtime_status()
        app.poll_runtime_status()
        app.poll_runtime_status()

        self.assertEqual(len(self.ask_calls), 1)

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

    def test_preflight_rejects_service_without_start_task(self):
        app = self._new_app()
        app.request_pi = lambda *_args, **_kwargs: {
            "status": "ok",
            "service_name": "wrong-service",
            "contract_version": "v1",
            "supports_start_task": False,
        }
        with self.assertRaises(PiRequestError) as cm:
            app._preflight_start_task()
        self.assertEqual(cm.exception.kind, "protocol_reject")

    def test_preflight_maps_connection_error_to_connect_failed(self):
        app = self._new_app()

        def _raise(*_args, **_kwargs):
            raise PiRequestError("connection_error", "socket failed")
        app.request_pi = _raise

        with self.assertRaises(PiRequestError) as cm:
            app._preflight_start_task()
        self.assertEqual(cm.exception.kind, "connect_failed")

    def test_wait_ack_freeze_keeps_snapshot_and_idle_poll_does_not_drift(self):
        app = self._new_app()
        snapshot = [{"type": "press", "button": "space", "at": 1.0, "at_jitter": 0.0, "buff_group": "", "buff_cycle_sec": 0.0, "buff_jitter_sec": 0.0, "replicatedRow": 0}]
        app.timeline = [{"type": "press", "button": "x", "at": 9.9, "at_jitter": 0.0, "buff_group": "", "buff_cycle_sec": 0.0, "buff_jitter_sec": 0.0, "replicatedRow": 0}]
        app._freeze_runtime_for_wait_ack(snapshot)
        app.request_pi = lambda *_args, **_kwargs: {"timeline_runtime": {"state": "idle", "events": []}}

        app.poll_runtime_status()

        self.assertTrue(app.runtime_display_frozen)
        self.assertEqual(app.timeline[0]["button"], "space")

    def test_ack_success_unfreezes_and_enters_running(self):
        app = self._new_app()
        app.task_monitor = {
            "phase": "wait_ack",
            "client_task_id": "ct-1",
            "last_processed_count": 99,
            "last_progress_at": 777.0
        }
        app.runtime_wait_ack_active = True
        app.runtime_display_frozen = True
        app.request_pi = lambda *_args, **_kwargs: {"timeline_runtime": {"state": "running", "events": []}}

        app._monitor_after_ack({"server_task_id": "srv-1"})
        self.assertEqual(app.task_monitor.get("last_processed_count"), 0)
        self.assertIsNone(app.task_monitor.get("last_progress_at"))
        app.poll_runtime_status()

        self.assertFalse(app.runtime_wait_ack_active)
        self.assertFalse(app.runtime_display_frozen)

    def test_loop_start_auto_switches_json_mode_to_runtime_and_renders(self):
        app = self._new_app()
        app.send_timeline_loop = App.send_timeline_loop.__get__(app, App)
        app.on_json_mode_change = App.on_json_mode_change.__get__(app, App)
        app._auto_switch_runtime_view_on_loop_start = App._auto_switch_runtime_view_on_loop_start.__get__(app, App)
        app.timeline = [{"type": "press", "button": "space", "at": 1.0}]
        app.loop_preview_pending = False
        app.front_loop_enabled = False
        app.loop_preview_cached_payload = None
        app.loop_preview_origin_snapshot = []
        app.show_warning = lambda *_args, **_kwargs: None
        app._update_runtime_control_buttons = lambda: None
        app.config["auto_switch_runtime_on_loop_start"] = True
        app.json_view_mode_var = _FakeStringVar("preview")
        render_calls = {"runtime": 0}
        app.render_runtime_analysis = lambda force=False: render_calls.__setitem__("runtime", render_calls["runtime"] + 1)
        app.prepare_events_for_send = lambda **_kwargs: (
            [{"type": "press", "button": "space", "at": 1.0}],
            [{"type": "press", "button": "space", "at": 1.0}],
            ""
        )

        app.send_timeline_loop()

        self.assertEqual(app.json_view_mode_var.get(), "runtime")
        self.assertGreaterEqual(render_calls["runtime"], 1)


class RuntimeProgressTimeoutPredictionTests(unittest.TestCase):
    def _new_monitor_app(self):
        app = App.__new__(App)
        app.task_monitor = {
            "phase": "watch_progress",
            "server_task_id": "srv-1",
            "last_state": "running",
            "last_processed_count": 1,
            "last_progress_at": 100.0,
            "failed": False
        }
        app.timeline_runtime_info = {"progress": {}}
        app.error_messages = []
        app.set_frontend_error = lambda msg: app.error_messages.append(msg)
        app.set_frontend_monitor_alert = lambda msg: app.error_messages.append(msg)
        app._mark_task_monitor_failed = App._mark_task_monitor_failed.__get__(app, App)
        app._monitor_task_status_progress = App._monitor_task_status_progress.__get__(app, App)
        app.first_event_progress_watch = {}
        app._reset_task_monitor = App._reset_task_monitor.__get__(app, App)
        return app

    def test_monitor_no_false_stall_when_first_event_is_late(self):
        app = self._new_monitor_app()
        app.task_monitor["last_progress_at"] = 10.0
        app.timeline_runtime_info["progress"] = {
            "event_time_ms": 200000,
            "next_expected_idx": 0,
            "next_expected_at_ms": 212000
        }
        with mock.patch("front.time.monotonic", return_value=17.0):
            app._monitor_task_status_progress("running", 1)
        self.assertFalse(app.task_monitor.get("failed"))
        self.assertEqual(len(app.error_messages), 0)

    def test_monitor_no_false_stall_for_long_mid_gap(self):
        app = self._new_monitor_app()
        app.timeline_runtime_info["progress"] = {
            "event_time_ms": 300000,
            "next_expected_idx": 5,
            "next_expected_at_ms": 314000
        }
        with mock.patch("front.time.monotonic", return_value=111.0):
            app._monitor_task_status_progress("running", 1)
        self.assertFalse(app.task_monitor.get("failed"))
        self.assertEqual(len(app.error_messages), 0)

    def test_running_paused_resumed_does_not_trigger_progress_stall(self):
        app = self._new_monitor_app()
        with mock.patch("front.time.monotonic", return_value=100.0):
            app._monitor_task_status_progress("running", 1)
        with mock.patch("front.time.monotonic", return_value=105.0):
            app._monitor_task_status_progress("paused", 1)
        with mock.patch("front.time.monotonic", return_value=106.0):
            app._monitor_task_status_progress("resumed", 1)
        with mock.patch("front.time.monotonic", return_value=106.1):
            app._monitor_task_status_progress("running", 1)
        self.assertFalse(app.task_monitor.get("failed"))
        self.assertEqual(len(app.error_messages), 0)


class FrontLoopRoundTransitionTests(unittest.TestCase):
    def _new_loop_app(self):
        app = App.__new__(App)
        app.front_loop_enabled = True
        app.front_loop_after_id = None
        app.front_round_state = "running"
        app.front_round_last_running_at = 10.0
        app.front_inflight_client_task_id = "ct-1"
        app.root = types.SimpleNamespace(after=lambda _ms, _cb: "after-id")
        app.request_pi = lambda *_args, **_kwargs: {"timeline_runtime": {"state": "finished"}}
        app._set_front_round_state = App._set_front_round_state.__get__(app, App)
        app._poll_front_loop_round_done = App._poll_front_loop_round_done.__get__(app, App)
        app._schedule_next_round_dispatch = lambda *_args, **_kwargs: None
        app._mark_loop_terminal = lambda: None
        app._update_runtime_control_buttons = lambda: None
        return app

    def test_finished_transition_has_grace_window(self):
        app = self._new_loop_app()
        with mock.patch("front.time.monotonic", return_value=10.0 + (ROUND_DONE_FINISH_GRACE_MS / 1000.0) - 0.01):
            app._poll_front_loop_round_done("demo")
        self.assertEqual(app.front_round_state, "running")
        self.assertEqual(app.front_inflight_client_task_id, "ct-1")

    def test_round2_waiting_ack_not_marked_terminal_after_round1_finished(self):
        app = self._new_loop_app()
        app.request_pi = lambda *_args, **_kwargs: {"timeline_runtime": {"state": "waiting_ack"}}
        app._poll_front_loop_round_done("demo")
        self.assertEqual(app.front_round_state, "running")
        self.assertEqual(app.front_inflight_client_task_id, "ct-1")


class TimelineWorkflowTests(unittest.TestCase):
    def _new_workflow_app(self):
        app = App.__new__(App)
        app.timeline = []
        app.timeline_meta = {"original_events": [], "latest_saved_events": []}
        app.current_name = "demo"
        app.current_loaded_from_saved = True
        app.timeline_runtime_info = {"events": []}
        app.timeline_runtime_by_index = {}
        app.runtime_round_traces = []
        app.runtime_trace_status_note = ""
        app.runtime_trace_owner = {"run_id": 0, "server_task_id": ""}
        app.runtime_recent_ok_indices = []
        app.runtime_recent_skipped_indices = []
        app.runtime_latest_index = None
        app.last_runtime_signature = ""
        app.runtime_move_gap_logs = []
        app.first_event_progress_watch = {}
        app.task_monitor = {}
        app.runtime_wait_ack_active = False
        app.loop_preview_pending = False
        app.loop_preview_cached_payload = None
        app.loop_preview_origin_snapshot = []
        app.runtime_working_timeline = []
        app.front_loop_enabled = False
        app.runtime_display_frozen = False
        app.runtime_manual_restore_active = False
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
        app._reset_first_event_progress_watch = App._reset_first_event_progress_watch.__get__(app, App)
        app._reset_task_monitor = App._reset_task_monitor.__get__(app, App)
        app.clear_runtime_highlight = App.clear_runtime_highlight.__get__(app, App)
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
        app._update_runtime_control_buttons = lambda: None
        app.set_status = lambda *_args, **_kwargs: None
        app.set_frontend_error = lambda *_args, **_kwargs: None
        app.validate_negative_group_monotonic_by_index = App.validate_negative_group_monotonic_by_index.__get__(app, App)
        app.recalculate_timeline_for_runtime = App.recalculate_timeline_for_runtime.__get__(app, App)
        app.config = {"buff_skip_mode": "pass"}
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
        runtime_meta = app.last_prepared_payload.get("runtime_meta", {})
        self.assertIn("placement_ledger", runtime_meta)
        self.assertIn("group_final_positions", runtime_meta)
        self.assertIn("apply_order", runtime_meta)
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

    def test_load_selected_timeline_resets_runtime_related_state(self):
        app = self._new_workflow_app()
        app.get_selected_saved_name = lambda: "legacy"
        app.load_selected_timeline = App.load_selected_timeline.__get__(app, App)
        app.normalize_meta = App.normalize_meta.__get__(app, App)
        app.new_meta = App.new_meta.__get__(app, App)
        app._is_script_switch_for_load = lambda _name: False
        app.timeline_runtime_info = {"events": [{"idx": 1}]}
        app.runtime_round_traces = [{"anchorIdx": 1}]
        app.loop_preview_pending = True
        app.loop_preview_cached_payload = {"runtime_display_events": [{"at": 1.0}]}
        app.loop_preview_origin_snapshot = [{"at": 1.0}]
        app.runtime_working_timeline = [{"at": 5.0}]
        app.runtime_display_frozen = True
        app.runtime_manual_restore_active = True

        refreshed = {"tree": 0, "preview": 0}
        app.refresh_tree = lambda: refreshed.__setitem__("tree", refreshed["tree"] + 1)
        app.refresh_preview = lambda: refreshed.__setitem__("preview", refreshed["preview"] + 1)
        statuses = []
        app.set_status = lambda msg: statuses.append(msg)

        loaded_event = {"type": "press", "button": "space", "at": 9.0, "at_jitter": 0.0, "buff_group": "", "buff_cycle_sec": 0.0, "buff_jitter_sec": 0.0, "replicatedRow": 0}
        with mock.patch("front.load_named_timeline", return_value={"name": "legacy", "events": [loaded_event], "_meta": {}}), \
             mock.patch("front.save_config"):
            app.load_selected_timeline()

        self.assertEqual(app.timeline_runtime_info, {"events": []})
        self.assertEqual(app.runtime_round_traces, [])
        self.assertFalse(app.loop_preview_pending)
        self.assertIsNone(app.loop_preview_cached_payload)
        self.assertEqual(app.loop_preview_origin_snapshot, [])
        self.assertEqual(app.runtime_working_timeline, [])
        self.assertFalse(app.runtime_display_frozen)
        self.assertFalse(app.runtime_manual_restore_active)
        self.assertEqual(refreshed["tree"], 1)
        self.assertEqual(refreshed["preview"], 1)
        self.assertTrue(statuses)
        self.assertIn("Runtime 已清空", statuses[-1])

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

    def test_prepare_events_multi_round_still_uses_origin_snapshot(self):
        app = self._new_workflow_app()
        app.timeline = [
            {"type": "press", "button": "space", "at": 1.0, "at_jitter": 0.0, "buff_group": "", "buff_cycle_sec": 0.0, "buff_jitter_sec": 0.0, "replicatedRow": 0}
        ]
        origin_snapshot = app.copy_events(app.timeline)

        with mock.patch("front.random.uniform", return_value=0.0):
            round1, _ = app.prepare_events_for_send("before_send_loop", base_events=origin_snapshot)
            # 模擬 runtime 顯示層被改寫，下一輪仍不可當作 base。
            round1[0]["at"] = 9.9
            round2, _ = app.prepare_events_for_send("before_send_loop", base_events=origin_snapshot)

        self.assertEqual(round2[0]["at"], 1.0)

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
