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


class _FakeTextWidget:
    def __init__(self):
        self.value = ""
        self.fg = "#b30000"

    def config(self, **kwargs):
        if "fg" in kwargs:
            self.fg = kwargs["fg"]

    def delete(self, _start, _end):
        self.value = ""

    def insert(self, _where, text):
        self.value = str(text)


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
        app.last_prepared_payload = {}
        app.front_inflight_client_task_id = ""
        app.copy_events = App.copy_events.__get__(app, App)
        app._normalize_replicated_row_flag = App._normalize_replicated_row_flag.__get__(app, App)
        app._sync_replicated_row = App._sync_replicated_row.__get__(app, App)
        app.update_runtime_from_status = App.update_runtime_from_status.__get__(app, App)
        app.render_runtime_analysis = App.render_runtime_analysis.__get__(app, App)
        app._build_runtime_consistency_key = App._build_runtime_consistency_key.__get__(app, App)
        app._consistency_key_from_runtime_meta = App._consistency_key_from_runtime_meta.__get__(app, App)
        app._consistency_key_from_runtime_status = App._consistency_key_from_runtime_status.__get__(app, App)
        app._is_runtime_key_compatible = App._is_runtime_key_compatible.__get__(app, App)
        app.refresh_tree = App.refresh_tree.__get__(app, App)
        app.poll_runtime_status = App.poll_runtime_status.__get__(app, App)
        app.is_auto_connect_enabled = App.is_auto_connect_enabled.__get__(app, App)
        app.update_auto_connect_ui = App.update_auto_connect_ui.__get__(app, App)
        app.set_auto_connect_enabled = App.set_auto_connect_enabled.__get__(app, App)
        app.toggle_auto_connect = App.toggle_auto_connect.__get__(app, App)
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

    def test_poll_runtime_status_keeps_sync_when_auto_connect_paused(self):
        app = self._new_app()
        app.config["auto_connect_enabled"] = False
        app.request_pi = mock.Mock(return_value={
            "timeline_runtime": {
                "run_id": 8,
                "state": "running",
                "processed_count": 3,
                "events": [{"original_index": 1, "status": "ok"}]
            }
        })
        app._classify_request_error = lambda *_args, **_kwargs: "request_error"
        app._set_channel_error = lambda *_args, **_kwargs: None
        app.root = _FakeRoot()

        app.poll_runtime_status()

        app.request_pi.assert_called_once_with({"action": "status"}, write_response=False, channel="status")
        self.assertEqual(app.timeline_runtime_info.get("state"), "running")
        self.assertEqual(app.timeline_runtime_info.get("processed_count"), 3)

    def test_toggle_auto_connect_persists_and_triggers_auto_connect(self):
        app = self._new_app()
        app.config["auto_connect_enabled"] = False
        app.auto_connect_state_var = _FakeStringVar("")
        app.auto_connect_toggle_btn = types.SimpleNamespace(config=lambda **kwargs: None)
        app.auto_connect = mock.Mock()
        app.close_connection = mock.Mock()
        app.set_status = mock.Mock()
        app.monitor_reconnect_pending = True

        with mock.patch("front.save_config") as save_config_mock:
            app.toggle_auto_connect()

        self.assertTrue(app.config["auto_connect_enabled"])
        app.auto_connect.assert_called_once()
        save_config_mock.assert_called_once()

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
        self.assertAlmostEqual(at_values[2] - at_values[1], 0.2, places=4)
        self.assertAlmostEqual(at_values[3] - at_values[2], 0.8, places=4)
        self.assertNotEqual(assignments["1"]["landed_index"], assignments["1"]["anchor_index"])
        self.assertEqual(debug.get("apply_order"), ["2", "1"])
        self.assertEqual([item.get("group_id") for item in debug.get("placement_ledger", [])], ["2", "1"])
        self.assertIn("1", debug.get("group_final_positions", {}))
        self.assertIn("2", debug.get("group_final_positions", {}))
        group1_info = debug.get("group_final_positions", {}).get("1", {})
        self.assertIn("final_b_range_slot", group1_info)
        self.assertIn("final_runtime_row_range", group1_info)
        self.assertEqual(group1_info.get("final_b_range"), group1_info.get("final_b_range_slot"))
        self.assertTrue(isinstance(group1_info.get("final_runtime_row_range"), list))

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

    def test_update_runtime_keeps_local_round_traces_even_if_backend_reports_new_round(self):
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
        self.assertEqual(app.runtime_round_traces[0]["execution_round"], 1)
        self.assertEqual(app.runtime_round_traces[0]["buffGroup"], "A")
        self.assertEqual(app.runtime_trace_owner, {"run_id": 0, "server_task_id": "srv-round2"})

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
        self.assertNotIn("Execution Round #2", captured["text"])
        self.assertIn("picked A-slot=", captured["text"])
        self.assertIn("final slot B", captured["text"])
        self.assertIn("final row idx", captured["text"])

    def test_render_runtime_analysis_uses_local_traces_without_backend_key_check(self):
        app = self._new_app()
        captured = {"text": ""}
        app.text = types.SimpleNamespace(
            delete=lambda *_args, **_kwargs: captured.__setitem__("text", ""),
            insert=lambda *_args, **_kwargs: captured.__setitem__("text", captured["text"] + (_args[1] if len(_args) > 1 else ""))
        )
        app.json_view_mode_var = types.SimpleNamespace(get=lambda: "runtime")
        app.runtime_round_traces = [
            {"execution_round": 2, "draw_order": 1, "buffGroup": "A", "pickedReason": "random_pick", "picked_slot": 0}
        ]
        app.timeline_runtime_info = {"run_id": 10, "server_task_id": "srv-old", "execution_round": 2}
        app.last_prepared_payload = {
            "runtime_version": 3,
            "execution_round": 3,
            "runtime_meta": {
                "run_id": 99,
                "server_task_id": "srv-new",
                "runtime_version": 3,
                "execution_round": 3,
                "apply_order": ["A"],
                "group_final_positions": {"A": {"picked_slot_a_idx": 0, "base_b_idx_before_offset": 0, "final_b_range": [0, 1]}}
            }
        }

        app.render_runtime_analysis(force=True)

        self.assertNotIn("key 不一致", captured["text"])
        self.assertTrue(("Trace diagnosis:" in captured["text"]) or ("Execution Round #1" in captured["text"]))
        self.assertNotIn("Current round final positions:", captured["text"])
        self.assertIn("Execution Round #1", captured["text"])
        self.assertNotIn("backend 未提供 round_traces", captured["text"])

    def test_render_runtime_analysis_shows_previous_round_summary_for_round2(self):
        app = self._new_app()
        captured = {"text": ""}
        app.text = types.SimpleNamespace(
            delete=lambda *_args, **_kwargs: captured.__setitem__("text", ""),
            insert=lambda *_args, **_kwargs: captured.__setitem__("text", captured["text"] + (_args[1] if len(_args) > 1 else ""))
        )
        app.json_view_mode_var = types.SimpleNamespace(get=lambda: "runtime")
        app.runtime_round_traces = [
            {
                "execution_round": 1,
                "draw_order": 1,
                "buffGroup": "A",
                "pickedReason": "random_pick",
                "picked_slot": 2,
                "placement": {"picked_slot_a_idx": 2, "base_b_idx_before_offset": 4, "final_b_range_slot": [4, 5], "final_runtime_row_range": [7, 8]},
            },
            {
                "execution_round": 2,
                "draw_order": 1,
                "buffGroup": "B",
                "pickedReason": "random_pick",
                "picked_slot": 1,
                "placement": {"picked_slot_a_idx": 1, "base_b_idx_before_offset": 2, "final_b_range_slot": [2, 3], "final_runtime_row_range": [3, 4]},
            },
        ]
        app.timeline_runtime_info = {"execution_round": 2}

        app.render_runtime_analysis(force=True)

        self.assertNotIn("Previous round final positions", captured["text"])
        self.assertIn("Execution Round #1", captured["text"])

    def test_render_runtime_analysis_shows_previous_round_summary_for_latest_round3(self):
        app = self._new_app()
        captured = {"text": ""}
        app.text = types.SimpleNamespace(
            delete=lambda *_args, **_kwargs: captured.__setitem__("text", ""),
            insert=lambda *_args, **_kwargs: captured.__setitem__("text", captured["text"] + (_args[1] if len(_args) > 1 else ""))
        )
        app.json_view_mode_var = types.SimpleNamespace(get=lambda: "runtime")
        app.runtime_round_traces = [
            {"execution_round": 1, "draw_order": 1, "buffGroup": "A", "pickedReason": "random_pick", "placement": {"picked_slot_a_idx": 0, "base_b_idx_before_offset": 0, "final_b_range_slot": [0, 1], "final_runtime_row_range": [0, 1]}},
            {"execution_round": 2, "draw_order": 1, "buffGroup": "B", "pickedReason": "random_pick", "placement": {"picked_slot_a_idx": 3, "base_b_idx_before_offset": 6, "final_b_range_slot": [6, 7], "final_runtime_row_range": [9, 10]}},
            {"execution_round": 3, "draw_order": 1, "buffGroup": "C", "pickedReason": "random_pick", "placement": {"picked_slot_a_idx": 5, "base_b_idx_before_offset": 8, "final_b_range_slot": [8, 9], "final_runtime_row_range": [12, 13]}},
        ]
        app.timeline_runtime_info = {"execution_round": 3}

        app.render_runtime_analysis(force=True)

        self.assertNotIn("Previous round final positions", captured["text"])
        self.assertIn("Execution Round #2", captured["text"])
        self.assertIn("Execution Round #1", captured["text"])
        self.assertLess(captured["text"].find("Execution Round #2"), captured["text"].find("Execution Round #1"))
        self.assertNotIn("Execution Round #3", captured["text"])

    def test_render_runtime_analysis_does_not_show_previous_summary_for_round1(self):
        app = self._new_app()
        captured = {"text": ""}
        app.text = types.SimpleNamespace(
            delete=lambda *_args, **_kwargs: captured.__setitem__("text", ""),
            insert=lambda *_args, **_kwargs: captured.__setitem__("text", captured["text"] + (_args[1] if len(_args) > 1 else ""))
        )
        app.json_view_mode_var = types.SimpleNamespace(get=lambda: "runtime")
        app.runtime_round_traces = [
            {"execution_round": 1, "draw_order": 1, "buffGroup": "A", "pickedReason": "random_pick", "placement": {"picked_slot_a_idx": 0, "base_b_idx_before_offset": 0, "final_b_range": [0, 1]}},
        ]
        app.timeline_runtime_info = {"execution_round": 1}

        app.render_runtime_analysis(force=True)

        self.assertNotIn("Previous round final positions", captured["text"])

    def test_render_runtime_analysis_supports_camel_case_execution_round(self):
        app = self._new_app()
        captured = {"text": ""}
        app.text = types.SimpleNamespace(
            delete=lambda *_args, **_kwargs: captured.__setitem__("text", ""),
            insert=lambda *_args, **_kwargs: captured.__setitem__("text", captured["text"] + (_args[1] if len(_args) > 1 else ""))
        )
        app.json_view_mode_var = types.SimpleNamespace(get=lambda: "runtime")
        app.runtime_round_traces = [
            {
                "executionRound": 1,
                "drawOrder": 1,
                "buffGroup": "A",
                "pickedReason": "random_pick",
                "placement": {"picked_slot_a_idx": 2, "base_b_idx_before_offset": 4, "final_b_range_slot": [4, 5], "final_runtime_row_range": [11, 12]},
            },
            {
                "executionRound": 2,
                "drawOrder": 1,
                "buffGroup": "B",
                "pickedReason": "random_pick",
                "placement": {"picked_slot_a_idx": 1, "base_b_idx_before_offset": 2, "final_b_range_slot": [2, 3], "final_runtime_row_range": [6, 7]},
            },
        ]
        app.timeline_runtime_info = {"executionRound": 2}

        app.render_runtime_analysis(force=True)

        self.assertNotIn("Previous round final positions", captured["text"])
        self.assertIn("Execution Round #1", captured["text"])

    def test_render_runtime_analysis_shows_slot_and_row_idx_ranges_together(self):
        app = self._new_app()
        captured = {"text": ""}
        app.text = types.SimpleNamespace(
            delete=lambda *_args, **_kwargs: captured.__setitem__("text", ""),
            insert=lambda *_args, **_kwargs: captured.__setitem__("text", captured["text"] + (_args[1] if len(_args) > 1 else ""))
        )
        app.json_view_mode_var = types.SimpleNamespace(get=lambda: "runtime")
        app.runtime_round_traces = [
            {
                "execution_round": 2,
                "draw_order": 1,
                "buffGroup": "A",
                "pickedReason": "random_pick",
                "placement": {
                    "picked_slot_a_idx": 4,
                    "base_b_idx_before_offset": 1,
                    "final_b_range_slot": [2, 3],
                    "final_runtime_row_range": [8, 9]
                },
            }
        ]
        app.timeline_runtime_info = {"execution_round": 2}

        app.render_runtime_analysis(force=True)

        self.assertIn("final slot B=2~3", captured["text"])
        self.assertIn("final row idx=8~9", captured["text"])
        self.assertNotIn("final row idx=2~3", captured["text"])

    def test_render_runtime_analysis_emits_trace_diagnosis_when_previous_round_missing(self):
        app = self._new_app()
        captured = {"text": ""}
        app.text = types.SimpleNamespace(
            delete=lambda *_args, **_kwargs: captured.__setitem__("text", ""),
            insert=lambda *_args, **_kwargs: captured.__setitem__("text", captured["text"] + (_args[1] if len(_args) > 1 else ""))
        )
        app.json_view_mode_var = types.SimpleNamespace(get=lambda: "runtime")
        app.runtime_round_traces = [
            {"execution_round": 2, "draw_order": 1, "buffGroup": "A", "pickedReason": "random_pick"},
        ]
        app.timeline_runtime_info = {"execution_round": 2}
        app.runtime_trace_status_note = "⚠ 已拒收後端舊 traces（server_task_id mismatch）"

        app.render_runtime_analysis(force=True)

        self.assertTrue(("Trace diagnosis:" in captured["text"]) or ("Execution Round #1" in captured["text"]))
        self.assertIn("server_task_id mismatch", app.runtime_trace_diagnostic)

    def test_render_runtime_analysis_consistency_only_counts_current_round(self):
        app = self._new_app()
        captured = {"text": ""}
        app.text = types.SimpleNamespace(
            delete=lambda *_args, **_kwargs: captured.__setitem__("text", ""),
            insert=lambda *_args, **_kwargs: captured.__setitem__("text", captured["text"] + (_args[1] if len(_args) > 1 else ""))
        )
        app.json_view_mode_var = types.SimpleNamespace(get=lambda: "runtime")
        app.runtime_round_traces = [
            {"execution_round": 1, "draw_order": 1, "buffGroup": "A", "pickedReason": "random_pick"},
            {"execution_round": 1, "draw_order": 2, "buffGroup": "B", "pickedReason": "random_pick"},
            {"execution_round": 2, "draw_order": 1, "buffGroup": "C", "pickedReason": "random_pick"},
        ]
        app.timeline_runtime_info = {"execution_round": 2}

        app.render_runtime_analysis(force=True)

        self.assertIn("Consistency: expected_groups=1 / actual_draws=1", captured["text"])
        self.assertNotIn("expected_groups=3 / actual_draws=3", captured["text"])

    def test_render_runtime_analysis_current_round_duplicate_group_emits_diagnosis(self):
        app = self._new_app()
        captured = {"text": ""}
        app.text = types.SimpleNamespace(
            delete=lambda *_args, **_kwargs: captured.__setitem__("text", ""),
            insert=lambda *_args, **_kwargs: captured.__setitem__("text", captured["text"] + (_args[1] if len(_args) > 1 else ""))
        )
        app.json_view_mode_var = types.SimpleNamespace(get=lambda: "runtime")
        app.runtime_round_traces = [
            {"execution_round": 2, "draw_order": 1, "buffGroup": "A", "pickedReason": "random_pick"},
            {"execution_round": 2, "draw_order": 2, "buffGroup": "A", "pickedReason": "random_pick"},
        ]

        app.render_runtime_analysis(force=True)

        self.assertIn("Trace diagnosis:", captured["text"])
        self.assertIn("current round duplicate traces: A", app.runtime_trace_diagnostic)

    def test_update_runtime_ignores_old_backend_traces_against_prepared_meta(self):
        app = self._new_app()
        app.last_prepared_payload = {
            "runtime_version": 3,
            "execution_round": 3,
            "runtime_meta": {
                "server_task_id": "srv-new",
                "runtime_version": 3,
                "execution_round": 3
            }
        }

        app.update_runtime_from_status({
            "timeline_runtime": {
                "state": "running",
                "run_id": 20,
                "server_task_id": "srv-old",
                "runtime_version": 2,
                "execution_round": 2,
                "round_traces": [
                    {"execution_round": 2, "draw_order": 1, "buffGroup": "B", "pickedReason": "random_pick"}
                ],
                "events": []
            }
        })

        self.assertEqual(app.runtime_round_traces, [])

    def test_update_runtime_ignores_backend_round_traces_and_keeps_local_trace(self):
        app = self._new_app()
        app.runtime_round_traces = [
            {"execution_round": 7, "draw_order": 1, "buffGroup": "A", "pickedReason": "random_pick"}
        ]
        app.last_prepared_payload = {
            "runtime_version": 7,
            "execution_round": 7,
            "runtime_meta": {
                "server_task_id": "srv-7",
                "runtime_version": 7,
                "execution_round": 7
            }
        }
        app.runtime_trace_status_note = ""

        app.update_runtime_from_status({
            "timeline_runtime": {
                "state": "running",
                "run_id": 42,
                "server_task_id": "srv-7",
                "runtime_version": 0,
                "execution_round": 0,
                "round_traces": [
                    {"execution_round": 1, "draw_order": 1, "buffGroup": "B", "pickedReason": "random_pick"}
                ],
                "events": []
            }
        })

        self.assertEqual(app.runtime_round_traces[0]["buffGroup"], "A")
        self.assertEqual(app.runtime_trace_status_note, "")

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
        app._set_frontend_monitor_info = lambda msg: app.error_messages.append("INFO:" + str(msg))
        app._is_ack_timeout_message = App._is_ack_timeout_message.__get__(app, App)
        app._mark_ack_timeout_recovered = App._mark_ack_timeout_recovered.__get__(app, App)
        app._mark_task_monitor_failed = App._mark_task_monitor_failed.__get__(app, App)
        app._monitor_task_status_progress = App._monitor_task_status_progress.__get__(app, App)
        app._status_matches_inflight_task = App._status_matches_inflight_task.__get__(app, App)
        app._adopt_running_task_after_ack_timeout = App._adopt_running_task_after_ack_timeout.__get__(app, App)
        app._calibrate_monitor_after_reconnect = App._calibrate_monitor_after_reconnect.__get__(app, App)
        app._monitor_after_ack = App._monitor_after_ack.__get__(app, App)
        app._set_front_round_state = App._set_front_round_state.__get__(app, App)
        app.first_event_progress_watch = {}
        app._reset_task_monitor = App._reset_task_monitor.__get__(app, App)
        app.monitor_reconnect_pending = False
        app.front_inflight_client_task_id = "ct-1"
        app.front_round_state = "waiting_ack"
        app.front_round_state_changed_at = 0.0
        app.runtime_display_frozen = True
        app.runtime_wait_ack_active = True
        app.runtime_trace_owner = {"run_id": 0, "server_task_id": ""}
        app.runtime_trace_status_note = ""
        app.frontend_error_main = ""
        app.frontend_monitor_alert = ""
        app.frontend_monitor_alert_level = ""
        app.ack_timeout_recovered = False
        app.ack_timeout_recovered_source = ""
        app.front_loop_enabled = False
        app.set_status = lambda *_args, **_kwargs: None
        app.write_text = lambda *_args, **_kwargs: None
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

    def test_ack_timeout_reconcile_adopts_inflight_running_task(self):
        app = self._new_monitor_app()
        recovered_logs = []
        app.write_text = lambda payload: recovered_logs.append(payload)
        status_response = {
            "timeline_runtime": {
                "state": "running",
                "client_task_id": "ct-1",
                "server_task_id": "srv-99"
            }
        }
        adopted = app._adopt_running_task_after_ack_timeout(status_response)
        self.assertTrue(adopted)
        self.assertEqual(app.task_monitor.get("server_task_id"), "srv-99")
        self.assertEqual(app.task_monitor.get("phase"), "wait_running")
        self.assertFalse(app.runtime_wait_ack_active)
        self.assertEqual(app.front_round_state, "running")
        self.assertTrue(app.monitor_reconnect_pending)
        self.assertTrue(app.ack_timeout_recovered)
        self.assertEqual(app.runtime_trace_status_note, "ACK timeout 已校正，後端任務持續中")
        self.assertTrue(any(item.get("ack_timeout_recovered") for item in recovered_logs))
        self.assertTrue(any(str(msg).startswith("INFO:ACK timeout 已恢復") for msg in app.error_messages))

    def test_ack_timeout_reconcile_does_not_override_loop_round_status_text(self):
        app = self._new_monitor_app()
        app.front_loop_enabled = True
        status_messages = []
        app.set_status = lambda msg: status_messages.append(msg)
        adopted = app._adopt_running_task_after_ack_timeout({
            "timeline_runtime": {
                "state": "running",
                "client_task_id": "ct-1",
                "server_task_id": "srv-99"
            }
        })
        self.assertTrue(adopted)
        self.assertEqual(status_messages, [])

    def test_ack_timeout_reconcile_does_not_write_operation_status_text(self):
        app = self._new_monitor_app()
        app.front_loop_enabled = False
        status_messages = []
        app.set_status = lambda msg: status_messages.append(msg)
        adopted = app._adopt_running_task_after_ack_timeout({
            "timeline_runtime": {
                "state": "running",
                "client_task_id": "ct-1",
                "server_task_id": "srv-99"
            }
        })
        self.assertTrue(adopted)
        self.assertEqual(status_messages, [])

    def test_reconnect_calibration_resets_progress_baseline(self):
        app = self._new_monitor_app()
        recovered_logs = []
        app.write_text = lambda payload: recovered_logs.append(payload)
        app.task_monitor.update({
            "phase": "watch_progress",
            "last_progress_at": 10.0,
            "last_processed_count": 1,
            "failed": False
        })
        app.monitor_reconnect_pending = True
        with mock.patch("front.time.monotonic", return_value=123.0):
            app._calibrate_monitor_after_reconnect("running", 4)
        self.assertEqual(app.task_monitor.get("last_processed_count"), 4)
        self.assertEqual(app.task_monitor.get("last_progress_at"), 123.0)
        self.assertFalse(app.monitor_reconnect_pending)
        self.assertTrue(app.ack_timeout_recovered)
        self.assertTrue(any(item.get("ack_timeout_recovered") for item in recovered_logs))

    def test_ack_timeout_alert_auto_downgraded_after_reconcile(self):
        app = App.__new__(App)
        app.frontend_error_text = _FakeTextWidget()
        app.frontend_error_main = ""
        app.frontend_monitor_alert = ""
        app.frontend_monitor_alert_level = ""
        app.runtime_trace_diagnostic = ""
        app.last_control_error = ""
        app.last_status_error = ""
        app.ack_timeout_recovered = False
        app.ack_timeout_recovered_source = ""
        app.runtime_trace_status_note = ""
        app.task_monitor = {
            "phase": "wait_ack",
            "client_task_id": "ct-1",
            "server_task_id": "",
            "phase_started_at": 0.0,
            "last_state": "submitted",
            "last_processed_count": 0,
            "last_progress_at": None,
            "paused_at": None,
            "failed": False
        }
        app.monitor_reconnect_pending = False
        app.front_inflight_client_task_id = "ct-1"
        app.front_round_state = "waiting_ack"
        app.front_round_state_changed_at = 0.0
        app.runtime_display_frozen = True
        app.runtime_wait_ack_active = True
        app.runtime_trace_owner = {"run_id": 0, "server_task_id": ""}
        app.write_text = lambda *_args, **_kwargs: None
        app.set_status = lambda *_args, **_kwargs: None
        app.update_error_text = App.update_error_text.__get__(app, App)
        app._render_frontend_error = App._render_frontend_error.__get__(app, App)
        app.set_frontend_monitor_alert = App.set_frontend_monitor_alert.__get__(app, App)
        app._set_frontend_monitor_info = App._set_frontend_monitor_info.__get__(app, App)
        app._is_ack_timeout_message = App._is_ack_timeout_message.__get__(app, App)
        app._mark_ack_timeout_recovered = App._mark_ack_timeout_recovered.__get__(app, App)
        app._status_matches_inflight_task = App._status_matches_inflight_task.__get__(app, App)
        app._monitor_after_ack = App._monitor_after_ack.__get__(app, App)
        app._set_front_round_state = App._set_front_round_state.__get__(app, App)
        app._adopt_running_task_after_ack_timeout = App._adopt_running_task_after_ack_timeout.__get__(app, App)

        app.set_frontend_monitor_alert("任務監控告警（ACK_TIMEOUT）")
        self.assertEqual(app.frontend_error_text.fg, "#b30000")

        adopted = app._adopt_running_task_after_ack_timeout({
            "timeline_runtime": {"state": "running", "client_task_id": "ct-1", "server_task_id": "srv-1"}
        })
        self.assertTrue(adopted)
        self.assertEqual(app.frontend_monitor_alert_level, "info")
        self.assertIn("已恢復", app.frontend_error_text.value)
        self.assertEqual(app.frontend_error_text.fg, "#3566b8")


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


class SendDelayPersistTests(unittest.TestCase):
    def _new_delay_app(self):
        app = App.__new__(App)
        app.config = {"send_delay_sec": 1.0}
        app.send_delay_entry = types.SimpleNamespace(get=lambda: "1.0")
        app.parse_send_delay_sec = App.parse_send_delay_sec.__get__(app, App)
        app.apply_send_delay_if_needed = App.apply_send_delay_if_needed.__get__(app, App)
        app.set_status = lambda *_args, **_kwargs: None
        app.root = types.SimpleNamespace(after=lambda _ms, cb: cb())
        return app

    def test_delay_override_does_not_persist_when_disabled(self):
        app = self._new_delay_app()
        called = []

        with mock.patch("front.save_config") as save_mock:
            app.apply_send_delay_if_needed(
                lambda payload: called.append(payload),
                delay_override=0.0,
                persist_config=False
            )

        self.assertEqual(app.config.get("send_delay_sec"), 1.0)
        self.assertEqual(len(called), 1)
        self.assertEqual(called[0].get("send_delay_sec"), 0.0)
        save_mock.assert_not_called()


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
        app.confirm = lambda *_args, **_kwargs: True
        app.request_pi = lambda *_args, **_kwargs: {"status": "ok"}
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
        placement_ledger = runtime_meta.get("placement_ledger", [])
        if placement_ledger:
            self.assertIn("final_b_range_slot", placement_ledger[0])
            self.assertIn("final_runtime_row_range", placement_ledger[0])
            self.assertIn("final_b_range", placement_ledger[0])

    def test_prepare_events_for_send_randat_gate_only_when_rslot_present(self):
        app = self._new_workflow_app()
        app.show_warning = lambda *_args, **_kwargs: None
        app.show_error = lambda *_args, **_kwargs: None
        app.render_prepared_payload = lambda: None
        app.render_runtime_analysis = lambda **_kwargs: None
        app.timeline = [
            {"type": "press", "button": "space", "at": 0.0, "at_jitter": 0.0, "buff_group": "1", "buff_cycle_sec": 0.0, "buff_jitter_sec": 0.0},
            {"type": "press", "button": "space", "at": 0.1, "at_jitter": 0.0, "buff_group": "", "buff_cycle_sec": 0.0, "buff_jitter_sec": 0.0},
            {"type": "press", "button": "space", "at": 0.2, "at_jitter": 0.0, "buff_group": "1", "buff_cycle_sec": 0.0, "buff_jitter_sec": 0.0},
        ]

        prepared_without_randat, _ = app.prepare_events_for_send("before_send")
        self.assertIsNotNone(prepared_without_randat)

        app.timeline.insert(1, {"type": "randat", "button": "", "at": 0.05, "at_jitter": 0.0, "buff_group": "", "buff_cycle_sec": 0.0, "buff_jitter_sec": 0.0})
        prepared_with_randat, _ = app.prepare_events_for_send("before_send")
        self.assertIsNone(prepared_with_randat)

    def test_build_start_task_payload_promotes_execution_round_runtime_fields_to_top_level(self):
        app = self._new_workflow_app()
        app._build_start_task_payload = App._build_start_task_payload.__get__(app, App)
        app._next_client_task_id = lambda: "ct-1"
        app.front_loop_round = 0
        app.last_prepared_payload = {
            "runtime_version": 9,
            "execution_round": 1,
            "round_traces": [{"execution_round": 1, "buffGroup": "A"}],
            "runtime_meta": {
                "execution_round": 0,
                "runtime_version": 0,
                "draw_result": [{"group": "A"}],
                "round_traces": [{"execution_round": 1, "buffGroup": "A"}],
                "placement_ledger": [{"group_id": "A"}],
            }
        }

        payload = app._build_start_task_payload([{"type": "press", "button": "space", "at": 0.1, "skip_mode": "pass"}])

        self.assertEqual(payload["execution_round"], 1)
        self.assertEqual(payload["runtime_version"], 9)
        self.assertNotIn("round_traces", payload)
        self.assertEqual(payload["runtime_meta"]["execution_round"], 1)
        self.assertEqual(payload["runtime_meta"]["runtime_version"], 9)
        self.assertIn("draw_result", payload["runtime_meta"])
        self.assertNotIn("round_traces", payload["runtime_meta"])
        self.assertNotIn("placement_ledger", payload["runtime_meta"])

    def test_extract_runtime_progress_brief_fallback_never_returns_round_zero(self):
        app = self._new_workflow_app()
        app._extract_runtime_progress_brief = App._extract_runtime_progress_brief.__get__(app, App)
        app.front_loop_round = 0
        app.last_prepared_payload = {"execution_round": 1}

        round_no, processed_count, events_total = app._extract_runtime_progress_brief({
            "execution_round": 0,
            "processed_count": 0,
            "events_total": 5
        })

        self.assertEqual(round_no, 1)
        self.assertEqual(processed_count, 0)
        self.assertEqual(events_total, 5)

    def test_pause_runtime_resume_and_stop_keep_same_round_display(self):
        app = self._new_workflow_app()
        app.pause_runtime = App.pause_runtime.__get__(app, App)
        app.stop_pi = App.stop_pi.__get__(app, App)
        app._extract_runtime_progress_brief = App._extract_runtime_progress_brief.__get__(app, App)
        app._mark_loop_terminal = lambda: None
        app._set_restore_pre_run_button_state = lambda: None
        app._reset_first_event_progress_watch = lambda: None
        app.show_error = lambda *_args, **_kwargs: None
        app.close_connection = lambda *_args, **_kwargs: None
        app.pi_ip_entry = types.SimpleNamespace(get=lambda: "127.0.0.1")
        app.timeline_runtime_info = {"state": "running"}
        app.has_pre_run_snapshot = False
        status_messages = []
        app.set_status = lambda msg: status_messages.append(msg)

        app.last_prepared_payload = {"execution_round": 1}

        responses = {
            "pause": {"execution_round": 0, "processed_count": 2, "events_total": 5},
            "resume": {"execution_round": 0, "processed_count": 2, "events_total": 5},
            "stop": {"execution_round": 0, "processed_count": 2, "events_total": 5},
        }

        def fake_request(payload, **_kwargs):
            action = payload.get("action")
            if action:
                return responses[action]
            return responses["stop"]

        app.request_pi = fake_request
        app.pause_runtime()
        app.timeline_runtime_info["state"] = "paused"
        app.pause_runtime()
        app.stop_pi()

        self.assertIn("已暫停（Round #1，進度 2/5）", status_messages)
        self.assertIn("已繼續（Round #1，進度 2/5）", status_messages)
        self.assertIn("已終止 Round #1（2/5）", status_messages)

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
        request_calls = []
        app.request_pi = lambda payload, **_kwargs: (request_calls.append(dict(payload)) or {"status": "ok"})

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
        self.assertIn({"action": "reset_runtime"}, request_calls)
        self.assertTrue(statuses)
        self.assertIn("Runtime 已清空", statuses[-1])

    def test_load_selected_timeline_reset_runtime_failure_shows_warning_status(self):
        app = self._new_workflow_app()
        app.get_selected_saved_name = lambda: "legacy"
        app.load_selected_timeline = App.load_selected_timeline.__get__(app, App)
        app.normalize_meta = App.normalize_meta.__get__(app, App)
        app.new_meta = App.new_meta.__get__(app, App)
        statuses = []
        app.set_status = lambda msg: statuses.append(msg)
        app.request_pi = lambda *_args, **_kwargs: {"status": "error"}

        loaded_event = {"type": "press", "button": "space", "at": 9.0, "at_jitter": 0.0, "buff_group": "", "buff_cycle_sec": 0.0, "buff_jitter_sec": 0.0, "replicatedRow": 0}
        with mock.patch("front.load_named_timeline", return_value={"name": "legacy", "events": [loaded_event], "_meta": {}}), \
             mock.patch("front.save_config"):
            app.load_selected_timeline()

        self.assertTrue(statuses)
        self.assertIn("後端未清空，可能顯示舊 round", statuses[-1])

    def test_load_selected_timeline_same_and_diff_name_share_confirm_and_reset_flow(self):
        for selected_name, loaded_name in (("demo", "demo"), ("legacy", "legacy")):
            app = self._new_workflow_app()
            app.get_selected_saved_name = lambda selected=selected_name: selected
            app.load_selected_timeline = App.load_selected_timeline.__get__(app, App)
            app.normalize_meta = App.normalize_meta.__get__(app, App)
            app.new_meta = App.new_meta.__get__(app, App)
            app.timeline = [
                {"type": "press", "button": "space", "at": 1.0, "at_jitter": 0.0, "buff_group": "", "buff_cycle_sec": 0.0, "buff_jitter_sec": 0.0, "replicatedRow": 0}
            ]
            app.timeline_histories = {
                app._history_key(): {
                    "undo": [app.copy_events(app.timeline)],
                    "redo": [app.copy_events(app.timeline)],
                }
            }

            confirm_calls = []
            app.confirm = lambda title, msg: (confirm_calls.append((title, msg)) or True)

            loaded_event = {"type": "press", "button": "space", "at": 9.0, "at_jitter": 0.0, "buff_group": "", "buff_cycle_sec": 0.0, "buff_jitter_sec": 0.0, "replicatedRow": 0}
            with mock.patch("front.load_named_timeline", return_value={"name": loaded_name, "events": [loaded_event], "_meta": {}}), \
                 mock.patch("front.save_config"):
                app.load_selected_timeline()

            self.assertEqual(len(confirm_calls), 1)
            self.assertIn("SESSION", confirm_calls[0][1])
            bucket = app.timeline_histories[app._history_key()]
            self.assertEqual(bucket["undo"], [])
            self.assertEqual(bucket["redo"], [])

    def test_load_selected_timeline_cancel_keeps_timeline_runtime_and_history_unchanged(self):
        app = self._new_workflow_app()
        app.get_selected_saved_name = lambda: "legacy"
        app.load_selected_timeline = App.load_selected_timeline.__get__(app, App)
        app.normalize_meta = App.normalize_meta.__get__(app, App)
        app.new_meta = App.new_meta.__get__(app, App)
        app.timeline = [
            {"type": "press", "button": "space", "at": 3.0, "at_jitter": 0.0, "buff_group": "", "buff_cycle_sec": 0.0, "buff_jitter_sec": 0.0, "replicatedRow": 0}
        ]
        app.timeline_runtime_info = {"events": [{"idx": 1}]}
        app.runtime_round_traces = [{"anchorIdx": 1}]
        app.loop_preview_pending = True
        app.loop_preview_cached_payload = {"runtime_display_events": [{"at": 1.0}]}
        app.loop_preview_origin_snapshot = [{"at": 1.0}]
        app.runtime_working_timeline = [{"at": 5.0}]
        app.runtime_display_frozen = True
        app.runtime_manual_restore_active = True
        app.timeline_histories = {
            app._history_key(): {
                "undo": [[{"at": 1.0}]],
                "redo": [[{"at": 2.0}]],
            }
        }
        statuses = []
        app.set_status = lambda msg: statuses.append(msg)
        app.confirm = lambda *_args, **_kwargs: False

        timeline_before = [dict(ev) for ev in app.timeline]
        runtime_info_before = {"events": [dict(item) for item in app.timeline_runtime_info["events"]]}
        round_traces_before = [dict(item) for item in app.runtime_round_traces]
        history_before = {
            app._history_key(): {
                "undo": [[dict(ev) for ev in rows] for rows in app.timeline_histories[app._history_key()]["undo"]],
                "redo": [[dict(ev) for ev in rows] for rows in app.timeline_histories[app._history_key()]["redo"]],
            }
        }

        with mock.patch("front.load_named_timeline") as load_mock, \
             mock.patch("front.save_config") as save_config_mock:
            app.load_selected_timeline()

        load_mock.assert_not_called()
        save_config_mock.assert_not_called()
        self.assertEqual(app.timeline, timeline_before)
        self.assertEqual(app.timeline_runtime_info, runtime_info_before)
        self.assertEqual(app.runtime_round_traces, round_traces_before)
        self.assertTrue(app.loop_preview_pending)
        self.assertIsNotNone(app.loop_preview_cached_payload)
        self.assertEqual(app.loop_preview_origin_snapshot, [{"at": 1.0}])
        self.assertEqual(app.runtime_working_timeline, [{"at": 5.0}])
        self.assertTrue(app.runtime_display_frozen)
        self.assertTrue(app.runtime_manual_restore_active)
        self.assertEqual(app.timeline_histories, history_before)
        self.assertIn("已取消載入", statuses[-1])

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
        app.prepare_events_for_send = lambda **_kwargs: ([
            {"type": "press", "button": "space", "at": 1.23, "at_jitter": 0.0, "buff_group": "-1", "buff_cycle_sec": 0.0, "buff_jitter_sec": 0.0, "replicatedRow": 0, "row_color": ""}
        ], "")

        app.calculate_offsets_only()

        self.assertEqual(app.timeline[0]["at"], 1.23)
        self.assertEqual(app.timeline[0]["buff_group"], "-1")
        self.assertEqual(app.timeline_meta["original_events"][0]["at"], 6.12)
        self.assertEqual(refreshed["tree"], 1)
        self.assertEqual(refreshed["preview"], 1)

    def test_calculate_offsets_only_disables_randat_gate_allocate_and_jitter(self):
        app = self._new_workflow_app()
        app.timeline = [
            {"type": "press", "button": "space", "at": 6.12, "at_jitter": 0.5, "buff_group": "-1", "buff_cycle_sec": 0.0, "buff_jitter_sec": 0.0, "replicatedRow": 0}
        ]
        captured = {}

        def fake_prepare_events_for_send(**kwargs):
            captured["kwargs"] = dict(kwargs)
            return (app.copy_events(app.timeline), "")

        app.prepare_events_for_send = fake_prepare_events_for_send
        app.calculate_offsets_only()

        self.assertIn("kwargs", captured)
        self.assertEqual(captured["kwargs"].get("run_randat_gate"), False)
        self.assertEqual(captured["kwargs"].get("run_randat_allocate"), False)
        self.assertEqual(captured["kwargs"].get("apply_at_jitter"), False)

    def test_prepare_events_for_send_can_skip_randat_gate_for_offset_only(self):
        app = self._new_workflow_app()
        app.show_warning = lambda *_args, **_kwargs: None
        app.show_error = lambda *_args, **_kwargs: None
        app.render_prepared_payload = lambda: None
        app.render_runtime_analysis = lambda **_kwargs: None
        app.timeline = [
            {"type": "press", "button": "space", "at": 0.0, "at_jitter": 0.0, "buff_group": "2", "buff_cycle_sec": 0.0, "buff_jitter_sec": 0.0},
            {"type": "randat", "button": "", "at": 0.05, "at_jitter": 0.0, "buff_group": "", "buff_cycle_sec": 0.0, "buff_jitter_sec": 0.0},
            {"type": "press", "button": "space", "at": 0.1, "at_jitter": 0.0, "buff_group": "1", "buff_cycle_sec": 0.0, "buff_jitter_sec": 0.0},
        ]

        prepared_blocked, _ = app.prepare_events_for_send("before_send")
        prepared_unblocked, _ = app.prepare_events_for_send(
            "calculate_offset_only",
            run_randat_gate=False,
            run_randat_allocate=False,
            apply_at_jitter=False
        )

        self.assertIsNone(prepared_blocked)
        self.assertIsNotNone(prepared_unblocked)
        self.assertEqual([ev.get("buff_group", "") for ev in prepared_unblocked if ev.get("type") == "press"], ["2", "1"])

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

    def test_first_loop_preview_after_load_forces_execution_round_one(self):
        app = self._new_workflow_app()
        app.send_timeline_loop = App.send_timeline_loop.__get__(app, App)
        app._auto_switch_runtime_view_on_loop_start = lambda: None
        app.render_runtime_analysis = lambda *_args, **_kwargs: None
        app.timeline = [
            {"type": "press", "button": "space", "at": 1.0, "at_jitter": 0.0, "buff_group": "", "buff_cycle_sec": 0.0, "buff_jitter_sec": 0.0, "replicatedRow": 0}
        ]
        # 模擬前一次執行到 round 2；第一次預覽仍應強制顯示 round 1。
        app.front_loop_round = 2
        app.pending_runtime_version = 2
        app.runtime_version = 2

        captured = {}

        def fake_prepare_events_for_send(**kwargs):
            captured["kwargs"] = dict(kwargs)
            return (
                app.copy_events(app.timeline),
                app.copy_events(app.timeline),
                ""
            )

        app.prepare_events_for_send = fake_prepare_events_for_send

        app.send_timeline_loop()

        self.assertTrue(app.loop_preview_pending)
        self.assertIn("kwargs", captured)
        self.assertEqual(captured["kwargs"].get("execution_round_override"), 1)

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
