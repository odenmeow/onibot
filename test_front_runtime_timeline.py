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


class _FakeLabel:
    def __init__(self):
        self.text = ""
        self.background = ""
        self.visible = False
        self.place_calls = []

    def configure(self, **kwargs):
        self.text = kwargs.get("text", self.text)
        self.background = kwargs.get("background", self.background)

    def place(self, **kwargs):
        self.visible = True
        self.place_calls.append(kwargs)

    def place_forget(self):
        self.visible = False

    def destroy(self):
        self.visible = False


class _FakeTree:
    def __init__(self, children, bbox_by_iid):
        self._children = [str(v) for v in children]
        self._bbox_by_iid = bbox_by_iid
        self.deleted = []

    def get_children(self):
        return tuple(self._children)

    def bbox(self, iid, column=None):
        return self._bbox_by_iid.get(str(iid))

    def delete(self, item):
        self.deleted.append(str(item))
        self._children = [iid for iid in self._children if iid != str(item)]


class OverlayStabilityTests(unittest.TestCase):
    def setUp(self):
        self._orig_label = sys.modules["front"].tk.Label
        sys.modules["front"].tk.Label = lambda *args, **kwargs: _FakeLabel()

    def tearDown(self):
        sys.modules["front"].tk.Label = self._orig_label

    def _new_app(self):
        app = App.__new__(App)
        app._tree_font = "TkDefaultFont"
        app.tree_overlay_labels = {}
        app.runtime_recent_ok_indices = []
        app.runtime_recent_skipped_indices = []
        app.timeline_runtime_by_index = {}
        return app

    def test_overlay_keeps_previous_visible_state_when_bbox_temporarily_missing(self):
        app = self._new_app()
        app.timeline = [{"buff_group": "", "replicatedRow": 0} for _ in range(11)]
        app.timeline[9]["replicatedRow"] = 1
        app.timeline[10]["replicatedRow"] = 1
        app.tree = _FakeTree(
            children=[9, 10],
            bbox_by_iid={
                "9": (0, 90, 120, 20),
                "10": (0, 110, 120, 20),
            },
        )

        app._refresh_tree_buff_group_overlays()
        self.assertTrue(app.tree_overlay_labels["9"].visible)
        self.assertTrue(app.tree_overlay_labels["10"].visible)
        self.assertEqual(app.tree_overlay_labels["9"].background, "#fff4b3")
        self.assertEqual(app.tree_overlay_labels["10"].background, "#fff4b3")

        app.tree._bbox_by_iid = {}
        app._refresh_tree_buff_group_overlays()

        self.assertTrue(app.tree_overlay_labels["9"].visible)
        self.assertTrue(app.tree_overlay_labels["10"].visible)
        self.assertEqual(app.tree_overlay_labels["9"].place_calls[-1]["y"], 90)
        self.assertEqual(app.tree_overlay_labels["10"].place_calls[-1]["y"], 110)

    def test_overlay_state_is_stable_and_deterministic_across_repeated_refresh(self):
        app = self._new_app()
        app.timeline = [{"buff_group": "", "replicatedRow": 0} for _ in range(11)]
        app.timeline[9]["replicatedRow"] = 1
        app.timeline[10]["replicatedRow"] = 1
        app.tree = _FakeTree(
            children=[9, 10],
            bbox_by_iid={
                "9": (0, 90, 120, 20),
                "10": (0, 110, 120, 20),
            },
        )

        snapshots = []
        for _ in range(5):
            app._refresh_tree_buff_group_overlays()
            snapshots.append(
                (
                    app.tree_overlay_labels["9"].visible,
                    app.tree_overlay_labels["9"].background,
                    app.tree_overlay_labels["9"].place_calls[-1]["y"],
                    app.tree_overlay_labels["10"].visible,
                    app.tree_overlay_labels["10"].background,
                    app.tree_overlay_labels["10"].place_calls[-1]["y"],
                )
            )

        self.assertEqual(len(set(snapshots)), 1)

        app.timeline[10]["replicatedRow"] = 0
        app._refresh_tree_buff_group_overlays()
        self.assertTrue(app.tree_overlay_labels["9"].visible)
        self.assertFalse(app.tree_overlay_labels["10"].visible)


if __name__ == "__main__":
    unittest.main()
