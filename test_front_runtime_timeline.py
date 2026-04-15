import unittest
import sys
import types

if "pynput" not in sys.modules:
    pynput_stub = types.ModuleType("pynput")
    pynput_stub.keyboard = types.SimpleNamespace(Listener=object)
    sys.modules["pynput"] = pynput_stub

from front import recalculate_runtime_events_by_index


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


if __name__ == "__main__":
    unittest.main()
