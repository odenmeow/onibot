import sys
import types
import unittest

if "pynput" not in sys.modules:
    pynput_stub = types.ModuleType("pynput")
    pynput_stub.keyboard = types.SimpleNamespace(Listener=object)
    sys.modules["pynput"] = pynput_stub

from front import analyze_pr_gap_events, apply_minimum_gap_by_pairs, detect_jitter_order_risk_pairs


class FrontPrGapTaskTests(unittest.TestCase):
    def test_analyze_pr_gap_events_outputs_top_100_and_segments(self):
        events = []
        t = 0.0
        for i in range(120):
            t += 0.05
            events.append({
                "idx": i,
                "at": t,
                "type": "press" if i % 2 == 0 else "release",
                "button_id": "space"
            })
        analyzed = analyze_pr_gap_events(events)
        self.assertEqual(len(analyzed["pr_pairs"]), 100)
        self.assertEqual(analyzed["pr_pairs"][0]["pr_rank"], 0)
        self.assertEqual(analyzed["pr_pairs"][99]["pr_rank"], 99)
        self.assertEqual([s["range"] for s in analyzed["segments"]], [
            "pr0~pr25", "pr26~pr50", "pr51~pr70", "pr71~pr95", "pr95~pr99"
        ])

    def test_analyze_pr_gap_events_with_top_none_returns_all_adjacent_pairs(self):
        events = []
        t = 0.0
        for i in range(12):
            t += 0.1
            events.append({"idx": i, "at": t, "type": "press", "button_id": "x"})
        analyzed = analyze_pr_gap_events(events, top_n=None)
        self.assertEqual(len(analyzed["pr_pairs"]), 11)
        self.assertEqual(analyzed["pair_count"], 11)

    def test_apply_minimum_gap_by_pairs_shifts_idx_b_and_following(self):
        events = [
            {"idx": 0, "at": 1.00, "type": "press", "button_id": "a"},
            {"idx": 1, "at": 1.01, "type": "release", "button_id": "a"},
            {"idx": 2, "at": 1.02, "type": "press", "button_id": "b"},
            {"idx": 3, "at": 1.03, "type": "release", "button_id": "b"},
        ]
        before = analyze_pr_gap_events(events, top_n=100)
        adjusted, logs = apply_minimum_gap_by_pairs(events, before["pr_pairs"], minimum_gap=0.05)

        by_idx = {row["idx"]: row for row in adjusted}
        self.assertGreaterEqual(by_idx[1]["at"] - by_idx[0]["at"], 0.05)
        self.assertGreaterEqual(by_idx[2]["at"] - by_idx[1]["at"], 0.0)
        self.assertGreaterEqual(len(logs), 1)

    def test_apply_minimum_gap_accepts_zero(self):
        events = [
            {"idx": 0, "at": 1.0, "type": "press", "button_id": "a"},
            {"idx": 1, "at": 1.1, "type": "release", "button_id": "a"},
        ]
        before = analyze_pr_gap_events(events, top_n=100)
        adjusted, logs = apply_minimum_gap_by_pairs(events, before["pr_pairs"], minimum_gap=0.0)
        self.assertEqual(adjusted[0]["at"], 1.0)
        self.assertEqual(adjusted[1]["at"], 1.1)
        self.assertEqual(logs[0]["delta"], 0.0)

    def test_detect_jitter_order_risk_pairs_flags_overtake_and_tie(self):
        events = [
            {"idx": 0, "at": 1.00, "type": "press", "button_id": "a", "at_jitter": 0.15},
            {"idx": 1, "at": 1.10, "type": "release", "button_id": "a", "at_jitter": 0.15},
            {"idx": 2, "at": 1.25, "type": "press", "button_id": "b", "at_jitter": 0.10},
        ]
        risks = detect_jitter_order_risk_pairs(events)
        self.assertEqual(len(risks), 2)
        self.assertEqual(risks[0]["risk_level"], "overtake_risk")
        self.assertEqual(risks[1]["risk_level"], "tie_risk")


if __name__ == "__main__":
    unittest.main()
