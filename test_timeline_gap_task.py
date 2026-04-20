import unittest

from timeline_gap_task import run_single_task, analyze_events, apply_discrete_at_jitter


class TimelineGapTaskTests(unittest.TestCase):
    def _sample_events(self, n=120):
        rows = []
        at = 0.0
        for i in range(n):
            at += 0.05
            rows.append({
                "idx": i,
                "at": round(at, 3),
                "type": "press" if i % 2 == 0 else "release",
                "button_id": "space" if i % 3 == 0 else "x",
            })
        return rows

    def test_discrete_jitter_mode_b_is_reproducible_and_discrete(self):
        events = self._sample_events(10)
        out_a = apply_discrete_at_jitter(events, jitter_max=0.01, jitter_step=0.001, seed=11)
        out_b = apply_discrete_at_jitter(events, jitter_max=0.01, jitter_step=0.001, seed=11)
        self.assertEqual(out_a, out_b)
        for row in out_a:
            sampled = row["at_jitter_sample"]
            self.assertGreaterEqual(sampled, 0.0)
            self.assertLessEqual(sampled, 0.01)
            unit = round(sampled / 0.001)
            self.assertAlmostEqual(sampled, unit * 0.001, places=9)

    def test_analysis_outputs_pr0_to_pr99_and_default_segments(self):
        report = analyze_events(self._sample_events(120), top_n=100)
        self.assertEqual(len(report["pr_pairs"]), 100)
        self.assertEqual(report["pr_pairs"][0]["pr_rank"], 0)
        self.assertEqual(report["pr_pairs"][99]["pr_rank"], 99)
        ranges = [row["range"] for row in report["segments"]]
        self.assertEqual(ranges, ["pr0~pr25", "pr26~pr50", "pr51~pr70", "pr71~pr95", "pr95~pr99"])

    def test_manual_batch_adjustment_pushes_idx_b_and_following_events(self):
        events = self._sample_events(130)
        baseline = analyze_events(events, top_n=100)
        target_pairs = [p for p in baseline["pr_pairs"] if 0 <= p["pr_rank"] <= 2]
        self.assertEqual(len(target_pairs), 3)

        task = run_single_task(
            events=events,
            jitter_max=0.0,
            jitter_step=0.001,
            seed=0,
            manual_pr_range="pr0~pr2",
            manual_deltas=[0.03, 0.0, 0.05],
        )

        self.assertEqual(len(task["manual_adjustment_log"]), 3)
        self.assertEqual(task["manual_adjustment_log"][1]["delta"], 0.0)

        before_min = task["analysis_before"]["min_all_pairs"]
        after_min = task["analysis_after"]["min_all_pairs"]
        self.assertGreaterEqual(after_min, before_min)

        for log in task["manual_adjustment_log"]:
            self.assertGreaterEqual(log["affected_idx_range"]["count"], 1)


if __name__ == "__main__":
    unittest.main()
