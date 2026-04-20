# -*- coding: utf-8 -*-
import argparse
import copy
import json
import random
from typing import Dict, List, Tuple

DEFAULT_SEGMENTS = [
    (0, 25),
    (26, 50),
    (51, 70),
    (71, 95),
    (95, 99),
]


def _safe_float(value, default=0.0):
    try:
        return float(value)
    except Exception:
        return float(default)


def _safe_int(value, default=0):
    try:
        return int(value)
    except Exception:
        return int(default)


def _validate_and_normalize_events(events: List[dict]) -> List[dict]:
    normalized = []
    for i, ev in enumerate(events):
        idx = _safe_int(ev.get("idx", i), i)
        ev_type = str(ev.get("type", "")).strip().lower()
        if ev_type not in ("press", "release"):
            raise ValueError("第 {} 筆 type 必須是 press/release".format(i))
        normalized.append({
            "idx": idx,
            "at": max(0.0, _safe_float(ev.get("at", 0.0))),
            "type": ev_type,
            "button_id": str(ev.get("button_id", "")).strip(),
        })
    return normalized


def discrete_jitter_values(jitter_max: float, jitter_step: float) -> List[float]:
    jitter_max = max(0.0, _safe_float(jitter_max, 0.0))
    jitter_step = max(0.0, _safe_float(jitter_step, 0.001))
    if jitter_step <= 0:
        raise ValueError("jitter_step 必須 > 0")
    steps = int(round(jitter_max / jitter_step))
    values = [round(i * jitter_step, 10) for i in range(steps + 1)]
    if not values:
        values = [0.0]
    return values


def apply_discrete_at_jitter(events: List[dict], jitter_max: float, jitter_step: float, seed: int = 0) -> List[dict]:
    values = discrete_jitter_values(jitter_max, jitter_step)
    rng = random.Random(seed)
    output = []
    for ev in events:
        picked = values[rng.randrange(0, len(values))]
        row = dict(ev)
        row["at_jitter_mode"] = "B"
        row["at_jitter_max"] = round(max(0.0, _safe_float(jitter_max)), 10)
        row["at_jitter_step"] = round(max(0.0, _safe_float(jitter_step)), 10)
        row["at_jitter_sample"] = picked
        row["at"] = round(max(0.0, _safe_float(ev.get("at", 0.0)) + picked), 10)
        output.append(row)
    return output


def build_pair_gaps(events: List[dict]) -> Tuple[List[dict], List[dict]]:
    ordered = sorted(events, key=lambda e: (_safe_float(e.get("at", 0.0)), _safe_int(e.get("idx", 0))))
    pairs = []
    for i in range(len(ordered) - 1):
        a = ordered[i]
        b = ordered[i + 1]
        gap = round(_safe_float(b.get("at", 0.0)) - _safe_float(a.get("at", 0.0)), 10)
        pairs.append({
            "idx_a": _safe_int(a.get("idx", 0)),
            "idx_b": _safe_int(b.get("idx", 0)),
            "button_a": str(a.get("button_id", "")),
            "button_b": str(b.get("button_id", "")),
            "type_a": str(a.get("type", "")),
            "type_b": str(b.get("type", "")),
            "at_a": round(_safe_float(a.get("at", 0.0)), 10),
            "at_b": round(_safe_float(b.get("at", 0.0)), 10),
            "gap": gap,
        })
    pairs = sorted(pairs, key=lambda p: (p["gap"], p["idx_a"], p["idx_b"]))
    for rank, item in enumerate(pairs):
        item["pr_rank"] = rank
    return ordered, pairs


def _range_label(start: int, end: int) -> str:
    return "pr{}~pr{}".format(start, end)


def _pairs_for_rank_range(pairs: List[dict], start: int, end: int) -> List[dict]:
    if start > end:
        start, end = end, start
    return [p for p in pairs if start <= int(p["pr_rank"]) <= end]


def _avg_gap(items: List[dict]) -> float:
    if not items:
        return 0.0
    return round(sum(_safe_float(item.get("gap", 0.0)) for item in items) / len(items), 10)


def build_segment_stats(pairs: List[dict], segments=None) -> List[dict]:
    if segments is None:
        segments = DEFAULT_SEGMENTS
    rows = []
    for start, end in segments:
        section_pairs = _pairs_for_rank_range(pairs, start, end)
        rows.append({
            "range": _range_label(start, end),
            "avg_gap": _avg_gap(section_pairs),
            "pairs": [
                {
                    "pr_rank": int(item["pr_rank"]),
                    "idx_a": int(item["idx_a"]),
                    "idx_b": int(item["idx_b"]),
                }
                for item in section_pairs
            ]
        })
    return rows


def analyze_events(events: List[dict], top_n: int = 100) -> dict:
    validated = _validate_and_normalize_events(events)
    ordered, pairs = build_pair_gaps(validated)
    top_pairs = pairs[:max(0, int(top_n))]
    min_gap = min((_safe_float(p["gap"]) for p in pairs), default=0.0)
    return {
        "events_sorted": ordered,
        "min_all_pairs": round(min_gap, 10),
        "pr_pairs": top_pairs,
        "segments": build_segment_stats(pairs),
        "pair_count": len(pairs),
    }


def parse_pr_range(range_text: str) -> Tuple[int, int]:
    text = str(range_text or "").strip().lower().replace(" ", "")
    if not text:
        raise ValueError("pr range 不可空白")
    if "~" not in text:
        raise ValueError("pr range 格式需為 prX~prY")
    left, right = text.split("~", 1)
    if not left.startswith("pr") or not right.startswith("pr"):
        raise ValueError("pr range 格式需為 prX~prY")
    return int(left[2:]), int(right[2:])


def apply_manual_batch_deltas(events: List[dict], pairs_snapshot: List[dict], pr_range: str, deltas: List[float]) -> Tuple[List[dict], List[dict]]:
    start_rank, end_rank = parse_pr_range(pr_range)
    target_pairs = _pairs_for_rank_range(pairs_snapshot, start_rank, end_rank)
    if len(deltas) != len(target_pairs):
        raise ValueError("deltas 數量({})需等於指定範圍 pair 數量({})".format(len(deltas), len(target_pairs)))

    working = [dict(ev) for ev in _validate_and_normalize_events(events)]
    log_rows = []
    for pair, delta_raw in zip(sorted(target_pairs, key=lambda p: int(p["pr_rank"])), deltas):
        delta = max(0.0, _safe_float(delta_raw, 0.0))
        idx_b = int(pair["idx_b"])
        affected = []
        for ev in working:
            if int(ev["idx"]) >= idx_b:
                ev["at"] = round(_safe_float(ev["at"]) + delta, 10)
                affected.append(int(ev["idx"]))
        log_rows.append({
            "pr_rank": int(pair["pr_rank"]),
            "idx_a": int(pair["idx_a"]),
            "idx_b": idx_b,
            "delta": round(delta, 10),
            "affected_idx_range": {
                "from_idx": idx_b,
                "count": len(affected),
                "affected_indices": affected,
            }
        })
    return working, log_rows


def run_single_task(
    events: List[dict],
    jitter_max: float,
    jitter_step: float,
    seed: int,
    manual_pr_range: str,
    manual_deltas: List[float],
) -> dict:
    source = _validate_and_normalize_events(events)
    jittered = apply_discrete_at_jitter(source, jitter_max=jitter_max, jitter_step=jitter_step, seed=seed)
    before = analyze_events(jittered)
    adjusted_events, logs = apply_manual_batch_deltas(
        events=jittered,
        pairs_snapshot=before["pr_pairs"],
        pr_range=manual_pr_range,
        deltas=manual_deltas,
    )
    after = analyze_events(adjusted_events)
    return {
        "config": {
            "jitter_mode": "B",
            "jitter_max": round(max(0.0, _safe_float(jitter_max)), 10),
            "jitter_step": round(max(0.0, _safe_float(jitter_step)), 10),
            "seed": int(seed),
            "manual_pr_range": manual_pr_range,
        },
        "analysis_before": before,
        "manual_adjustment_log": logs,
        "analysis_after": after,
    }


def _load_json(path: str):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def _write_json(path: str, payload: dict):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)


def main(argv=None):
    parser = argparse.ArgumentParser(description="Timeline gap analysis + manual batch adjustment task")
    parser.add_argument("--input", required=True, help="input events JSON path")
    parser.add_argument("--output", required=True, help="output report JSON path")
    parser.add_argument("--jitter-max", type=float, default=0.150)
    parser.add_argument("--jitter-step", type=float, default=0.001)
    parser.add_argument("--seed", type=int, default=0)
    parser.add_argument("--manual-range", required=True, help="like pr0~pr25")
    parser.add_argument("--deltas", required=True, help="comma separated deltas, one per pair in range")
    args = parser.parse_args(argv)

    raw = _load_json(args.input)
    if isinstance(raw, dict) and "events" in raw:
        events = raw["events"]
    else:
        events = raw
    deltas = [float(x.strip()) for x in str(args.deltas).split(",") if x.strip() != ""]
    report = run_single_task(
        events=events,
        jitter_max=args.jitter_max,
        jitter_step=args.jitter_step,
        seed=args.seed,
        manual_pr_range=args.manual_range,
        manual_deltas=deltas,
    )
    _write_json(args.output, report)


if __name__ == "__main__":
    main()
