# =============================================================================
# sample_data.py – Generate realistic sample MineStar cycle/shift data
# =============================================================================
from __future__ import annotations
import io
import numpy as np
import pandas as pd


def generate_sample_csv() -> bytes:
    """Return bytes of a sample MineStar cycle report CSV."""
    rng = np.random.default_rng(42)
    n = 400

    machines  = [f"CAT-{i:02d}" for i in range(1, 9)]
    operators = ["J. Smith", "M. Johnson", "R. Williams", "K. Brown",
                 "T. Davis",  "L. Martinez", "S. Wilson",  "A. Garcia",
                 "B. Taylor", "C. Anderson"]
    shovels   = ["EX-01", "EX-02", "EX-03"]
    dumps     = ["Crusher", "Waste Dump N", "Waste Dump S"]
    shifts    = ["Day", "Night"]

    machine_ids  = rng.choice(machines,  n)
    operator_ids = rng.choice(operators, n)
    shovel_ids   = rng.choice(shovels,   n)
    dump_ids     = rng.choice(dumps,     n)
    shift_ids    = rng.choice(shifts,    n)

    # Cycle times (minutes)
    base_cycle    = rng.normal(42, 8, n).clip(20, 90)
    # Some trucks are slower
    slow_mask = np.isin(machine_ids, ["CAT-05", "CAT-06"])
    base_cycle[slow_mask] += rng.uniform(5, 12, slow_mask.sum())

    queue_time    = rng.exponential(4, n).clip(0.5, 20)
    spot_time     = rng.normal(1.5, 0.5, n).clip(0.5, 4)
    load_time     = rng.normal(3.2, 0.8, n).clip(1.5, 7)
    travel_loaded = rng.normal(16, 3, n).clip(8, 30)
    travel_empty  = rng.normal(13, 2.5, n).clip(6, 25)
    dump_time     = rng.normal(2.5, 0.6, n).clip(1.0, 6)
    total_cycle   = queue_time + spot_time + load_time + travel_loaded + travel_empty + dump_time

    target_payload = 220.0
    payload = rng.normal(target_payload, 18, n).clip(120, 280)
    # Some operators consistently underload
    ul_mask = np.isin(operator_ids, ["T. Davis", "L. Martinez"])
    payload[ul_mask] -= rng.uniform(20, 35, ul_mask.sum())

    dates = pd.date_range("2024-10-01", periods=14, freq="D")
    date_ids = rng.choice(dates.strftime("%Y-%m-%d"), n)

    df = pd.DataFrame({
        "Machine":               machine_ids,
        "Operator":              operator_ids,
        "Shift":                 shift_ids,
        "Shift Date":            date_ids,
        "Load Location":         shovel_ids,
        "Dump Location":         dump_ids,
        "Payload (t)":           payload.round(1),
        "Target Payload (t)":    target_payload,
        "Queue Time (min)":      queue_time.round(1),
        "Spot Time (min)":       spot_time.round(1),
        "Load Time (min)":       load_time.round(1),
        "Travel Loaded (min)":   travel_loaded.round(1),
        "Travel Empty (min)":    travel_empty.round(1),
        "Dump Time (min)":       dump_time.round(1),
        "Total Cycle Time (min)": total_cycle.round(1),
    })

    buf = io.StringIO()
    df.to_csv(buf, index=False)
    return buf.getvalue().encode("utf-8")
