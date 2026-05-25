# =============================================================================
# data_processor.py – Load, detect columns, and analyse MineStar exports
# =============================================================================
from __future__ import annotations
import logging
from typing import Optional

import numpy as np
import pandas as pd

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Column detection patterns (keyword → standardised field name)
# ---------------------------------------------------------------------------
COLUMN_PATTERNS = {
    "machine":          ["machine", "truck", "equipment id", "asset", "unit id", "vehicle"],
    "operator":         ["operator", "driver", "personnel"],
    "date":             ["date", "shift date", "activity date"],
    "shift":            ["shift"],
    "total_cycle_time": ["total cycle", "cycle time", "total time (min"],
    "queue_time":       ["queue time", "waiting time", "wait time"],
    "load_time":        ["load time", "loading time", "load duration"],
    "travel_loaded":    ["travel loaded", "haul time", "loaded travel", "travel full"],
    "travel_empty":     ["travel empty", "return time", "empty travel", "travel return"],
    "dump_time":        ["dump time", "dumping time", "dump duration"],
    "spot_time":        ["spot time", "spotting"],
    "payload":          ["payload (t", "gross payload", "net payload", "load weight", "payload t"],
    "target_payload":   ["target payload", "nominal payload", "rated payload", "design payload"],
    "load_location":    ["load location", "shovel", "excavator", "load site", "source"],
    "dump_location":    ["dump location", "dump site", "destination", "crusher", "waste dump"],
    "operating_hours":  ["operating hours", "productive hours", "working hours", "engine hours"],
    "idle_hours":       ["idle hours", "standby hours", "idle time"],
    "down_hours":       ["down hours", "downtime hours", "maintenance hours", "repair hours"],
    "scheduled_hours":  ["scheduled hours", "available hours", "total hours", "calendar hours"],
    "availability_pct": ["physical availability", "mechanical availability", "availability %", "avail %"],
    "utilization_pct":  ["use of availability", "utilization %", "utilisation %", "util %"],
}


def detect_columns(df: pd.DataFrame) -> dict:
    col_map: dict = {}
    cols_lower = {c.lower().strip(): c for c in df.columns}
    for field, patterns in COLUMN_PATTERNS.items():
        for pattern in patterns:
            for col_lower, col_orig in cols_lower.items():
                if pattern in col_lower:
                    col_map[field] = col_orig
                    break
            if field in col_map:
                break
    return col_map


def get_sheet_names(file) -> list[str]:
    """Return sheet names for Excel files; empty list for CSV."""
    name = getattr(file, "name", "").lower()
    if name.endswith(".csv"):
        return []
    try:
        file.seek(0)
        xf = pd.ExcelFile(file)
        return xf.sheet_names
    except Exception:
        return []


def load_data(file, sheet_name: str | None = None) -> pd.DataFrame:
    """Load CSV or a specific Excel sheet into a DataFrame."""
    name = getattr(file, "name", "").lower()
    try:
        if name.endswith(".csv"):
            for enc in ("utf-8", "utf-8-sig", "latin-1"):
                try:
                    file.seek(0)
                    df = pd.read_csv(file, encoding=enc, thousands=",")
                    df.columns = df.columns.str.strip()
                    return df
                except UnicodeDecodeError:
                    continue
        else:
            file.seek(0)
            df = pd.read_excel(file, sheet_name=sheet_name, thousands=",")
            df.columns = df.columns.str.strip()
            return df
    except Exception:
        pass
    raise ValueError("Could not read file. Ensure it is a valid CSV or Excel export from MineStar.")


def _col(df: pd.DataFrame, col_map: dict, key: str) -> Optional[pd.Series]:
    c = col_map.get(key)
    return df[c] if c and c in df.columns else None


def _num(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s, errors="coerce")


# ---------------------------------------------------------------------------
# Cycle Time Analysis
# ---------------------------------------------------------------------------
def analyze_cycle_times(df: pd.DataFrame, col_map: dict) -> dict:
    result: dict = {"available": False}
    cycle_s = _col(df, col_map, "total_cycle_time")
    if cycle_s is None:
        return result

    d = df.copy()
    d["_ct"] = _num(cycle_s)
    d = d[d["_ct"].notna() & (d["_ct"] > 0)]
    if len(d) == 0:
        return result

    result.update({
        "available":        True,
        "total_cycles":     len(d),
        "avg_cycle":        d["_ct"].mean(),
        "median_cycle":     d["_ct"].median(),
        "min_cycle":        d["_ct"].min(),
        "max_cycle":        d["_ct"].max(),
        "std_cycle":        d["_ct"].std(),
    })

    bd: dict = {}
    for key in ["queue_time", "load_time", "travel_loaded", "travel_empty", "dump_time", "spot_time"]:
        s = _col(d, col_map, key)
        if s is not None:
            bd[key] = _num(s).mean()
    result["time_breakdown"] = bd

    mach_s = _col(d, col_map, "machine")
    if mach_s is not None:
        d["_m"] = mach_s.astype(str)
        bm = (d.groupby("_m")
              .agg(avg_cycle=("_ct", "mean"), cycles=("_ct", "count"))
              .reset_index().rename(columns={"_m": "Machine"})
              .sort_values("avg_cycle", ascending=False))
        result["by_machine"] = bm

    shift_s = _col(d, col_map, "shift")
    if shift_s is not None:
        d["_sh"] = shift_s.astype(str)
        bs = (d.groupby("_sh")
              .agg(avg_cycle=("_ct", "mean"), cycles=("_ct", "count"))
              .reset_index().rename(columns={"_sh": "Shift"}))
        result["by_shift"] = bs

    result["distribution"] = d["_ct"].values
    return result


# ---------------------------------------------------------------------------
# Payload Analysis
# ---------------------------------------------------------------------------
def analyze_payload(df: pd.DataFrame, col_map: dict) -> dict:
    result: dict = {"available": False}
    pay_s = _col(df, col_map, "payload")
    if pay_s is None:
        return result

    d = df.copy()
    d["_pl"] = _num(pay_s)
    d = d[d["_pl"].notna() & (d["_pl"] > 0)]
    if len(d) == 0:
        return result

    result.update({
        "available":    True,
        "total_loads":  len(d),
        "avg_payload":  d["_pl"].mean(),
        "median_payload": d["_pl"].median(),
        "total_tonnes": d["_pl"].sum(),
        "target_payload": None,
    })

    tgt_s = _col(d, col_map, "target_payload")
    if tgt_s is not None:
        d["_tgt"] = _num(tgt_s)
        v = d.dropna(subset=["_tgt"])
        if len(v) > 0:
            avg_tgt = v["_tgt"].mean()
            result["target_payload"]  = avg_tgt
            result["pct_overloaded"]  = (v["_pl"] > v["_tgt"] * 1.10).mean() * 100
            result["pct_underloaded"] = (v["_pl"] < v["_tgt"] * 0.90).mean() * 100
            result["pct_on_target"]   = 100 - result["pct_overloaded"] - result["pct_underloaded"]
            underloads = v[v["_pl"] < v["_tgt"] * 0.90]
            result["tonnage_gap"] = (underloads["_tgt"] - underloads["_pl"]).sum()

    mach_s = _col(d, col_map, "machine")
    if mach_s is not None:
        d["_m"] = mach_s.astype(str)
        bm = (d.groupby("_m")
              .agg(avg_payload=("_pl", "mean"), loads=("_pl", "count"), total_t=("_pl", "sum"))
              .reset_index().rename(columns={"_m": "Machine"})
              .sort_values("avg_payload", ascending=False))
        result["by_machine"] = bm

    op_s = _col(d, col_map, "operator")
    if op_s is not None:
        d["_op"] = op_s.astype(str)
        bo = (d.groupby("_op")
              .agg(avg_payload=("_pl", "mean"), loads=("_pl", "count"))
              .reset_index().rename(columns={"_op": "Operator"})
              .sort_values("avg_payload", ascending=False))
        result["by_operator"] = bo

    result["distribution"] = d["_pl"].values
    return result


# ---------------------------------------------------------------------------
# Utilization Analysis
# ---------------------------------------------------------------------------
def analyze_utilization(df: pd.DataFrame, col_map: dict) -> dict:
    result: dict = {"available": False}

    avail_s = _col(df, col_map, "availability_pct")
    util_s  = _col(df, col_map, "utilization_pct")
    op_s    = _col(df, col_map, "operating_hours")

    if avail_s is None and util_s is None and op_s is None:
        return result

    d = df.copy()
    if avail_s is not None:
        d["_av"] = _num(avail_s)
        if d["_av"].max() <= 1.0:
            d["_av"] *= 100
    if util_s is not None:
        d["_ut"] = _num(util_s)
        if d["_ut"].max() <= 1.0:
            d["_ut"] *= 100
    if op_s is not None:
        d["_op"] = _num(op_s)

    idle_s  = _col(df, col_map, "idle_hours")
    down_s  = _col(df, col_map, "down_hours")
    sched_s = _col(df, col_map, "scheduled_hours")
    if idle_s  is not None: d["_id"] = _num(idle_s)
    if down_s  is not None: d["_dn"] = _num(down_s)
    if sched_s is not None: d["_sc"] = _num(sched_s)

    result["available"] = True
    if "_av" in d: result["avg_availability"] = d["_av"].mean()
    if "_ut" in d: result["avg_utilization"]  = d["_ut"].mean()
    if "_op" in d:
        result["total_operating_hours"] = d["_op"].sum()
        result["avg_operating_hours"]   = d["_op"].mean()

    mach_s = _col(d, col_map, "machine")
    if mach_s is not None:
        d["_m"] = mach_s.astype(str)
        agg: dict = {}
        if "_av" in d.columns: agg["Availability %"]  = ("_av", "mean")
        if "_ut" in d.columns: agg["Utilization %"]   = ("_ut", "mean")
        if "_op" in d.columns: agg["Operating Hrs"]   = ("_op", "sum")
        if "_id" in d.columns: agg["Idle Hrs"]        = ("_id", "sum")
        if "_dn" in d.columns: agg["Down Hrs"]        = ("_dn", "sum")
        if agg:
            bm = d.groupby("_m").agg(**agg).reset_index().rename(columns={"_m": "Machine"})
            result["by_machine"] = bm
            if "Availability %" in bm.columns:
                result["low_availability"] = bm[bm["Availability %"] < 80]["Machine"].tolist()
            if "Utilization %" in bm.columns:
                result["low_utilization"] = bm[bm["Utilization %"] < 60]["Machine"].tolist()

    return result


# ---------------------------------------------------------------------------
# Operator Analysis
# ---------------------------------------------------------------------------
def analyze_operators(df: pd.DataFrame, col_map: dict) -> dict:
    result: dict = {"available": False}
    op_s = _col(df, col_map, "operator")
    if op_s is None:
        return result

    d = df.copy()
    d["_op"] = op_s.astype(str)
    agg: dict = {"Cycles": ("_op", "count")}

    pay_s = _col(d, col_map, "payload")
    if pay_s is not None:
        d["_pl"] = _num(pay_s)
        agg["Avg Payload (t)"] = ("_pl", "mean")

    ct_s = _col(d, col_map, "total_cycle_time")
    if ct_s is not None:
        d["_ct"] = _num(ct_s)
        agg["Avg Cycle (min)"] = ("_ct", "mean")

    by_op = (d.groupby("_op").agg(**agg).reset_index()
              .rename(columns={"_op": "Operator"})
              .sort_values("Cycles", ascending=False))

    # Efficiency score: composite of payload (high=good) + cycle time (low=good)
    scores = pd.Series(np.zeros(len(by_op)))
    n = 0
    for col, ascending in [("Avg Payload (t)", False), ("Avg Cycle (min)", True)]:
        if col in by_op.columns:
            vals = by_op[col].fillna(by_op[col].median())
            mn, mx = vals.min(), vals.max()
            norm = (vals - mn) / (mx - mn + 1e-9)
            scores += (1 - norm) if ascending else norm
            n += 1
    if n > 0:
        by_op["Efficiency Score"] = (scores / n * 100).round(1)

    result.update({
        "available":      True,
        "operator_count": by_op["Operator"].nunique(),
        "by_operator":    by_op,
        "top_performers": by_op.nlargest(3, "Cycles")["Operator"].tolist(),
    })
    if "Efficiency Score" in by_op.columns and len(by_op) >= 3:
        result["needs_coaching"] = by_op.nsmallest(3, "Efficiency Score")["Operator"].tolist()

    return result


# ---------------------------------------------------------------------------
# Run all
# ---------------------------------------------------------------------------
def run_all_analyses(df: pd.DataFrame, col_map: dict) -> dict:
    return {
        "cycle_times":  analyze_cycle_times(df, col_map),
        "payload":      analyze_payload(df, col_map),
        "utilization":  analyze_utilization(df, col_map),
        "operators":    analyze_operators(df, col_map),
        "row_count":    len(df),
    }
