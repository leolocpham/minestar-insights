# =============================================================================
# insights_engine.py – Generate findings and recommendations from analysis
# =============================================================================
from __future__ import annotations
from typing import List, Dict


def generate_all_insights(analysis: dict) -> List[Dict]:
    insights: list = []
    ct = analysis.get("cycle_times", {})
    pl = analysis.get("payload", {})
    ut = analysis.get("utilization", {})
    op = analysis.get("operators", {})

    if ct.get("available"): insights.extend(_cycle_insights(ct))
    if pl.get("available"): insights.extend(_payload_insights(pl))
    if ut.get("available"): insights.extend(_util_insights(ut))
    if op.get("available"): insights.extend(_operator_insights(op))

    order = {"critical": 0, "warning": 1, "info": 2, "positive": 3}
    insights.sort(key=lambda x: order.get(x["severity"], 9))
    return insights


# ---------------------------------------------------------------------------
def _cycle_insights(ct: dict) -> list:
    out = []
    avg = ct.get("avg_cycle", 0)
    std = ct.get("std_cycle", 0)

    if avg > 0 and std / avg > 0.30:
        out.append({
            "category": "Cycle Times",
            "finding": (f"High cycle-time variability detected — avg {avg:.1f} min, "
                        f"std deviation {std:.1f} min (CV = {std/avg*100:.0f}%)."),
            "severity": "warning",
            "recommendation": (
                "Standardise haul route selection. Conduct time-motion studies on "
                "outlier trucks. Check for inconsistent dispatch instructions causing "
                "operators to choose sub-optimal routes."
            ),
        })

    bd = ct.get("time_breakdown", {})
    queue = bd.get("queue_time", 0)
    if avg > 0 and queue and queue / avg > 0.18:
        out.append({
            "category": "Cycle Times",
            "finding": (f"Queue time averages {queue:.1f} min, representing "
                        f"{queue/avg*100:.0f}% of total cycle time. "
                        "Excessive queuing indicates a shovel or crusher bottleneck."),
            "severity": "critical",
            "recommendation": (
                "Review dispatch rules to prevent truck bunching. Implement dynamic "
                "truck allocation to balance shovel queues. Target queue time < 10% "
                "of total cycle. Consider staggered truck release intervals."
            ),
        })

    if avg > 45:
        out.append({
            "category": "Cycle Times",
            "finding": (f"Fleet average cycle time of {avg:.1f} min is above the "
                        "45-min benchmark, directly reducing tonnes moved per shift."),
            "severity": "warning",
            "recommendation": (
                "Audit haul road conditions for speed restrictions, grade, and surface. "
                "Review dump point congestion and consider additional dump positions. "
                "Ensure speed limits are set correctly in dispatch system."
            ),
        })

    bm = ct.get("by_machine")
    if bm is not None and len(bm) > 1:
        worst = bm.iloc[0]
        best  = bm["avg_cycle"].min()
        if worst["avg_cycle"] > best * 1.35:
            out.append({
                "category": "Cycle Times",
                "finding": (f"Truck {worst['Machine']} has the longest avg cycle "
                            f"({worst['avg_cycle']:.1f} min), "
                            f"{(worst['avg_cycle']/best - 1)*100:.0f}% above fleet best."),
                "severity": "warning",
                "recommendation": (
                    f"Investigate truck {worst['Machine']} for mechanical issues "
                    "(speed limiter, tyre condition, retarder), operator habits, "
                    "and route assignment. Consider reassigning to a shorter haul."
                ),
            })

    return out


# ---------------------------------------------------------------------------
def _payload_insights(pl: dict) -> list:
    out = []
    avg    = pl.get("avg_payload", 0)
    target = pl.get("target_payload")

    if target and avg > 0:
        ratio = avg / target * 100
        gap   = pl.get("tonnage_gap", 0)
        if ratio < 90:
            out.append({
                "category": "Payload",
                "finding": (f"Fleet avg payload ({avg:.1f} t) is {100-ratio:.0f}% below "
                            f"target ({target:.1f} t). Estimated lost production: "
                            f"{gap:,.0f} t in this dataset."),
                "severity": "critical",
                "recommendation": (
                    "Conduct immediate operator loading training. Verify bucket/dipper "
                    "sizing is appropriate for the material. Ensure payload meters are "
                    "calibrated. Review bench preparation and fragmentation quality."
                ),
            })
        elif ratio > 112:
            out.append({
                "category": "Payload",
                "finding": (f"Fleet avg payload ({avg:.1f} t) exceeds target by "
                            f"{ratio-100:.0f}%. Systematic overloading accelerates "
                            "tyre, suspension, and frame wear."),
                "severity": "warning",
                "recommendation": (
                    "Counsel operators on overloading costs and risks. Calibrate shovel "
                    "payload meters. Review dig face bench preparation. Implement "
                    "load-and-haul audits with supervisors present."
                ),
            })

    pct_over  = pl.get("pct_overloaded", 0)
    pct_under = pl.get("pct_underloaded", 0)

    if pct_over > 15:
        out.append({
            "category": "Payload",
            "finding": (f"{pct_over:.0f}% of loads are overloaded (>10% above target). "
                        "This is significantly above the acceptable threshold of 5%."),
            "severity": "critical",
            "recommendation": (
                "Implement immediate operator coaching programme. Calibrate all shovel "
                "payload meters. Consider auto-payload alerts in MineStar for loads "
                "exceeding target by >10%. Target overloads < 5%."
            ),
        })

    if pct_under > 25:
        out.append({
            "category": "Payload",
            "finding": (f"{pct_under:.0f}% of loads are underloaded (<90% of target). "
                        "Underloading wastes truck capacity and reduces shift production."),
            "severity": "warning",
            "recommendation": (
                "Investigate root causes: poor bench preparation, excessive fragmentation, "
                "or conservative operator loading. Adjust pass strategy. Run loading "
                "efficiency workshops with shovel and truck operators."
            ),
        })

    if pct_over < 8 and pct_under < 15 and avg > 0:
        out.append({
            "category": "Payload",
            "finding": (f"Payload management is well-controlled — "
                        f"{pl.get('pct_on_target', 0):.0f}% of loads within target range. "
                        f"Fleet avg: {avg:.1f} t."),
            "severity": "positive",
            "recommendation": (
                "Maintain current loading practices. Continue monitoring for drift. "
                "Share loading best practices across all operators."
            ),
        })

    return out


# ---------------------------------------------------------------------------
def _util_insights(ut: dict) -> list:
    out = []
    avail = ut.get("avg_availability")
    util  = ut.get("avg_utilization")

    if avail is not None:
        if avail < 75:
            out.append({
                "category": "Utilization",
                "finding": (f"Fleet physical availability is {avail:.1f}% — critically "
                            "below the 80% industry benchmark. Excessive unplanned downtime."),
                "severity": "critical",
                "recommendation": (
                    "Audit PM (preventive maintenance) scheduling compliance immediately. "
                    "Identify top 3 downtime codes and implement root cause corrective "
                    "actions. Review component life tracking to prevent surprise failures. "
                    "Escalate to maintenance management."
                ),
            })
        elif avail < 85:
            out.append({
                "category": "Utilization",
                "finding": (f"Fleet physical availability is {avail:.1f}% — below the "
                            "85% best-practice target. Improvement potential exists."),
                "severity": "warning",
                "recommendation": (
                    "Audit maintenance practices and parts inventory levels. Identify "
                    "chronic downtime offenders and implement targeted improvement plans. "
                    "Review MTTR (Mean Time to Repair) for top failure modes."
                ),
            })
        else:
            out.append({
                "category": "Utilization",
                "finding": (f"Fleet physical availability is strong at {avail:.1f}%, "
                            "meeting the 85% best-practice benchmark."),
                "severity": "positive",
                "recommendation": (
                    "Sustain current maintenance practices. Shift focus to utilization "
                    "improvement (reducing operator delays, fuel delays, blast waits)."
                ),
            })

    if util is not None and util < 65:
        out.append({
            "category": "Utilization",
            "finding": (f"Use-of-availability is {util:.1f}% — below the 70% target. "
                        "Significant idle/standby time is eroding production hours."),
            "severity": "warning",
            "recommendation": (
                "Classify delay categories (fuel, blast, congestion, operator). "
                "Implement delay reduction initiatives for the top 2–3 categories. "
                "Review shift handover procedures to reduce transition delays."
            ),
        })

    low_av = ut.get("low_availability", [])
    if low_av:
        machines = ", ".join(str(m) for m in low_av[:5])
        suffix = f" (+{len(low_av)-5} more)" if len(low_av) > 5 else ""
        out.append({
            "category": "Utilization",
            "finding": (f"{len(low_av)} unit(s) below 80% availability: "
                        f"{machines}{suffix}."),
            "severity": "warning",
            "recommendation": (
                "Focus maintenance attention on these specific units. Conduct root cause "
                "analysis on their individual top downtime codes. Consider temporary "
                "removal from production fleet if chronic."
            ),
        })

    return out


# ---------------------------------------------------------------------------
def _operator_insights(op: dict) -> list:
    out = []
    by_op = op.get("by_operator")
    if by_op is None or len(by_op) == 0:
        return out

    if "Cycles" in by_op.columns and len(by_op) > 1:
        mx = by_op["Cycles"].max()
        mn = by_op["Cycles"].min()
        if mx > mn * 1.5:
            out.append({
                "category": "Operators",
                "finding": (f"Large productivity spread across operators: top performer "
                            f"completed {mx} cycles vs {mn} for the lowest — a "
                            f"{mx/mn:.1f}x gap."),
                "severity": "warning",
                "recommendation": (
                    "Investigate whether the gap is driven by machine allocation, shift "
                    "assignment, or skill. Pair high-performing operators as mentors. "
                    "Conduct one-on-one cycle replay reviews in MineStar with lower performers."
                ),
            })

    top = op.get("top_performers", [])
    if top:
        out.append({
            "category": "Operators",
            "finding": f"Top 3 operators by cycle count: {', '.join(str(o) for o in top)}.",
            "severity": "positive",
            "recommendation": (
                "Formally recognise top performers. Document their practices — "
                "loading technique, route choices, pre-start habits — and incorporate "
                "into operator training materials."
            ),
        })

    coaching = op.get("needs_coaching", [])
    if coaching:
        out.append({
            "category": "Operators",
            "finding": (f"Operators with the lowest composite efficiency scores: "
                        f"{', '.join(str(o) for o in coaching)}."),
            "severity": "info",
            "recommendation": (
                "Schedule individualised coaching sessions. Review their MineStar cycle "
                "replays to identify specific improvement areas (loading dwell time, "
                "travel speed, payload consistency). Set 30-day improvement targets."
            ),
        })

    return out
