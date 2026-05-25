# =============================================================================
# app.py – MineStar Insights: upload data → analyse → download PowerPoint
# Run with: streamlit run app.py
# =============================================================================
from __future__ import annotations
import logging
from datetime import datetime

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from data_processor import (load_data, get_sheet_names, preview_sheets,
                             load_multiple_sheets, detect_columns,
                             run_all_analyses, COLUMN_PATTERNS)
from db_connector import (build_conn_str, test_connection, get_tables,
                           get_date_columns, preview_table, score_table,
                           query_table, save_settings, load_settings,
                           get_available_driver)
from insights_engine import generate_all_insights
from presentation_builder import build_presentation
from sample_data import generate_sample_csv

logging.basicConfig(level=logging.WARNING)

st.set_page_config(
    page_title="MineStar Insights",
    page_icon="⛏️",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
  .main .block-container { padding-top: 1.2rem; }
  .severity-critical { color: #c62828; font-weight: 700; }
  .severity-warning  { color: #e65c00; font-weight: 600; }
  .severity-info     { color: #1a3a5c; }
  .severity-positive { color: #2e7d32; font-weight: 600; }
  .kpi-card { background:#1e2130; border-radius:8px; padding:14px 10px;
               text-align:center; border-top: 3px solid #b87333; }
</style>
""", unsafe_allow_html=True)

SEV_ICON  = {"critical": "🔴", "warning": "🟡", "info": "🔵", "positive": "🟢"}
SEV_COLOR = {"critical": "#c62828", "warning": "#e65c00",
             "info": "#1a3a5c", "positive": "#2e7d32"}


# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------
def render_sidebar(df: pd.DataFrame, col_map: dict) -> dict:
    st.sidebar.title("⛏️ MineStar Insights")
    st.sidebar.caption("Pinto Valley Mine — Operations Analytics")
    st.sidebar.divider()

    st.sidebar.header("🏭 Mine Settings")
    mine_name = st.sidebar.text_input("Mine / Site Name", value="Pinto Valley Mine")

    st.sidebar.header("🔧 Column Mapping")
    st.sidebar.caption(
        "Auto-detected columns shown below. Override if your export uses different names."
    )

    all_cols = ["(not available)"] + list(df.columns)
    updated_map = {}
    key_fields = [
        ("machine",          "Machine / Truck ID"),
        ("operator",         "Operator"),
        ("shift",            "Shift"),
        ("total_cycle_time", "Total Cycle Time (min)"),
        ("queue_time",       "Queue Time — Shovel (min)"),
        ("queue_sink",       "Queue Time — Crusher (min)"),
        ("load_time",        "Load Time (min)"),
        ("travel_loaded",    "Travel Loaded (min)"),
        ("travel_empty",     "Travel Empty (min)"),
        ("dump_time",        "Dump Time (min)"),
        ("payload",          "Payload (t)"),
        ("target_payload",   "Target Payload (t)"),
        ("operating_hours",  "Operating Hours"),
        ("idle_hours",       "Idle Hours"),
        ("down_hours",       "Down Hours"),
        ("availability_pct", "Physical Availability %"),
        ("utilization_pct",  "Utilization %"),
    ]
    for key, label in key_fields:
        detected = col_map.get(key, "(not available)")
        idx = all_cols.index(detected) if detected in all_cols else 0
        chosen = st.sidebar.selectbox(label, all_cols, index=idx, key=f"col_{key}")
        if chosen != "(not available)":
            updated_map[key] = chosen

    return {"mine_name": mine_name, "col_map": updated_map}


# ---------------------------------------------------------------------------
# Tab renderers
# ---------------------------------------------------------------------------
def render_summary_tab(analysis: dict, insights: list):
    ct = analysis.get("cycle_times", {})
    pl = analysis.get("payload", {})
    ut = analysis.get("utilization", {})
    op = analysis.get("operators", {})

    critical = sum(1 for i in insights if i["severity"] == "critical")
    warnings = sum(1 for i in insights if i["severity"] == "warning")

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("Records",       f"{analysis.get('row_count', 0):,}")
    c2.metric("Avg Cycle",     f"{ct['avg_cycle']:.1f} min"  if ct.get("available") else "N/A")
    c3.metric("Avg Payload",   f"{pl['avg_payload']:.1f} t"  if pl.get("available") else "N/A")
    c4.metric("Fleet Avail.",  f"{ut['avg_availability']:.0f}%" if ut.get("avg_availability") else "N/A")
    c5.metric("Critical Finds", critical, delta=None)
    c6.metric("Warnings",      warnings,  delta=None)

    st.divider()
    st.subheader("All Findings")

    for ins in insights:
        icon  = SEV_ICON.get(ins["severity"], "ℹ️")
        color = SEV_COLOR.get(ins["severity"], "#333")
        with st.expander(f"{icon} **[{ins['category']}]** {ins['finding'][:100]}…"):
            st.markdown(f"**Finding:** {ins['finding']}")
            st.markdown(f"**Recommended Action:**")
            st.info(ins["recommendation"])


def render_cycle_tab(ct: dict):
    if not ct.get("available"):
        st.info("No cycle time data detected. Ensure your file has a 'Total Cycle Time' column.")
        return

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Cycles",   f"{ct['total_cycles']:,}")
    c2.metric("Avg Cycle Time", f"{ct['avg_cycle']:.1f} min")
    c3.metric("Median",         f"{ct['median_cycle']:.1f} min")
    c4.metric("Variability (σ)", f"{ct['std_cycle']:.1f} min")

    col_a, col_b = st.columns(2)

    bm = ct.get("by_machine")
    if bm is not None and len(bm) > 0:
        with col_a:
            st.markdown("#### Avg Cycle Time by Machine")
            df_plot = bm.head(15).sort_values("avg_cycle")
            fig = px.bar(df_plot, x="avg_cycle", y="Machine", orientation="h",
                         color="avg_cycle",
                         color_continuous_scale=["#2e7d32", "#e65c00", "#c62828"],
                         labels={"avg_cycle": "Avg Cycle (min)"})
            fig.update_layout(template="plotly_dark", height=380, showlegend=False,
                              margin=dict(l=10, r=10, t=10, b=10),
                              coloraxis_showscale=False)
            st.plotly_chart(fig, use_container_width=True)

    bd = ct.get("time_breakdown", {})
    if bd:
        labels_map = {
            "queue_time":    "Queue (Shovel)",
            "queue_sink":    "Queue (Crusher)",
            "load_time":     "Load",
            "travel_loaded": "Travel (Loaded)",
            "travel_empty":  "Travel (Empty)",
            "dump_time":     "Dump",
            "spot_time":     "Spot",
        }
        labels = [labels_map.get(k, k) for k, v in bd.items() if v and v > 0]
        values = [v for v in bd.values() if v and v > 0]
        with col_b:
            st.markdown("#### Cycle Time Breakdown")
            fig2 = px.pie(names=labels, values=values, hole=0.4,
                          color_discrete_sequence=["#c62828", "#1a3a5c", "#b87333",
                                                    "#2e7d32", "#e65c00", "#7b1fa2"])
            fig2.update_layout(template="plotly_dark", height=380,
                               margin=dict(l=10, r=10, t=10, b=10))
            st.plotly_chart(fig2, use_container_width=True)

    bs = ct.get("by_shift")
    if bs is not None:
        st.markdown("#### Avg Cycle Time by Shift")
        fig3 = px.bar(bs, x="Shift", y="avg_cycle", color="Shift",
                      labels={"avg_cycle": "Avg Cycle Time (min)"},
                      color_discrete_sequence=["#1a3a5c", "#b87333"])
        fig3.update_layout(template="plotly_dark", height=250, showlegend=False,
                           margin=dict(l=10, r=10, t=10, b=10))
        st.plotly_chart(fig3, use_container_width=True)


def render_payload_tab(pl: dict):
    if not pl.get("available"):
        st.info("No payload data detected. Ensure your file has a 'Payload (t)' column.")
        return

    tgt = pl.get("target_payload")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Avg Payload",   f"{pl['avg_payload']:.1f} t")
    c2.metric("Total Tonnes",  f"{pl['total_tonnes']:,.0f} t")
    c3.metric("Target Payload", f"{tgt:.1f} t" if tgt else "N/A")
    gap = pl.get("tonnage_gap", 0)
    c4.metric("Tonnage Gap (underloads)", f"{gap:,.0f} t", delta=f"-{gap:,.0f}" if gap > 0 else None,
              delta_color="inverse")

    if tgt:
        c5, c6, c7 = st.columns(3)
        c5.metric("Overloaded (>10%)",  f"{pl.get('pct_overloaded', 0):.0f}%")
        c6.metric("On Target",          f"{pl.get('pct_on_target', 0):.0f}%")
        c7.metric("Underloaded (<90%)", f"{pl.get('pct_underloaded', 0):.0f}%")

    col_a, col_b = st.columns(2)

    dist = pl.get("distribution")
    if dist is not None and len(dist) > 0:
        with col_a:
            st.markdown("#### Payload Distribution")
            fig = go.Figure()
            fig.add_trace(go.Histogram(x=dist, nbinsx=30,
                                       marker_color="#1a3a5c", name="Payload"))
            if tgt:
                fig.add_vline(x=tgt,        line_color="#2e7d32", line_dash="dash",
                              annotation_text=f"Target {tgt:.0f}t", annotation_position="top right")
                fig.add_vline(x=tgt * 0.90, line_color="#e65c00", line_dash="dot",
                              annotation_text="−10%")
                fig.add_vline(x=tgt * 1.10, line_color="#c62828", line_dash="dot",
                              annotation_text="+10%")
            fig.update_layout(template="plotly_dark", height=360,
                              xaxis_title="Payload (t)", yaxis_title="Count",
                              margin=dict(l=10, r=10, t=10, b=10))
            st.plotly_chart(fig, use_container_width=True)

    bm = pl.get("by_machine")
    if bm is not None and len(bm) > 0:
        with col_b:
            st.markdown("#### Avg Payload by Machine")
            df_plot = bm.head(14).sort_values("avg_payload")
            colors = ["#2e7d32" if (tgt and v >= tgt * 0.95) else "#e65c00"
                      for v in df_plot["avg_payload"]]
            fig2 = go.Figure(go.Bar(
                x=df_plot["avg_payload"], y=df_plot["Machine"].astype(str),
                orientation="h", marker_color=colors,
            ))
            if tgt:
                fig2.add_vline(x=tgt, line_color="#ffffff", line_dash="dash",
                               annotation_text=f"Target {tgt:.0f}t")
            fig2.update_layout(template="plotly_dark", height=360,
                               xaxis_title="Avg Payload (t)",
                               margin=dict(l=10, r=10, t=10, b=10))
            st.plotly_chart(fig2, use_container_width=True)


def render_util_tab(ut: dict):
    if not ut.get("available"):
        st.info("No utilization data detected. Ensure your file has availability/utilization columns.")
        return

    c1, c2, c3 = st.columns(3)
    if ut.get("avg_availability"):
        c1.metric("Avg Availability", f"{ut['avg_availability']:.1f}%",
                  delta=f"{ut['avg_availability']-85:.1f}% vs 85% target",
                  delta_color="normal")
    if ut.get("avg_utilization"):
        c2.metric("Avg Utilization", f"{ut['avg_utilization']:.1f}%",
                  delta=f"{ut['avg_utilization']-70:.1f}% vs 70% target",
                  delta_color="normal")
    if ut.get("total_operating_hours"):
        c3.metric("Total Operating Hours", f"{ut['total_operating_hours']:,.0f} h")

    bm = ut.get("by_machine")
    if bm is None or len(bm) == 0:
        return

    has_stacked = all(c in bm.columns for c in ["Operating Hrs", "Idle Hrs", "Down Hrs"])
    has_pct = "Availability %" in bm.columns or "Utilization %" in bm.columns

    col_a, col_b = st.columns(2)

    if has_stacked:
        with col_a:
            st.markdown("#### Hours by Machine")
            df_p = bm.sort_values("Operating Hrs", ascending=False).head(14)
            fig = go.Figure()
            fig.add_trace(go.Bar(name="Operating", x=df_p["Machine"].astype(str),
                                  y=df_p["Operating Hrs"], marker_color="#2e7d32"))
            if "Idle Hrs" in df_p:
                fig.add_trace(go.Bar(name="Idle", x=df_p["Machine"].astype(str),
                                      y=df_p["Idle Hrs"], marker_color="#e65c00"))
            if "Down Hrs" in df_p:
                fig.add_trace(go.Bar(name="Down", x=df_p["Machine"].astype(str),
                                      y=df_p["Down Hrs"], marker_color="#c62828"))
            fig.update_layout(barmode="stack", template="plotly_dark", height=360,
                              margin=dict(l=10, r=10, t=10, b=10))
            st.plotly_chart(fig, use_container_width=True)

    if has_pct:
        col_use = col_b if has_stacked else col_a
        pct_col = "Availability %" if "Availability %" in bm.columns else "Utilization %"
        target_line = 85.0 if "Availability" in pct_col else 70.0
        with col_use:
            st.markdown(f"#### {pct_col} by Machine")
            df_p = bm.sort_values(pct_col).head(14)
            colors = ["#c62828" if v < target_line - 10 else
                      "#e65c00" if v < target_line else "#2e7d32"
                      for v in df_p[pct_col]]
            fig2 = go.Figure(go.Bar(
                x=df_p[pct_col], y=df_p["Machine"].astype(str),
                orientation="h", marker_color=colors,
            ))
            fig2.add_vline(x=target_line, line_color="#ffffff", line_dash="dash",
                           annotation_text=f"Target {target_line:.0f}%")
            fig2.update_layout(template="plotly_dark", height=360,
                               xaxis_range=[0, 110],
                               margin=dict(l=10, r=10, t=10, b=10))
            st.plotly_chart(fig2, use_container_width=True)

    st.markdown("#### Machine Utilization Table")
    st.dataframe(bm.round(1), use_container_width=True, height=300)


def render_operator_tab(op: dict):
    if not op.get("available"):
        st.info("No operator data detected. Ensure your file has an 'Operator' column.")
        return

    by_op = op.get("by_operator", pd.DataFrame())
    c1, c2 = st.columns(2)
    c1.metric("Total Operators", op.get("operator_count", 0))
    top = op.get("top_performers", [])
    c2.metric("Top Performer", str(top[0]) if top else "N/A")

    col_a, col_b = st.columns(2)

    if "Cycles" in by_op.columns:
        with col_a:
            st.markdown("#### Cycles Completed by Operator")
            df_p = by_op.sort_values("Cycles").tail(15)
            colors = ["#b87333" if op in top else "#1a3a5c"
                      for op in df_p["Operator"].astype(str)]
            fig = go.Figure(go.Bar(
                x=df_p["Cycles"], y=df_p["Operator"].astype(str),
                orientation="h", marker_color=colors,
            ))
            fig.update_layout(template="plotly_dark", height=380,
                              margin=dict(l=10, r=10, t=10, b=10))
            st.plotly_chart(fig, use_container_width=True)

    if "Avg Payload (t)" in by_op.columns:
        with col_b:
            st.markdown("#### Avg Payload by Operator")
            df_p2 = by_op.sort_values("Avg Payload (t)").tail(15)
            fig2 = px.bar(df_p2, x="Avg Payload (t)", y="Operator",
                          orientation="h", color="Avg Payload (t)",
                          color_continuous_scale=["#e65c00", "#2e7d32"])
            fig2.update_layout(template="plotly_dark", height=380, showlegend=False,
                               coloraxis_showscale=False,
                               margin=dict(l=10, r=10, t=10, b=10))
            st.plotly_chart(fig2, use_container_width=True)

    if "Efficiency Score" in by_op.columns:
        st.markdown("#### Operator Efficiency Scoreboard")
        disp = by_op.sort_values("Efficiency Score", ascending=False).reset_index(drop=True)
        disp.index += 1
        st.dataframe(disp.round(1), use_container_width=True, height=320)

    by_crew = op.get("by_crew")
    if by_crew is not None and len(by_crew) > 0:
        st.markdown("#### Performance by Crew")
        st.dataframe(by_crew.round(1), use_container_width=True, hide_index=True)


# ---------------------------------------------------------------------------
# Landing page
# ---------------------------------------------------------------------------
def render_landing():
    st.markdown("""
    ## Welcome to MineStar Insights ⛏️
    Upload a MineStar export (CSV or Excel) to automatically generate:
    - **Fleet cycle time analysis** — avg cycle, queue time, shift comparison
    - **Payload performance** — overloads, underloads, tonnage gap vs target
    - **Equipment utilization** — availability %, idle hours, downtime by machine
    - **Operator performance** — cycles, avg payload, efficiency scoring
    - **PowerPoint report** — professional slide deck with findings + recommendations

    ### How to export data from MineStar
    1. Open MineStar Fleet Management → Reports
    2. Select **Cycle Report**, **Shift Report**, or **Equipment Report**
    3. Set your date range and export as **CSV** or **Excel**
    4. Upload the file above

    ### Don't have data handy? Try the sample dataset below.
    """)

    if st.button("📥 Download sample MineStar data", key="sample_btn"):
        csv_bytes = generate_sample_csv()
        st.download_button(
            "⬇️ Save sample_minestar.csv",
            data=csv_bytes,
            file_name="sample_minestar.csv",
            mime="text/csv",
        )


# ---------------------------------------------------------------------------
# SQL Server connection UI
# ---------------------------------------------------------------------------

def render_db_section() -> Optional[pd.DataFrame]:
    """Direct SQL Server connection. Returns loaded DataFrame or None."""
    saved = load_settings()

    # ── Step 1: Connection form ───────────────────────────────────────────────
    st.markdown("#### Step 1 — Connection Details")

    c1, c2 = st.columns(2)
    server   = c1.text_input("Server", value=saved.get("server", ""),
                              placeholder=r"PVMINE-BACKUP\MINESTAR", key="db_server")
    database = c2.text_input("Database", value=saved.get("database", ""),
                              placeholder="MineStar_Prod", key="db_database")

    auth_choice = st.radio(
        "Authentication",
        ["Windows Authentication (Trusted)", "SQL Server Login"],
        index=0 if saved.get("use_windows_auth", True) else 1,
        horizontal=True, key="db_auth",
    )
    use_windows = auth_choice.startswith("Windows")

    username = password = ""
    if not use_windows:
        u1, u2 = st.columns(2)
        username = u1.text_input("Username", value=saved.get("username", ""), key="db_user")
        password = u2.text_input("Password", type="password", key="db_pass")

    driver = get_available_driver()
    if driver is None:
        st.warning(
            "No SQL Server ODBC driver detected on this machine. "
            "Download **ODBC Driver 17 for SQL Server** from Microsoft, install it, "
            "then restart this app."
        )

    if st.button("🔌 Test & Connect", type="primary", key="db_connect_btn"):
        if not server or not database:
            st.error("Enter both server and database name.")
        else:
            conn_str = build_conn_str(server, database, use_windows, username, password, driver)
            with st.spinner("Connecting to SQL Server…"):
                ok, msg = test_connection(conn_str)
            if ok:
                st.session_state["db_conn_str"]  = conn_str
                st.session_state["db_connected"] = True
                st.session_state.pop("db_df", None)   # clear any previous data
                save_settings(server, database, use_windows, username)
                st.success(f"✅ {msg}")
                st.rerun()
            else:
                st.error(f"❌ Connection failed: {msg}")

    if not st.session_state.get("db_connected"):
        return None

    conn_str = st.session_state["db_conn_str"]

    # ── Step 2: Table browser ─────────────────────────────────────────────────
    st.divider()
    st.markdown("#### Step 2 — Browse & Select Table")

    if "db_tables" not in st.session_state:
        with st.spinner("Loading table list…"):
            st.session_state["db_tables"] = get_tables(conn_str)

    all_tables = st.session_state["db_tables"]
    if not all_tables:
        st.warning("No tables found in this database. Check permissions.")
        return None

    search = st.text_input("🔍 Filter tables", placeholder="type to search…", key="db_search")
    filtered = [t for t in all_tables if search.lower() in t.lower()] if search else all_tables

    default_idx = 0
    if saved.get("last_table") in filtered:
        default_idx = filtered.index(saved["last_table"])

    selected_table = st.selectbox(
        f"Select table ({len(filtered)} shown)", filtered,
        index=default_idx, key="db_table_select",
    )

    # Show table preview + detected data types
    if selected_table:
        with st.expander("Preview table (first 5 rows)", expanded=False):
            try:
                prev = preview_table(conn_str, selected_table, n=5)
                st.dataframe(prev, use_container_width=True)
                info = score_table(conn_str, selected_table)
                if info["detected_types"]:
                    st.caption(
                        f"Detected MineStar data: {' · '.join(info['detected_types'])}  "
                        f"| Estimated rows: {info['row_count']:,}"
                    )
                else:
                    st.caption("No MineStar column patterns detected in this table.")
            except Exception as e:
                st.error(f"Could not preview: {e}")

    # ── Step 3: Date range & row limit ────────────────────────────────────────
    st.divider()
    st.markdown("#### Step 3 — Date Range & Load")

    date_cols = []
    if selected_table:
        try:
            date_cols = get_date_columns(conn_str, selected_table)
        except Exception:
            pass

    use_date_filter = st.toggle("Filter by date range", value=bool(date_cols), key="db_use_date")

    date_col = start_dt = end_dt = None
    if use_date_filter and date_cols:
        dc1, dc2, dc3 = st.columns(3)
        date_col = dc1.selectbox("Date column", date_cols, key="db_date_col")
        start_dt = dc2.date_input("From", value=pd.Timestamp.now() - pd.Timedelta(days=30),
                                   key="db_start")
        end_dt   = dc3.date_input("To",   value=pd.Timestamp.now(), key="db_end")
    elif use_date_filter and not date_cols:
        st.info("No date columns detected in this table.")

    row_limit = st.select_slider(
        "Max rows to load",
        options=[1_000, 5_000, 10_000, 25_000, 50_000, 100_000],
        value=25_000, key="db_row_limit",
    )

    if st.button("⬇️ Load Data from Database", type="primary", key="db_load_btn"):
        with st.spinner(f"Querying {selected_table}…"):
            try:
                df = query_table(
                    conn_str, selected_table,
                    date_col=date_col if use_date_filter else None,
                    start_date=start_dt, end_date=end_dt,
                    limit=row_limit,
                )
                df.columns = df.columns.str.strip()
                st.session_state["db_df"] = df
                save_settings(
                    st.session_state.get("db_server", ""),
                    st.session_state.get("db_database", ""),
                    use_windows, username, last_table=selected_table,
                )
                st.success(f"Loaded **{len(df):,} rows** from `{selected_table}`.")
                st.rerun()
            except Exception as e:
                st.error(f"Query failed: {e}")

    return st.session_state.get("db_df")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    st.title("⛏️ MineStar Operations Insights")
    st.caption("Connect to your MineStar SQL Server or upload an exported file.")

    source = st.radio(
        "Data source",
        ["📁 Upload File (CSV / Excel / Power BI export)",
         "🗄️ SQL Server — Direct Database Connection"],
        horizontal=True, key="data_source",
    )

    df: Optional[pd.DataFrame] = None

    # ── SQL Server path ───────────────────────────────────────────────────────
    if source.startswith("🗄️"):
        df = render_db_section()
        if df is None:
            return

    # ── File upload path ──────────────────────────────────────────────────────
    else:
        uploaded = st.file_uploader(
            "Upload MineStar export (CSV or Excel)",
            type=["csv", "xlsx", "xls"],
            label_visibility="collapsed",
        )

        if uploaded is None:
            render_landing()
            return

    # ── Multi-sheet selector for Excel / Power BI exports ────────────────────
    sheet_names = get_sheet_names(uploaded)

    if not sheet_names:
        # Plain CSV — load directly
        try:
            df = load_data(uploaded)
        except ValueError as e:
            st.error(str(e))
            return

    elif len(sheet_names) == 1:
        # Single-sheet Excel — load directly
        try:
            df = load_data(uploaded, sheet_name=sheet_names[0])
        except ValueError as e:
            st.error(str(e))
            return

    else:
        # Multi-sheet Excel — show preview table and let user pick sheets
        with st.spinner("Scanning sheets…"):
            previews = preview_sheets(uploaded)

        type_icons = {
            "cycle_times": "🔄", "payload": "⚖️",
            "utilization": "🔧", "operators": "👷",
        }
        type_labels = {
            "cycle_times": "Cycle Times", "payload": "Payload",
            "utilization": "Utilization",  "operators": "Operators",
        }

        st.markdown("### Select sheets to include in the report")
        st.caption(
            "The app detected which sheets contain MineStar data. "
            "Tick the ones you want — they will be merged into one combined analysis."
        )

        # Build a preview table
        preview_rows = []
        for sname, info in previews.items():
            types_str = "  ".join(
                f"{type_icons[t]} {type_labels[t]}"
                for t in info["detected_types"]
            ) or "⚠️ No MineStar data detected"
            preview_rows.append({
                "Sheet": sname,
                "Rows": f"{info['rows']:,}",
                "Columns": info["cols"],
                "Contains": types_str,
            })

        st.dataframe(pd.DataFrame(preview_rows), use_container_width=True,
                     hide_index=True, height=min(200, 45 + len(preview_rows) * 38))

        # Default: select sheets that have MineStar data
        default_selected = [s for s, info in previews.items() if info["score"] > 0]
        if not default_selected:
            default_selected = sheet_names[:1]

        selected_sheets = st.multiselect(
            "Sheets to merge:",
            options=sheet_names,
            default=default_selected,
            key="sheet_multiselect",
        )

        if not selected_sheets:
            st.warning("Select at least one sheet to continue.")
            st.stop()

        # Load and merge
        try:
            with st.spinner(f"Loading {len(selected_sheets)} sheet(s)…"):
                df = load_multiple_sheets(uploaded, selected_sheets)
            if len(selected_sheets) > 1:
                st.success(
                    f"Merged **{len(selected_sheets)} sheets** → "
                    f"**{len(df):,} rows** total."
                )
        except ValueError as e:
            st.error(str(e))
            return
    # ── end of data source section — df is now populated either way ──────────

    col_map_detected = detect_columns(df)
    n_detected = len(col_map_detected)
    st.success(
        f"Loaded **{len(df):,} rows × {len(df.columns)} columns**. "
        f"Auto-detected **{n_detected}** MineStar fields."
        + (f"  ·  Filtering to TruckCycle/LoaderCycle rows only."
           if col_map_detected.get("cycle_type") else "")
    )

    # Sidebar settings + column override
    settings = render_sidebar(df, col_map_detected)
    col_map  = settings["col_map"]
    mine_name = settings["mine_name"]

    if n_detected == 0:
        st.warning(
            "No MineStar columns were auto-detected. Please map your columns in the sidebar."
        )

    # Run analysis
    with st.spinner("Analysing data…"):
        analysis = run_all_analyses(df, col_map)
        insights = generate_all_insights(analysis)

    available = [k for k in ["cycle_times", "payload", "utilization", "operators"]
                 if analysis[k].get("available")]

    if not available:
        st.error(
            "No analysable columns found. Please map your columns using the sidebar "
            "and ensure the data contains numeric values."
        )
        with st.expander("Raw data preview"):
            st.dataframe(df.head(20), use_container_width=True)
        return

    # Tabs
    tab_labels = ["📊 Summary"]
    if analysis["cycle_times"].get("available"):  tab_labels.append("🔄 Cycle Times")
    if analysis["payload"].get("available"):       tab_labels.append("⚖️ Payload")
    if analysis["utilization"].get("available"):   tab_labels.append("🔧 Utilization")
    if analysis["operators"].get("available"):     tab_labels.append("👷 Operators")

    tabs = st.tabs(tab_labels)
    idx = 0
    with tabs[idx]:
        render_summary_tab(analysis, insights)
    idx += 1
    if analysis["cycle_times"].get("available"):
        with tabs[idx]: render_cycle_tab(analysis["cycle_times"])
        idx += 1
    if analysis["payload"].get("available"):
        with tabs[idx]: render_payload_tab(analysis["payload"])
        idx += 1
    if analysis["utilization"].get("available"):
        with tabs[idx]: render_util_tab(analysis["utilization"])
        idx += 1
    if analysis["operators"].get("available"):
        with tabs[idx]: render_operator_tab(analysis["operators"])

    # Generate PowerPoint
    st.divider()
    col_gen, col_info = st.columns([1, 3])
    with col_gen:
        gen_btn = st.button("📊 Generate PowerPoint", type="primary",
                            use_container_width=True)
    with col_info:
        st.caption(
            f"Will generate a {len(available)+3}-slide deck: title, executive summary, "
            f"{', '.join(available)}, and recommendations."
        )

    if gen_btn:
        with st.spinner("Building presentation…"):
            pptx_bytes = build_presentation(analysis, insights, mine_name)
        fname = f"minestar_insights_{datetime.now().strftime('%Y%m%d_%H%M')}.pptx"
        st.download_button(
            "⬇️ Download PowerPoint",
            data=pptx_bytes,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
        st.success(
            f"Presentation ready — {len(insights)} findings across "
            f"{len(available)} analysis categories."
        )


if __name__ == "__main__":
    main()
