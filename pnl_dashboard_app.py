"""
Streamlit UI for the PnL Dashboard checker.
All comparison logic lives in compare_pnl_dashboard.py.
Run with: streamlit run pnl_dashboard_app.py
"""

import io
import os
import subprocess
import sys

import pandas as pd
import streamlit as st

from compare_pnl_dashboard import run_comparison, EXCEL_PATH, REPORT_PATH, SHEET_NAME


# ── HELPERS ───────────────────────────────────────────────────────────────────
def open_excel(path):
    if sys.platform == "win32":
        os.startfile(path)
    elif sys.platform == "darwin":
        subprocess.run(["open", path])
    else:
        subprocess.run(["xdg-open", path])


def build_report(issues_df) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        if issues_df.empty:
            pd.DataFrame({"Result": ["All cells matched their corresponding SQL rows."]}).to_excel(
                writer, sheet_name="Issues", index=False
            )
        else:
            issues_df.to_excel(writer, sheet_name="Issues", index=False)
    buf.seek(0)
    return buf.read()


# ── PAGE SETUP ────────────────────────────────────────────────────────────────
st.set_page_config(page_title="PnL Dashboard Checker", layout="wide")

st.markdown("""
<style>
    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

    .title-bar {
        display: flex; align-items: baseline; gap: 12px;
        border-bottom: 1px solid #e2e2e2; padding-bottom: 12px; margin-bottom: 28px;
    }
    .title-bar h1 { font-size: 1.3rem; font-weight: 600; margin: 0; color: #5ab4d6; }
    .title-bar span { font-size: 0.82rem; color: #888; }

    .stat-row { display: flex; gap: 16px; margin-bottom: 28px; }
    .stat-card {
        flex: 1; padding: 18px 20px; border-radius: 8px;
        background: #0d1b2a; border: 1px solid #1e3a5f;
    }
    .stat-card .label { font-size: 0.72rem; text-transform: uppercase;
                        letter-spacing: .08em; color: #7bafd4; margin-bottom: 4px; }
    .stat-card .value { font-size: 1.7rem; font-weight: 700; color: #5ab4d6; }
    .stat-card.ok   .value { color: #16a34a; }
    .stat-card.warn .value { color: #dc2626; }

    div[data-testid="stDataFrame"] { border: 1px solid #e8e8e8; border-radius: 8px; overflow: hidden; }
    .block-container { padding-top: 2rem !important; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="title-bar">
    <h1>PnL Dashboard Checker</h1>
    <span>Excel ↔ MySQL reconciliation</span>
</div>
""", unsafe_allow_html=True)


# ── ACTION BUTTONS ────────────────────────────────────────────────────────────
col_btn1, col_btn2, col_spacer = st.columns([1, 1, 5])

with col_btn1:
    run = st.button("▶  Run comparison", use_container_width=True, type="primary")

with col_btn2:
    if st.button("📂  Open Dashboard in Excel", use_container_width=True):
        try:
            open_excel(EXCEL_PATH)
        except Exception as e:
            st.error(f"Could not open file: {e}")


# ── RUN ───────────────────────────────────────────────────────────────────────
if run:
    with st.spinner("Connecting to MySQL and comparing…"):
        try:
            total, issues_df = run_comparison()
        except Exception as e:
            st.error(f"**Error:** {e}")
            st.stop()

    report_bytes = build_report(issues_df)

    with open(REPORT_PATH, "wb") as fh:
        fh.write(report_bytes)

    st.session_state["total"]        = total
    st.session_state["issues_df"]    = issues_df
    st.session_state["report_bytes"] = report_bytes
    st.session_state["report_ready"] = True


# ── RENDER RESULTS ────────────────────────────────────────────────────────────
if st.session_state.get("report_ready"):
    total        = st.session_state["total"]
    issues_df    = st.session_state["issues_df"]
    report_bytes = st.session_state["report_bytes"]
    matched      = total - len(issues_df)

    st.markdown(f"""
    <div class="stat-row">
        <div class="stat-card">
            <div class="label">Cells checked</div>
            <div class="value">{total:,}</div>
        </div>
        <div class="stat-card ok">
            <div class="label">Matched</div>
            <div class="value">{matched:,}</div>
        </div>
        <div class="stat-card {'warn' if len(issues_df) else 'ok'}">
            <div class="label">Issues</div>
            <div class="value">{len(issues_df):,}</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    tab_issues, = st.tabs(["Issues"])

    with tab_issues:
        if issues_df.empty:
            st.success("All cells match their corresponding SQL rows.")
        else:
            with st.expander("Filter results", expanded=False):
                fc1, fc2, fc3 = st.columns(3)
                status_opts = ["All"] + sorted(issues_df["Status"].unique().tolist())
                flag_opts   = ["All"] + sorted(issues_df["Flag"].unique().tolist())
                freq_opts   = ["All"] + sorted(issues_df["Frequency"].unique().tolist())
                sel_status  = fc1.selectbox("Status",    status_opts)
                sel_flag    = fc2.selectbox("Exchange",  flag_opts)
                sel_freq    = fc3.selectbox("Frequency", freq_opts)

            filtered = issues_df.copy()
            if sel_status != "All": filtered = filtered[filtered["Status"] == sel_status]
            if sel_flag   != "All": filtered = filtered[filtered["Flag"]   == sel_flag]
            if sel_freq   != "All": filtered = filtered[filtered["Frequency"] == sel_freq]

            st.markdown(f"**{len(filtered):,} issue(s)** shown")
            st.dataframe(filtered, use_container_width=True,
                         height=min(600, 55 + 35 * len(filtered)), hide_index=True)

    dl1, dl2 = st.columns([1, 1])
    with dl1:
        st.download_button(
            label="⬇  Download report (.xlsx)",
            data=report_bytes,
            file_name="tbl_pnl_dashboard_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with dl2:
        if st.button("📂  Open report in Excel", use_container_width=True):
            try:
                open_excel(REPORT_PATH)
            except Exception as e:
                st.error(f"Could not open report: {e}")
