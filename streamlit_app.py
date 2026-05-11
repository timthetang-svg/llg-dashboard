#!/usr/bin/env python3
"""
LLG Industries — Cloud Dashboard
Upload LLE / LLG sales Excel → download your dashboard HTML
"""
import io
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

sys.path.insert(0, str(Path(__file__).parent))
from dashboard import (
    load, load_foc, load_full,
    build_dashboard, build_agent_dashboard,
)

st.set_page_config(
    page_title="LLG Dashboard",
    page_icon="📊",
    layout="wide",
)

st.markdown(
    "<h1 style='color:#1a3c5e;font-family:Segoe UI'>📊 LLG Industries — Sales Dashboard</h1>",
    unsafe_allow_html=True,
)
st.markdown("Upload your sales Excel file(s), choose the report type, then click **Generate**.")

st.divider()

col1, col2 = st.columns(2)
with col1:
    lle_file = st.file_uploader("LLE Sales Data (.xlsx)", type="xlsx", key="lle")
with col2:
    llg_file = st.file_uploader("LLG Sales Data (.xlsx)", type="xlsx", key="llg")

report_type = st.radio(
    "Report type",
    ["Company Overview", "Agent Dashboard (AGENT1 & AGENT2)", "Agent 3 Dashboard"],
    horizontal=True,
)

st.divider()

if st.button("🚀 Generate Dashboard", type="primary", use_container_width=True):
    if not lle_file and not llg_file:
        st.error("Please upload at least one Excel file.")
        st.stop()

    with st.spinner("Reading data and building charts… (may take 20–40 seconds)"):
        try:
            frames, foc_frames, full_frames = [], [], []

            for uploaded, company in [(lle_file, "LLE"), (llg_file, "LLG")]:
                if uploaded is None:
                    continue
                raw = uploaded.read()
                frames.append(load(io.BytesIO(raw), company))
                foc_frames.append(load_foc(io.BytesIO(raw), company))
                full_frames.append(load_full(io.BytesIO(raw)))

            combined = pd.concat(frames, ignore_index=True)
            foc_df   = pd.concat(foc_frames, ignore_index=True) if foc_frames else pd.DataFrame(columns=["Cost_RM"])
            foc_cost = float(foc_df["Cost_RM"].sum()) if not foc_df.empty else 0.0
            df_full  = pd.concat(full_frames, ignore_index=True) if full_frames else pd.DataFrame(
                columns=["Year", "Customer", "Revenue_RM", "Agent", "Date"]
            )

            today = datetime.now().strftime("%Y-%m-%d")

            if report_type == "Company Overview":
                html  = build_dashboard(combined, foc_cost, df_full)
                fname = f"dashboard_{today}.html"

            elif report_type == "Agent Dashboard (AGENT1 & AGENT2)":
                IN_HOUSE   = ["AGENT1", "AGENT2"]
                agent_df   = combined[combined["Agent"].isin(IN_HOUSE)].copy()
                agent_full = df_full[df_full["Agent"].astype(str).isin(IN_HOUSE)].copy()
                html  = build_agent_dashboard(agent_df, foc_cost, agent_full, agents=IN_HOUSE)
                fname = f"dashboard_agent_{today}.html"

            else:
                AGENT3     = ["AGENT3"]
                agent3_df  = combined[combined["Agent"].isin(AGENT3)].copy()
                agent3_full = df_full[df_full["Agent"].astype(str).isin(AGENT3)].copy()
                html  = build_agent_dashboard(agent3_df, foc_cost, agent3_full, agents=AGENT3)
                fname = f"dashboard_agent3_{today}.html"

        except Exception as e:
            st.error(f"Error generating dashboard: {e}")
            st.stop()

    st.success(f"✅ Dashboard ready!")

    st.download_button(
        label="📥 Download Dashboard HTML",
        data=html.encode("utf-8"),
        file_name=fname,
        mime="text/html",
        use_container_width=True,
        type="primary",
    )

    st.info(
        "**How to view:** After downloading, open the HTML file in any browser "
        "(Chrome, Safari). Works on phone and laptop."
    )
