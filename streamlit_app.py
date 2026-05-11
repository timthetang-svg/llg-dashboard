#!/usr/bin/env python3
"""
LLG Industries — Cloud Dashboard
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

# ── Password config ────────────────────────────────────────────────────────────

ROLE_MAP = {
    st.secrets["passwords"]["boss"]:   "boss",
    st.secrets["passwords"]["agent1"]: "agent1",
    st.secrets["passwords"]["agent2"]: "agent2",
}

# ── Login page ─────────────────────────────────────────────────────────────────

if "role" not in st.session_state:
    st.session_state.role = None

if st.session_state.role is None:
    st.markdown(
        "<h1 style='color:#1a3c5e;font-family:Segoe UI'>📊 LLG Industries — Sales Dashboard</h1>",
        unsafe_allow_html=True,
    )
    st.markdown(" ")
    with st.form("login"):
        pw = st.text_input("Password", type="password", placeholder="Enter your password")
        submitted = st.form_submit_button("Login →", use_container_width=True, type="primary")
    if submitted:
        role = ROLE_MAP.get(pw)
        if role:
            st.session_state.role = role
            st.rerun()
        else:
            st.error("Incorrect password. Please try again.")
    st.stop()

# ── Dashboard (logged in) ──────────────────────────────────────────────────────

role = st.session_state.role

col_title, col_logout = st.columns([5, 1])
with col_title:
    st.markdown(
        "<h1 style='color:#1a3c5e;font-family:Segoe UI'>📊 LLG Industries — Sales Dashboard</h1>",
        unsafe_allow_html=True,
    )
with col_logout:
    st.markdown(" ")
    if st.button("Logout", use_container_width=True):
        st.session_state.role = None
        st.rerun()

st.markdown("Upload your sales Excel file(s), choose report type, then click **Generate**.")
st.divider()

col1, col2 = st.columns(2)
with col1:
    lle_file = st.file_uploader("LLE Sales Data (.xlsx)", type="xlsx", key="lle")
with col2:
    llg_file = st.file_uploader("LLG Sales Data (.xlsx)", type="xlsx", key="llg")

# Report types & agent filtering per role
if role == "boss":
    report_type = st.radio(
        "Report type",
        ["Company Overview", "Agent Dashboard (AGENT1 & AGENT2)", "Agent 3 Dashboard"],
        horizontal=True,
    )
    filter_agents = None  # boss sees all
elif role == "agent1":
    report_type = "agent_individual"
    filter_agents = ["AGENT1"]
    st.info("📋 Your personal sales dashboard — AGENT1")
elif role == "agent2":
    report_type = "agent_individual"
    filter_agents = ["AGENT2"]
    st.info("📋 Your personal sales dashboard — AGENT2")

st.divider()

if st.button("🚀 Generate Dashboard", type="primary", use_container_width=True):
    if not lle_file and not llg_file:
        st.error("Please upload at least one Excel file.")
        st.stop()

    with st.spinner("Reading data and building charts… (20–40 seconds)"):
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
                agents     = ["AGENT1", "AGENT2"]
                agent_df   = combined[combined["Agent"].isin(agents)].copy()
                agent_full = df_full[df_full["Agent"].astype(str).isin(agents)].copy()
                html  = build_agent_dashboard(agent_df, foc_cost, agent_full, agents=agents)
                fname = f"dashboard_agent_{today}.html"

            elif report_type == "Agent 3 Dashboard":
                agents      = ["AGENT3"]
                agent3_df   = combined[combined["Agent"].isin(agents)].copy()
                agent3_full = df_full[df_full["Agent"].astype(str).isin(agents)].copy()
                html  = build_agent_dashboard(agent3_df, foc_cost, agent3_full, agents=agents)
                fname = f"dashboard_agent3_{today}.html"

            else:  # agent_individual
                agent_df   = combined[combined["Agent"].isin(filter_agents)].copy()
                agent_full = df_full[df_full["Agent"].astype(str).isin(filter_agents)].copy()
                html  = build_agent_dashboard(agent_df, foc_cost, agent_full, agents=filter_agents)
                fname = f"dashboard_{filter_agents[0].lower()}_{today}.html"

        except Exception as e:
            st.error(f"Error: {e}")
            st.stop()

    st.success("✅ Dashboard ready!")
    st.download_button(
        label="📥 Download Dashboard HTML",
        data=html.encode("utf-8"),
        file_name=fname,
        mime="text/html",
        use_container_width=True,
        type="primary",
    )
    st.info("Open the downloaded file in Chrome or Safari — works on phone and laptop.")
