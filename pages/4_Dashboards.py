# pages/4_Dashboards.py
from __future__ import annotations

import streamlit as st

from utils.state import init_state
from utils.session_restore import restore_session
from utils.persistence import list_runs
from utils.dashboards_core import render_dashboards

# Safe page config (won't crash if already set elsewhere)
try:
    st.set_page_config(page_title="Dashboard", layout="wide")
except Exception:
    pass

init_state()
st.title("Dashboard")

# ✅ Auto-select latest run if none selected
if st.session_state.get("active_run_id") is None:
    runs = list_runs(limit=1)
    if runs:
        st.session_state["active_run_id"] = runs[0]["id"]

rid = st.session_state.get("active_run_id")

# ✅ Auto-restore session if results missing/empty
if rid and (not isinstance(st.session_state.get("results"), dict) or not st.session_state["results"]):
    restore_session(rid)

# Render dashboards
if not isinstance(st.session_state.get("results"), dict) or not st.session_state["results"]:
    st.warning("No cleaning results found. Please run Cleaning first or select a Session.")
else:
    render_dashboards()