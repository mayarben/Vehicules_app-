# pages/1_Cleaning.py
import streamlit as st
import pandas as pd

from cleaning.pipeline import run_brand_pipeline
from utils.ui import inject_css, hero, section, card_open, card_close
from utils.persistence import DB_PATH, create_run, save_artifact, save_df_parquet, XLSX_MIME
from utils.state import init_state

st.set_page_config(page_title="Cleaning", layout="wide")

# Ensure all expected session keys exist
init_state()

inject_css()

hero(
    "Cleaning 3 Brands",
    "Upload the 3 files per brand (Main d'œuvre + Pièces + Décompte), then click Run Cleaning.",
)

st.caption(f"DB file: {DB_PATH}")

# ---------------------------------------------------------
# OPTIONAL: "Start new session" button (clean slate in UI)
# ---------------------------------------------------------
if st.session_state.get("active_run_id") is not None and st.session_state.get("cleaning_done"):
    st.info("A previous session is currently loaded. Running Cleaning will start a NEW session automatically.")
    if st.button("Start a new session now", use_container_width=True):
        st.session_state["active_run_id"] = None
        st.session_state["results"] = {}
        st.session_state["cleaning_done"] = False
        st.session_state["global_merge_bytes"] = None
        st.session_state["global_merge_error"] = None
        st.session_state["vehicle_dates_ds1"] = None
        st.session_state["vehicle_dates_ds2"] = None
        st.session_state["df_vehicle_dates_all"] = None
        st.session_state["df_vehicle_dates_earliest"] = None
        st.session_state["df_vehicle_dates_oldest"] = None
        st.rerun()

# Reset (kept, but you won't need it anymore)
colA, colB = st.columns([1, 6])
with colA:
    if st.button("Reset"):
        for k in ["results", "cleaning_done", "_cleaning_running", "active_run_id"]:
            st.session_state.pop(k, None)

        for k in list(st.session_state.keys()):
            if k.endswith(("_main", "_piece", "_decompte")) and (
                k.startswith("TAS") or k.startswith("Peugeot") or k.startswith("Citroen")
            ):
                st.session_state.pop(k, None)

        st.success("State reset. Re-upload files.")

def brand_upload_block(brand: str):
    section(brand)
    card_open()

    c1, c2, c3 = st.columns(3)
    with c1:
        f_main = st.file_uploader(f"{brand} - Main d'œuvre (.xlsx)", type=["xlsx"], key=f"{brand}_main")
    with c2:
        f_piece = st.file_uploader(f"{brand} - Pièces (.xlsx)", type=["xlsx"], key=f"{brand}_piece")
    with c3:
        f_decompte = st.file_uploader(f"{brand} - Décompte (.xlsx)", type=["xlsx"], key=f"{brand}_decompte")

    card_close()
    return f_main, f_piece, f_decompte

tas_main, tas_piece, tas_decompte = brand_upload_block("TAS")
peu_main, peu_piece, peu_decompte = brand_upload_block("Peugeot")
cit_main, cit_piece, cit_decompte = brand_upload_block("Citroen")

all_ok = all([tas_main, tas_piece, tas_decompte, peu_main, peu_piece, peu_decompte, cit_main, cit_piece, cit_decompte])
st.divider()

# ✅ FIX: do NOT disable just because cleaning_done from old session
run_disabled = (not all_ok) or st.session_state.get("_cleaning_running", False)

if st.button("Run Cleaning", type="primary", disabled=run_disabled):
    st.session_state["_cleaning_running"] = True
    try:
        # ✅ NEW: Start a brand new session EVERY time user clicks Run Cleaning
        run_id = create_run("full_session", run_name="Session (Cleaning)")
        st.session_state["active_run_id"] = run_id

        # ✅ Clear previous in-memory outputs (do NOT delete DB history)
        st.session_state["results"] = {}
        st.session_state["cleaning_done"] = False
        st.session_state["global_merge_bytes"] = None
        st.session_state["global_merge_error"] = None
        st.session_state["vehicle_dates_ds1"] = None
        st.session_state["vehicle_dates_ds2"] = None
        st.session_state["df_vehicle_dates_all"] = None
        st.session_state["df_vehicle_dates_earliest"] = None
        st.session_state["df_vehicle_dates_oldest"] = None

        results = {}

        with st.spinner("Cleaning TAS..."):
            results["TAS"] = run_brand_pipeline("TAS", tas_main, tas_piece, tas_decompte)

        with st.spinner("Cleaning Peugeot..."):
            results["Peugeot"] = run_brand_pipeline("Peugeot", peu_main, peu_piece, peu_decompte)

        with st.spinner("Cleaning Citroen..."):
            results["Citroen"] = run_brand_pipeline("Citroen", cit_main, cit_piece, cit_decompte)

        # ✅ Save artifacts + cleaned dfs into the NEW run_id
        for brand, res in results.items():
            b = res.get("final_xlsx")
            if b:
                save_artifact(run_id, brand, f"{brand}_cleaned.xlsx", XLSX_MIME, b)

            if isinstance(res.get("df_main_clean"), pd.DataFrame):
                save_df_parquet(run_id, brand, f"{brand}_main.parquet", res["df_main_clean"])
            if isinstance(res.get("df_piece_clean"), pd.DataFrame):
                save_df_parquet(run_id, brand, f"{brand}_piece.parquet", res["df_piece_clean"])
            if isinstance(res.get("df_decompte_sum"), pd.DataFrame):
                save_df_parquet(run_id, brand, f"{brand}_decompte.parquet", res["df_decompte_sum"])

        st.session_state["results"] = results
        st.session_state["cleaning_done"] = True

        st.success(f"✅ Cleaning done + NEW session saved (run_id={run_id}). Go to Sessions or Results.")
        st.rerun()

    except Exception as e:
        st.session_state["cleaning_done"] = False
        st.error(f"Cleaning failed: {e}")
        st.exception(e)

    finally:
        st.session_state["_cleaning_running"] = False

if st.session_state.get("cleaning_done", False):
    st.success("✅ Cleaning done. Go to Results page.")