# utils/state.py
import streamlit as st

def init_state():
    # -------------------------
    # Sessions / Persistence
    # -------------------------
    # This is the key that links ALL outputs (cleaning + merge + vehicle dates)
    # into ONE saved session.
    if "active_run_id" not in st.session_state:
        st.session_state.active_run_id = None

    # -------------------------
    # Cleaning pipeline state
    # -------------------------
    if "cleaning_done" not in st.session_state:
        st.session_state.cleaning_done = False

    # results[brand] = {
    #   "final_xlsx": bytes,
    #   "df_main_clean": df (REAL cleaned)
    #   "df_piece_clean": df (REAL cleaned)
    #   "df_decompte_sum": df (REAL summary)
    #   "kpi_main": df,
    #   "kpi_piece": df,
    # }
    if "results" not in st.session_state:
        st.session_state.results = {}

    # -------------------------
    # Global merge state
    # -------------------------
    if "global_merge_bytes" not in st.session_state:
        st.session_state.global_merge_bytes = None

    if "global_merge_error" not in st.session_state:
        st.session_state.global_merge_error = None

    # -------------------------
    # Vehicle Dates (Page 3) state
    # -------------------------
    if "vehicle_dates_ds1" not in st.session_state:
        st.session_state.vehicle_dates_ds1 = None  # bytes of vehicle_dates_all_rows.xlsx

    if "vehicle_dates_ds2" not in st.session_state:
        st.session_state.vehicle_dates_ds2 = None  # bytes of vehicle_dates_earliest_per_vehicle.xlsx

    # preview dataframe (optional)
    if "df_vehicle_dates_all" not in st.session_state:
        st.session_state.df_vehicle_dates_all = None

    # ✅ IMPORTANT: keep BOTH keys for compatibility
    # - Vehicle Dates page writes df_vehicle_dates_earliest
    # - Dashboard / restore expects df_vehicle_dates_oldest
    if "df_vehicle_dates_earliest" not in st.session_state:
        st.session_state.df_vehicle_dates_earliest = None

    if "df_vehicle_dates_oldest" not in st.session_state:
        st.session_state.df_vehicle_dates_oldest = None