import streamlit as st


def init_state():
    if "cleaning_done" not in st.session_state:
        st.session_state.cleaning_done = False

    # results[brand] = {
    #   "final_xlsx": bytes,
    #   "df_main_clean": df,
    #   "df_piece_clean": df,
    #   "df_decompte_sum": df,
    #   "kpi_main": df,
    #   "kpi_piece": df,
    # }
    if "results" not in st.session_state:
        st.session_state.results = {}

    if "global_merge_bytes" not in st.session_state:
        st.session_state.global_merge_bytes = None

    if "global_merge_error" not in st.session_state:
        st.session_state.global_merge_error = None