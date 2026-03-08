# utils/session_restore.py
from __future__ import annotations

from io import BytesIO
import pandas as pd
import streamlit as st

from utils.persistence import list_artifacts, get_artifact_bytes, load_df

PREFERRED_ORDER = ["TAS", "Peugeot", "Citroen"]


def _get_artifact_by_scope_and_suffix(run_id: int, scope: str, suffix: str) -> bytes | None:
    arts = list_artifacts(run_id)
    for a in arts:
        if a["scope"] == scope and a["name"].lower().endswith(suffix.lower()):
            return get_artifact_bytes(a["id"])
    return None


def _get_artifact_by_scope_and_name(run_id: int, scope: str, exact_name: str) -> bytes | None:
    arts = list_artifacts(run_id)
    for a in arts:
        if a["scope"] == scope and a["name"] == exact_name:
            return get_artifact_bytes(a["id"])
    return None


def _get_latest_artifact_matching(run_id: int, scope: str, prefix: str, suffix: str) -> bytes | None:
    """
    Find latest artifact in given scope whose name matches:
      name startswith(prefix) AND endswith(suffix)
    Uses artifact id order as "latest".
    """
    arts = list_artifacts(run_id)
    matches = []
    for a in arts:
        if a["scope"] != scope:
            continue
        name = a.get("name") or ""
        if name.startswith(prefix) and name.lower().endswith(suffix.lower()):
            matches.append(a)

    if not matches:
        return None

    matches.sort(key=lambda x: x["id"], reverse=True)
    return get_artifact_bytes(matches[0]["id"])


def _safe_df(x) -> pd.DataFrame:
    return x if isinstance(x, pd.DataFrame) else pd.DataFrame()


def restore_session(run_id: int) -> dict:
    st.session_state["active_run_id"] = run_id

    results: dict = {}
    for brand in PREFERRED_ORDER:
        wb = _get_artifact_by_scope_and_suffix(run_id, brand, "_cleaned.xlsx")
        if wb is None:
            continue

        df_main = _safe_df(load_df(run_id, brand, f"{brand}_main.parquet"))
        df_piece = _safe_df(load_df(run_id, brand, f"{brand}_piece.parquet"))
        df_dec = _safe_df(load_df(run_id, brand, f"{brand}_decompte.parquet"))

        results[brand] = {
            "final_xlsx": wb,
            "df_main_clean": df_main,
            "df_piece_clean": df_piece,
            "df_decompte_sum": df_dec,
            "kpi_main": pd.DataFrame(),
            "kpi_piece": pd.DataFrame(),
        }

    # ✅ Global merge (optional): exact name first, then fallback to latest Dataset_Complet*.xlsx
    b = _get_artifact_by_scope_and_name(run_id, "GLOBAL", "Dataset_Complet.xlsx")
    if not b:
        b = _get_latest_artifact_matching(run_id, "GLOBAL", "Dataset_Complet", ".xlsx")
    st.session_state["global_merge_bytes"] = b
    st.session_state["global_merge_error"] = None

    # ✅ Vehicle Dates (optional)
    vd_all = _get_artifact_by_scope_and_name(run_id, "VEHICLE_DATES", "vehicle_dates_all_rows.xlsx")
    vd_earliest = _get_artifact_by_scope_and_name(
        run_id, "VEHICLE_DATES", "vehicle_dates_earliest_per_vehicle.xlsx"
    )
    st.session_state["vehicle_dates_ds1"] = vd_all
    st.session_state["vehicle_dates_ds2"] = vd_earliest

    st.session_state["df_vehicle_dates_all"] = None
    st.session_state["df_vehicle_dates_earliest"] = None
    st.session_state["df_vehicle_dates_oldest"] = None

    if vd_all:
        try:
            st.session_state["df_vehicle_dates_all"] = pd.read_excel(BytesIO(vd_all), sheet_name="extraction")
        except Exception:
            pass

    # Prefer parquet (it matches dashboard expectations)
    df_oldest = load_df(run_id, "VEHICLE_DATES", "vehicle_dates_earliest.parquet")
    if isinstance(df_oldest, pd.DataFrame) and not df_oldest.empty:
        st.session_state["df_vehicle_dates_oldest"] = df_oldest
        st.session_state["df_vehicle_dates_earliest"] = df_oldest  # compat

    st.session_state["results"] = results
    st.session_state["cleaning_done"] = bool(results)

    return results