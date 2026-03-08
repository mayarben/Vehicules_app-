# pages/2_Results of cleaning.py
import streamlit as st
import pandas as pd

from utils.ui import inject_css, hero, section, card_open, card_close
from utils.merge_global import build_global_from_cleaned_bytes
from utils.state import init_state

# persistence (sessions)
from utils.persistence import (
    list_runs,
    list_artifacts,
    get_artifact_bytes,
    load_df,
    save_artifact,
    XLSX_MIME,
)

st.set_page_config(page_title="Results", layout="wide")

# ✅ Ensure all expected session keys exist
init_state()

inject_css()
hero("Results", "Download per-brand outputs, preview cleaned data, and build the global merged workbook.")


# =========================================================
# Helpers (load from persistence)
# =========================================================
PREFERRED_ORDER = ["TAS", "Peugeot", "Citroen"]

def _pick_latest_run_id() -> int | None:
    runs = list_runs(limit=1)
    return runs[0]["id"] if runs else None

def _get_artifact_by_scope_and_suffix(run_id: int, scope: str, suffix: str) -> bytes | None:
    arts = list_artifacts(run_id)
    for a in arts:
        if a["scope"] == scope and a["name"].lower().endswith(suffix.lower()):
            return get_artifact_bytes(a["id"])
    return None

def _load_results_from_run(run_id: int) -> dict:
    out = {}
    for brand in PREFERRED_ORDER:
        b = _get_artifact_by_scope_and_suffix(run_id, brand, "_cleaned.xlsx")
        df_main = load_df(run_id, brand, f"{brand}_main.parquet")
        df_piece = load_df(run_id, brand, f"{brand}_piece.parquet")
        df_dec = load_df(run_id, brand, f"{brand}_decompte.parquet")

        if b is not None:
            out[brand] = {
                "final_xlsx": b,
                "df_main_clean": df_main if isinstance(df_main, pd.DataFrame) else pd.DataFrame(),
                "df_piece_clean": df_piece if isinstance(df_piece, pd.DataFrame) else pd.DataFrame(),
                "df_decompte_sum": df_dec if isinstance(df_dec, pd.DataFrame) else pd.DataFrame(),
            }
    return out

def _get_cleaned_bytes(res: dict) -> bytes | None:
    return res.get("final_xlsx") or res.get("cleaned_xlsx") or res.get("final_bytes")


# =========================================================
# Determine source of truth: session (DB) or live state
# =========================================================
run_id = st.session_state.get("active_run_id", None)

if run_id is None:
    run_id = _pick_latest_run_id()
    if run_id is not None:
        st.session_state["active_run_id"] = run_id

results = {}
if run_id is not None:
    results = _load_results_from_run(run_id)

if not results:
    results = st.session_state.get("results", {})

if not results:
    st.warning("No results found. Run Cleaning first OR select a Session.")
    st.stop()


# =========================================================
# Per-brand downloads + preview
# =========================================================
section("Per-brand outputs")
card_open()

ordered_brands = [b for b in PREFERRED_ORDER if b in results] + [b for b in results if b not in PREFERRED_ORDER]

for brand in ordered_brands:
    res = results[brand]
    cleaned_bytes = res.get("final_xlsx")

    st.markdown(f"### {brand}")

    c1, c2 = st.columns([1, 2])
    with c1:
        if cleaned_bytes:
            st.download_button(
                label=f"⬇️ Download {brand}_cleaned.xlsx",
                data=cleaned_bytes,
                file_name=f"{brand}_cleaned.xlsx",
                mime=XLSX_MIME,
                key=f"dl_{brand}_cleaned",
                use_container_width=True,
            )
        else:
            st.warning(f"{brand}: No Excel bytes found.")

    with c2:
        st.markdown("<div class='muted'>Preview below uses cleaned DataFrames when available.</div>", unsafe_allow_html=True)

    with st.expander(f"Preview {brand} data"):
        cA, cB = st.columns(2)

        with cA:
            st.markdown("**Main d'œuvre (first 20 rows)**")
            df_main = res.get("df_main_clean")
            if isinstance(df_main, pd.DataFrame) and not df_main.empty:
                st.dataframe(df_main.head(20), use_container_width=True)
            else:
                st.info("No df_main_clean available.")

        with cB:
            st.markdown("**Pièces (first 20 rows)**")
            df_piece = res.get("df_piece_clean")
            if isinstance(df_piece, pd.DataFrame) and not df_piece.empty:
                st.dataframe(df_piece.head(20), use_container_width=True)
            else:
                st.info("No df_piece_clean available.")

        st.markdown("**Décompte (preview)**")
        df_dec = res.get("df_decompte_sum")
        if isinstance(df_dec, pd.DataFrame) and not df_dec.empty:
            st.dataframe(df_dec.head(50), use_container_width=True)
        else:
            st.info("No df_decompte_sum available.")

    st.divider()

card_close()
st.divider()


# =========================================================
# GLOBAL MERGE
# =========================================================
section("Global Merge (All Brands)")
card_open()

global_bytes = st.session_state.get("global_merge_bytes")
has_global = bool(global_bytes)

# ✅ Single status message (no duplicate green banners)
if has_global:
    st.info("Global workbook is ready to download.")

# ✅ Disable build if already present (prevents double messages + UNIQUE insert collisions)
build_clicked = st.button(
    "Build Global Merged Workbook",
    type="primary",
    key="btn_build_global",
    disabled=has_global,
)

# Optional: allow rebuild if you really want it
rebuild = False
if has_global:
    with st.expander("Need to rebuild?"):
        rebuild = st.checkbox("Allow rebuild (will overwrite the existing global workbook in session).", value=False)
    if rebuild:
        build_clicked = st.button(
            "Rebuild Global Merged Workbook",
            type="primary",
            key="btn_rebuild_global",
            disabled=False,
        )

if build_clicked:
    brand_to_bytes = {}

    for b in PREFERRED_ORDER:
        if b in results:
            bbytes = _get_cleaned_bytes(results[b])
            if bbytes:
                brand_to_bytes[b] = bbytes

    for b, res in results.items():
        if b not in brand_to_bytes:
            bbytes = _get_cleaned_bytes(res)
            if bbytes:
                brand_to_bytes[b] = bbytes

    if len(brand_to_bytes) < 2:
        st.error("Need at least 2 brand workbooks to build the global merge.")
    else:
        try:
            with st.spinner("Building global merged workbook..."):
                merged_bytes = build_global_from_cleaned_bytes(brand_to_bytes)

            st.session_state["global_merge_bytes"] = merged_bytes
            st.session_state["global_merge_error"] = None

            # ✅ Save into same session
            rid = st.session_state.get("active_run_id")
            if rid:
                # If your DB has a UNIQUE constraint on (run_id, scope, name),
                # rebuilding would fail on INSERT. The simplest safe approach is:
                # - only allow rebuild when checkbox is on (handled above)
                # - and save with a timestamped name OR overwrite logic in persistence.
                #
                # Here: use a stable name on first build; on rebuild use a timestamped name.
                name = "Dataset_Complet.xlsx"
                if rebuild:
                    # keep the stable download name too (so Sessions page works),
                    # but also avoid UNIQUE collisions if your DB is strict
                    # by appending a timestamped copy:
                    ts_name = f"Dataset_Complet_{pd.Timestamp.now():%Y%m%d_%H%M%S}.xlsx"
                    try:
                        save_artifact(rid, "GLOBAL", ts_name, XLSX_MIME, merged_bytes)
                    except Exception:
                        pass

                # Save/overwrite the main one (if your DB insert is strict and fails,
                # you'll want to apply the UPSERT change in utils/persistence.py)
                save_artifact(
                    run_id=rid,
                    scope="GLOBAL",
                    name=name,
                    content_type=XLSX_MIME,
                    data=merged_bytes,
                )

            st.success("✅ Global workbook created and saved to session.")
        except Exception as e:
            st.session_state["global_merge_bytes"] = None
            st.session_state["global_merge_error"] = str(e)
            st.error(f"Global merge failed: {e}")
            st.exception(e)

if st.session_state.get("global_merge_error"):
    st.error(st.session_state["global_merge_error"])

global_bytes = st.session_state.get("global_merge_bytes")
if global_bytes:
    st.download_button(
        label="⬇️ Download Dataset_Complet.xlsx",
        data=global_bytes,
        file_name="Dataset_Complet.xlsx",
        mime=XLSX_MIME,
        key="dl_global",
        use_container_width=True,
    )

card_close()