import streamlit as st
from datetime import datetime

from utils.state import init_state
from utils.merge_global import build_global_from_cleaned_bytes
from utils.ui import inject_css, hero, section, card_open, card_close

init_state()

st.set_page_config(page_title="Results", layout="wide")
inject_css()
hero("Results", "Download per-brand outputs, preview cleaned data, and build the global merged workbook.")

# ---- Guard rails ----
if not st.session_state.get("cleaning_done"):
    st.info("Run Cleaning first.")
    st.stop()

results = st.session_state.get("results", {})
if not results:
    st.warning("No results found.")
    st.stop()

# ---- helpers ----
def get_cleaned_bytes(res: dict):
    """
    Backward/forward compatible:
    - new key: cleaned_xlsx
    - old key: final_xlsx (or final_bytes)
    """
    return res.get("cleaned_xlsx") or res.get("final_xlsx") or res.get("final_bytes")

# ---- Per-brand downloads + preview ----
section("Per-brand outputs")

preferred_order = ["TAS", "Peugeot", "Citroen"]
ordered_brands = [b for b in preferred_order if b in results] + [b for b in results if b not in preferred_order]

for brand in ordered_brands:
    res = results[brand]

    section(brand)
    card_open()

    cleaned_bytes = get_cleaned_bytes(res)
    c1, c2 = st.columns([1, 2])

    with c1:
        if cleaned_bytes:
            st.download_button(
                label=f"⬇️ Download {brand}_cleaned.xlsx",
                data=cleaned_bytes,
                file_name=f"{brand}_cleaned.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_{brand}_cleaned",
                use_container_width=True,
            )
        else:
            st.warning(f"{brand}: No Excel bytes found (expected cleaned_xlsx / final_xlsx / final_bytes).")

    with c2:
        st.markdown(
            "<div class='muted'>Use the preview below to quickly verify the cleaned sheets.</div>",
            unsafe_allow_html=True,
        )

    with st.expander(f"Preview {brand} data"):
        cA, cB = st.columns(2)

        with cA:
            st.markdown("**Main d'œuvre (first 20 rows)**")
            df_main = res.get("df_main_clean")
            if df_main is not None:
                st.dataframe(df_main.head(20), use_container_width=True)
            else:
                st.info("df_main_clean not found.")

        with cB:
            st.markdown("**Pièces (first 20 rows)**")
            df_piece = res.get("df_piece_clean")
            if df_piece is not None:
                st.dataframe(df_piece.head(20), use_container_width=True)
            else:
                st.info("df_piece_clean not found.")

        st.markdown("**Décompte (preview)**")
        df_dec = res.get("df_decompte_sum")
        if df_dec is not None:
            st.dataframe(df_dec.head(50), use_container_width=True)
        else:
            st.info("df_decompte_sum not found.")

        st.markdown("**KPI Main**")
        kpi_main = res.get("kpi_main")
        if kpi_main is not None:
            st.dataframe(kpi_main, use_container_width=True)
        else:
            st.info("kpi_main not found.")

        st.markdown("**KPI Piece**")
        kpi_piece = res.get("kpi_piece")
        if kpi_piece is not None:
            st.dataframe(kpi_piece, use_container_width=True)
        else:
            st.info("kpi_piece not found.")

    card_close()

st.divider()

# ---- GLOBAL MERGE ----
section("Global Merge (All Brands)")
card_open()

if st.session_state.get("global_merge_bytes"):
    st.success("Global workbook is ready to download.")

build_clicked = st.button("Build Global Merged Workbook", type="primary", key="btn_build_global")

if build_clicked:
    brand_to_bytes = {}

    # add preferred brands first (if present)
    for b in preferred_order:
        if b in results:
            bbytes = get_cleaned_bytes(results[b])
            if bbytes:
                brand_to_bytes[b] = bbytes

    # add remaining brands
    for b, res in results.items():
        if b not in brand_to_bytes:
            bbytes = get_cleaned_bytes(res)
            if bbytes:
                brand_to_bytes[b] = bbytes

    missing = [b for b in preferred_order if b not in brand_to_bytes]
    if missing:
        st.warning(f"Missing brand files for merge: {', '.join(missing)} (will merge what exists).")

    if len(brand_to_bytes) < 2:
        st.error("Need at least 2 brand workbooks to build the global merge.")
    else:
        try:
            with st.spinner("Building global merged workbook..."):
                merged_bytes = build_global_from_cleaned_bytes(brand_to_bytes)

            st.session_state.global_merge_bytes = merged_bytes
            st.session_state.global_merge_error = None
            st.success("Global workbook created.")
        except Exception as e:
            st.session_state.global_merge_bytes = None
            st.session_state.global_merge_error = str(e)
            st.error(f"Global merge failed: {e}")

if st.session_state.get("global_merge_error"):
    st.error(st.session_state.global_merge_error)

global_bytes = st.session_state.get("global_merge_bytes")
if global_bytes:
    st.download_button(
        label="⬇️ Download Dataset_Complet.xlsx",
        data=global_bytes,
        file_name="Dataset_Complet.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="dl_global",
        use_container_width=True,
    )

card_close()