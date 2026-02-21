import streamlit as st

from cleaning.pipeline import run_brand_pipeline
from utils.ui import inject_css, hero, section, card_open, card_close

st.set_page_config(page_title="Cleaning", layout="wide")

inject_css()
hero("Cleaning", "Upload the 3 files per brand (Main d'œuvre + Pièces + Décompte), then click Run Cleaning.")

# -------------------------
# Reset state
# -------------------------
colA, colB = st.columns([1, 6])
with colA:
    if st.button("Reset"):
        st.session_state.clear()
        st.rerun()

# -------------------------
# Upload blocks
# -------------------------
def brand_upload_block(brand: str):
    section(brand)
    card_open()

    c1, c2, c3 = st.columns(3)
    with c1:
        f_main = st.file_uploader(
            f"{brand} - Main d'œuvre (.xlsx)",
            type=["xlsx"],
            key=f"{brand}_main",
        )
    with c2:
        f_piece = st.file_uploader(
            f"{brand} - Pièces (.xlsx)",
            type=["xlsx"],
            key=f"{brand}_piece",
        )
    with c3:
        f_decompte = st.file_uploader(
            f"{brand} - Décompte (.xlsx)",
            type=["xlsx"],
            key=f"{brand}_decompte",
        )

    card_close()
    return f_main, f_piece, f_decompte


tas_main, tas_piece, tas_decompte = brand_upload_block("TAS")
peu_main, peu_piece, peu_decompte = brand_upload_block("Peugeot")
cit_main, cit_piece, cit_decompte = brand_upload_block("Citroen")

all_ok = all(
    [
        tas_main, tas_piece, tas_decompte,
        peu_main, peu_piece, peu_decompte,
        cit_main, cit_piece, cit_decompte,
    ]
)

st.divider()

# -------------------------
# Run cleaning
# -------------------------
run_disabled = (not all_ok) or st.session_state.get("cleaning_done", False)

if st.button("Run Cleaning", type="primary", disabled=run_disabled):
    results = {}
    try:
        with st.spinner("Cleaning TAS..."):
            results["TAS"] = run_brand_pipeline("TAS", tas_main, tas_piece, tas_decompte)

        with st.spinner("Cleaning Peugeot..."):
            results["Peugeot"] = run_brand_pipeline("Peugeot", peu_main, peu_piece, peu_decompte)

        with st.spinner("Cleaning Citroen..."):
            results["Citroen"] = run_brand_pipeline("Citroen", cit_main, cit_piece, cit_decompte)

        st.session_state["results"] = results
        st.session_state["cleaning_done"] = True
        st.rerun()

    except Exception as e:
        st.session_state["cleaning_done"] = False
        st.error(f"Cleaning failed: {e}")

# -------------------------
# ✅ Message UNDER the button (persistent after rerun)
# -------------------------
if st.session_state.get("cleaning_done"):
    st.success("✅ Cleaning done. Go to Results page to download outputs.")
else:
    st.markdown("<div class='muted'>Tip: upload all 9 files first, then run cleaning.</div>", unsafe_allow_html=True)