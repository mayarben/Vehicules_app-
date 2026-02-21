import pandas as pd

from .exporters import export_brand_final_excel
from .rules import normalize_columns, safe_float, count_designations


def run_brand_pipeline(brand: str, file_main, file_piece, file_decompte) -> dict:
    """
    file_main / file_piece / file_decompte: Streamlit UploadedFile (or file-like)
    Returns dict with preview dfs + final bytes
    """

    # ---- READ EXCEL (single-sheet files) ----
    df_main = pd.read_excel(file_main) if file_main else pd.DataFrame()
    df_piece = pd.read_excel(file_piece) if file_piece else pd.DataFrame()
    df_decompte = pd.read_excel(file_decompte) if file_decompte else pd.DataFrame()

    # ---- Preview normalization (not the “real” cleaning) ----
    df_main_preview = normalize_columns(df_main)
    df_piece_preview = normalize_columns(df_piece)
    df_dec_preview = normalize_columns(df_decompte)

    for df in (df_main_preview, df_piece_preview, df_dec_preview):
        for c in df.columns:
            lc = str(c).lower()
            if ("total" in lc) or ("htva" in lc) or ("montant" in lc):
                df[c] = df[c].apply(safe_float)

    kpi_main = pd.DataFrame({
        "metric": ["rows", "designation_count"],
        "value": [len(df_main_preview), count_designations(df_main_preview)]
    })

    kpi_piece = pd.DataFrame({
        "metric": ["rows", "designation_count"],
        "value": [len(df_piece_preview), count_designations(df_piece_preview)]
    })

    # ---- Final bytes (exporter does the real cleaning + workbook) ----
    final_bytes = export_brand_final_excel(
        brand=brand,
        df_main_raw=df_main,
        df_piece_raw=df_piece,
        df_decompte_raw=df_decompte,
    )

    return {
        "final_xlsx": final_bytes,
        "df_main_clean": df_main_preview,
        "df_piece_clean": df_piece_preview,
        "df_decompte_sum": df_dec_preview,
        "kpi_main": kpi_main,
        "kpi_piece": kpi_piece,
    }