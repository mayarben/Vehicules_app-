# cleaning/pipeline.py
import pandas as pd

from cleaning.exporters import (
    export_brand_final_excel,
    clean_main_df,
    clean_piece_df,
    build_decompte_summary,
)

def load_raw_inputs(file_main, file_piece, file_decompte):
    """
    Utility (for debugging / reporting screenshots):
    Reads the 3 raw Excel inputs and prints a quick preview (columns + first rows).
    """
    df_main_raw = pd.read_excel(file_main) if file_main else pd.DataFrame()
    df_piece_raw = pd.read_excel(file_piece) if file_piece else pd.DataFrame()
    df_decompte_raw = pd.read_excel(file_decompte) if file_decompte else pd.DataFrame()

    print("MAIN columns:", list(df_main_raw.columns))
    print(df_main_raw.head(10))

    print("PIECE columns:", list(df_piece_raw.columns))
    print(df_piece_raw.head(10))

    print("DECOMPTE columns:", list(df_decompte_raw.columns))
    print(df_decompte_raw.head(20))

    return df_main_raw, df_piece_raw, df_decompte_raw


def run_brand_pipeline(brand: str, file_main, file_piece, file_decompte) -> dict:
    # Read raw inputs
    df_main_raw = pd.read_excel(file_main) if file_main else pd.DataFrame()
    df_piece_raw = pd.read_excel(file_piece) if file_piece else pd.DataFrame()
    df_decompte_raw = pd.read_excel(file_decompte) if file_decompte else pd.DataFrame()

    # Clean + summarize
    df_main_clean = clean_main_df(df_main_raw) if not df_main_raw.empty else pd.DataFrame()
    df_piece_clean = clean_piece_df(df_piece_raw) if not df_piece_raw.empty else pd.DataFrame()
    df_decompte_sum = build_decompte_summary(df_decompte_raw) if not df_decompte_raw.empty else pd.DataFrame()

    # Export formatted brand workbook
    final_bytes = export_brand_final_excel(
        brand=brand,
        df_main_raw=df_main_raw,
        df_piece_raw=df_piece_raw,
        df_decompte_raw=df_decompte_raw,
    )

    # Keep keys stable (Results page expects these)
    return {
        "final_xlsx": final_bytes,
        "df_main_clean": df_main_clean,
        "df_piece_clean": df_piece_clean,
        "df_decompte_sum": df_decompte_sum,
        "kpi_main": pd.DataFrame(),
        "kpi_piece": pd.DataFrame(),
    }