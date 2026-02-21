from __future__ import annotations

from io import BytesIO
import re
import unicodedata
from typing import Dict, List, Optional

import numpy as np
import pandas as pd
import xlsxwriter


# ======================================================
# Excel-safe sanitizer
# ======================================================
def _san(v):
    if v is None:
        return ""
    if isinstance(v, (float, int)) and not isinstance(v, bool):
        try:
            if np.isnan(v) or np.isinf(v):
                return 0.0
        except Exception:
            pass
        return v
    return v


# ======================================================
# Column normalization + picking
# ======================================================
def normalize_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = (
        df.columns.astype(str)
        .str.replace("\u00A0", " ", regex=False)   # NBSP -> space
        .str.strip()
        .str.replace("\n", " ", regex=False)
        .str.replace("\r", " ", regex=False)
    )
    return df


def pick_col(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    cols = {str(c).lower().strip(): c for c in df.columns}
    for cand in candidates:
        k = cand.lower().strip()
        if k in cols:
            return cols[k]
    return None


# ======================================================
# Numeric cleaning
# ======================================================
def clean_numeric(series: pd.Series) -> pd.Series:
    """
    Handles:
    - "1 234,567" -> 1234.567
    - "1234,567"  -> 1234.567
    - "1234.567"  -> 1234.567
    """
    s = series.astype(str)
    s = s.str.replace("\u00A0", " ", regex=False)
    s = s.str.replace("\xa0", " ", regex=False)
    s = s.str.replace(" ", "", regex=False)
    s = s.str.replace(",", ".", regex=False)
    return pd.to_numeric(s, errors="coerce")


def safe_float(x) -> float:
    try:
        if pd.isna(x) or np.isinf(x):
            return 0.0
        return float(x)
    except Exception:
        return 0.0


# ======================================================
# Vehicle ID normalization (17-xxxxxx)
# ======================================================
def normalize_vehicle_id(value) -> str:
    if pd.isna(value):
        return ""
    s = str(value).strip()
    if s == "":
        return ""

    # 356771.0 -> 356771
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]

    s = re.sub(r"\s+", "", s)

    # already correct: 17-XXXXXX (6 digits)
    if re.fullmatch(r"17-\d{6}", s):
        return s

    # match 17xxxxxx or 17-xxxxxx
    m = re.search(r"17-?(\d{6})", s)
    if m:
        return f"17-{m.group(1)}"

    # digits only
    if s.isdigit():
        if s.startswith("17") and len(s) >= 8:
            return f"17-{s[-6:]}"          # 17351426 -> 17-351426
        return f"17-{s.zfill(6)[-6:]}"     # 351304 -> 17-351304

    return s


def fix_vehicle(v):
    """More permissive fixer used in your merge scripts (7-xxxxx -> 17-xxxxx, etc.)."""
    if pd.isna(v):
        return np.nan
    s = str(v).strip()

    if s.endswith(".0"):
        s = s[:-2]
    s = s.replace("\u00A0", "").replace("\xa0", "").replace(" ", "")

    if s.startswith("17-") and s[3:].isdigit():
        return s

    if s.startswith("7-"):
        rest = s[2:]
        return "17-" + rest if rest.isdigit() else np.nan

    if s.startswith("17") and s[2:].isdigit():
        return "17-" + s[2:]

    if s.isdigit():
        return "17-" + s

    return np.nan


# ======================================================
# Text normalization for designation KPI + decompte parsing
# ======================================================
def norm_text(s: str) -> str:
    if s is None:
        return ""
    s = str(s).replace("\u00a0", " ").strip().lower()
    s = s.replace("œ", "oe").replace("’", "'")
    s = "".join(ch for ch in unicodedata.normalize("NFKD", s) if not unicodedata.combining(ch))
    s = re.sub(r"\s+", " ", s)
    return s


def clean_designation_for_kpi(x) -> Optional[str]:
    if pd.isna(x):
        return None
    s = str(x).strip().lower()
    s = "".join(ch for ch in unicodedata.normalize("NFKD", s) if not unicodedata.combining(ch))
    s = re.sub(r"\s+", " ", s)
    s = s.strip()
    return s if s else None


# ======================================================
# Enforce schema / types (display rules)
# ======================================================
def enforce_schema(df: pd.DataFrame) -> pd.DataFrame:
    df = normalize_cols(df)

    col_doss = pick_col(df, ["N° Doss", "N°Doss", "No Doss", "N Doss"])
    col_veh  = pick_col(df, ["Véhicule", "Vehicule", "Vehicle", "véhicule", "vehicule"])
    col_mar  = pick_col(df, ["Marque"])
    col_no   = pick_col(df, ["N°", "No", "N"])
    col_des  = pick_col(df, ["Designation", "Désignation"])
    col_qte  = pick_col(df, ["Qté", "Qte", "qte"])
    col_mnt  = pick_col(df, ["Montant", "montant"])
    col_tot  = pick_col(df, ["Total HTVA", "total htva", "Total"])

    # ints
    for c in [col_doss, col_no]:
        if c and c in df.columns:
            df[c] = clean_numeric(df[c]).round(0).astype("Int64")

    # strings
    for c in [col_veh, col_mar, col_des]:
        if c and c in df.columns:
            df[c] = df[c].fillna("").astype(str)

    # Qté -> 1 decimal always
    if col_qte and col_qte in df.columns:
        df[col_qte] = clean_numeric(df[col_qte]).astype(float).round(1)

    # money -> 3 decimals
    for c in [col_mnt, col_tot]:
        if c and c in df.columns:
            df[c] = clean_numeric(df[c]).astype(float).round(3)

    return df


# ======================================================
# CLEAN main/piece like your notebook (filter + normalize)
# ======================================================
def _clean_common(df_raw: pd.DataFrame) -> pd.DataFrame:
    df = normalize_cols(df_raw)

    # normalize vehicle col name
    if "Vehicule" in df.columns and "véhicule" not in df.columns and "Véhicule" not in df.columns:
        df = df.rename(columns={"Vehicule": "véhicule"})
    if "Véhicule" in df.columns and "véhicule" not in df.columns:
        df = df.rename(columns={"Véhicule": "véhicule"})

    # strip/normalize all headers (already done in normalize_cols, but keep safe)
    df.columns = df.columns.astype(str).str.replace("\u00A0", " ", regex=False).str.strip()

    # ✅ ROBUST DROP: handles "Date", "Date ", "Date\u00A0", "indexe", etc.
    def _norm_colname(c: str) -> str:
        return str(c).replace("\u00A0", " ").strip().lower()

    drop_set = {
        "date",
        "indexe",
        "index",
        "idx",
        "unnamed: 0",
        "unnamed:0",
    }
    to_drop = [c for c in df.columns if _norm_colname(c) in drop_set]
    df = df.drop(columns=to_drop, errors="ignore")

    # vehicle cleanup
    if "véhicule" in df.columns:
        df["véhicule"] = df["véhicule"].apply(normalize_vehicle_id).astype(str)
        df = df[df["véhicule"].notna() & (df["véhicule"].str.strip() != "")]
        df = df.reset_index(drop=True)

    return df


def clean_main_df(df_main_raw: pd.DataFrame) -> pd.DataFrame:
    df = _clean_common(df_main_raw)

    # filter: valid vehicle and (amount or code)
    if "véhicule" in df.columns:
        veh = df["véhicule"].astype(str).str.strip()
        is_vehicle = veh.fillna("").str.match(r"^(?:\d+-)?\d+$")
    else:
        is_vehicle = pd.Series([False] * len(df), index=df.index)

    montant = pd.to_numeric(clean_numeric(df.get("Montant", pd.Series([np.nan] * len(df)))), errors="coerce")
    code = pd.to_numeric(clean_numeric(df.get("N°", pd.Series([np.nan] * len(df)))), errors="coerce")
    keep = is_vehicle & (montant.notna() | code.notna())

    df = df.loc[keep].copy()

    # enforce types/rounding
    df = enforce_schema(df)

    # fix vehicle to strict format
    if "véhicule" in df.columns:
        df["véhicule"] = df["véhicule"].apply(normalize_vehicle_id).astype(str)

    return df.reset_index(drop=True)


def clean_piece_df(df_piece_raw: pd.DataFrame) -> pd.DataFrame:
    df = _clean_common(df_piece_raw)

    # filter: valid vehicle and has amount
    if "véhicule" in df.columns:
        veh = df["véhicule"].astype(str).str.strip()
        is_vehicle = veh.fillna("").str.match(r"^(?:\d+-)?\d+$")
    else:
        is_vehicle = pd.Series([False] * len(df), index=df.index)

    montant = pd.to_numeric(clean_numeric(df.get("Montant", pd.Series([np.nan] * len(df)))), errors="coerce")
    keep = is_vehicle & montant.notna()

    df = df.loc[keep].copy()
    df = enforce_schema(df)

    if "véhicule" in df.columns:
        df["véhicule"] = df["véhicule"].apply(normalize_vehicle_id).astype(str)

    # drop empty designation
    des_col = pick_col(df, ["Désignation", "Designation"])
    if des_col:
        df[des_col] = df[des_col].astype(str).str.strip()
        df = df[df[des_col].replace({"": np.nan, "nan": np.nan, "None": np.nan, "-": np.nan}).notna()]

    return df.reset_index(drop=True)


# ======================================================
# Decompte summary parser (robust)
# ======================================================
def build_decompte_summary(df_decompte_raw: pd.DataFrame) -> pd.DataFrame:
    df = normalize_cols(df_decompte_raw)

    # ensure columns exist
    for c in ["Désignation", "Total HTVA"]:
        if c not in df.columns:
            alt = pick_col(df, [c, c.lower(), c.replace("é", "e")])
            if alt and alt != c:
                df = df.rename(columns={alt: c})
            else:
                df[c] = pd.NA

    # numeric for Total HTVA
    df["Total HTVA"] = clean_numeric(df["Total HTVA"]).astype("Float64")

    results = []
    k = 0
    current = None

    def contains_any(text: str, keys: List[str]) -> bool:
        return any(k in text for k in keys)

    MAIN_KEYS = ["total main", "main d'oeuvre", "main doeuvre", "main dœuvre", "main d oeuvre"]
    PIECE_KEYS = [
        "total pieces", "total piece",
        "total pieces de rechange", "pieces de rechange", "piece de rechange",
        "fourniture piece", "fourniture pieces", "fournitures pieces",
        "total fourniture", "total fournitures", "fourniture", "fournitures",
        "pieces", "piece",
    ]
    HTVA_KEYS = ["total htv", "total htva", "total htv/htva", "total htva/htv"]

    for _, row in df.iterrows():
        des = norm_text(row.get("Désignation", ""))
        val = row.get("Total HTVA", pd.NA)
        if pd.isna(val):
            continue
        val = float(val)

        # Start block
        if des.startswith("total main") or des == "total" or contains_any(des, MAIN_KEYS):
            if current is not None:
                if current["Total HTV/HTVA"] == 0.0:
                    current["Total HTV/HTVA"] = current["Total main d'œuvre"]
                results.append(current)

            k += 1
            current = {
                "Decompte": f"Decompte {k}",
                "Total main d'œuvre": val,
                "Total pièces": 0.0,
                "FGB": 0.0,
                "Total HTV/HTVA": 0.0,
            }
            continue

        if current is None:
            continue

        if (("total" in des and contains_any(des, ["piece", "pieces", "fourniture", "fournitures", "rechange"]))
            or contains_any(des, PIECE_KEYS)):
            current["Total pièces"] = val
        elif "fgb" in des:
            current["FGB"] = val
        elif (des.startswith("total htv") or des.startswith("total htva")
              or ("total" in des and ("htv" in des or "htva" in des))
              or contains_any(des, HTVA_KEYS)):
            current["Total HTV/HTVA"] = val
            results.append(current)
            current = None

    if current is not None:
        if current["Total HTV/HTVA"] == 0.0:
            current["Total HTV/HTVA"] = current["Total main d'œuvre"]
        results.append(current)

    summary = pd.DataFrame(results)
    if summary.empty:
        return summary

    # add TOTAL GLOBAL
    summary.loc[len(summary)] = {
        "Decompte": "TOTAL GLOBAL",
        "Total main d'œuvre": summary["Total main d'œuvre"].sum(),
        "Total pièces": summary["Total pièces"].sum(),
        "FGB": summary["FGB"].sum(),
        "Total HTV/HTVA": summary["Total HTV/HTVA"].sum(),
    }

    for c in ["Total main d'œuvre", "Total pièces", "FGB", "Total HTV/HTVA"]:
        summary[c] = pd.to_numeric(summary[c], errors="coerce").fillna(0.0).round(3)

    return summary


# ======================================================
# KPI builder (designation counts)
# ======================================================
def build_designation_kpi(df: pd.DataFrame, designation_col: str, count_col_name: str) -> pd.DataFrame:
    s = df[designation_col].apply(clean_designation_for_kpi)
    out = s.dropna().value_counts().reset_index()
    out.columns = ["Designation", count_col_name]
    return out


# ======================================================
# FINAL workbook builder (your full layout)
# ======================================================
def export_brand_final_excel(
    brand: str,
    df_main_raw: pd.DataFrame,
    df_piece_raw: pd.DataFrame,
    df_decompte_raw: pd.DataFrame,
) -> bytes:
    # -------- clean ----------
    df_main = clean_main_df(df_main_raw)
    df_piece = clean_piece_df(df_piece_raw)

    # unify headers for export workbook (use Véhicule + Designation consistently)
    rename_map = {
        "véhicule": "Véhicule",
        "Vehicule": "Véhicule",
        "Désignation": "Designation",
        "designation": "Designation",
    }
    df_main = df_main.rename(columns=rename_map)
    df_piece = df_piece.rename(columns=rename_map)

    # safety: fix vehicles again using permissive fix then drop invalid
    veh_main_col = pick_col(df_main, ["Véhicule"])
    veh_piece_col = pick_col(df_piece, ["Véhicule"])
    if veh_main_col:
        df_main[veh_main_col] = df_main[veh_main_col].apply(fix_vehicle)
        df_main = df_main.dropna(subset=[veh_main_col])
    if veh_piece_col:
        df_piece[veh_piece_col] = df_piece[veh_piece_col].apply(fix_vehicle)
        df_piece = df_piece.dropna(subset=[veh_piece_col])

    # add Type + combine
    df_main["Type"] = "Main d'œuvre"
    df_piece["Type"] = "Pièce"
    full_df = pd.concat([df_main, df_piece], ignore_index=True)
    full_df = enforce_schema(full_df)

    # columns
    veh_col = pick_col(full_df, ["Véhicule"])
    total_col = pick_col(full_df, ["Total HTVA", "Total"])
    des_col = pick_col(full_df, ["Designation"])
    qte_col = pick_col(full_df, ["Qté", "Qte", "qte"])
    mnt_col = pick_col(full_df, ["Montant", "montant"])
    doss_col = pick_col(full_df, ["N° Doss", "N°Doss", "No Doss", "N Doss"])
    no_col = pick_col(full_df, ["N°", "No", "N"])

    if veh_col is None:
        raise ValueError("Vehicle column not found after cleaning.")
    if total_col is None:
        raise ValueError("Total column not found after cleaning.")
    if des_col is None:
        des_col = None

    # decompte summary + totals
    df_decompte_sum = build_decompte_summary(df_decompte_raw)

    decompte_total = 0.0
    fgb_global = 0.0
    if not df_decompte_sum.empty:
        mask_total = df_decompte_sum["Decompte"].astype(str).str.upper().str.contains("TOTAL")
        if mask_total.any():
            decompte_total = safe_float(df_decompte_sum.loc[mask_total, "Total HTV/HTVA"].iloc[0])
            fgb_global = safe_float(df_decompte_sum.loc[mask_total, "FGB"].iloc[0])

    # KPI (designation counts)
    df_kpi_main = pd.DataFrame()
    df_kpi_piece = pd.DataFrame()
    if "Designation" in df_main.columns:
        df_kpi_main = build_designation_kpi(df_main, "Designation", "Main Count")
    if "Designation" in df_piece.columns:
        df_kpi_piece = build_designation_kpi(df_piece, "Designation", "Piece Count")

    # -------- workbook ----------
    bio = BytesIO()
    wb = xlsxwriter.Workbook(bio, {"in_memory": True, "nan_inf_to_errors": True})

    header_fmt = wb.add_format({"bold": True, "align": "center", "border": 1, "bg_color": "#2F5597", "font_color": "white"})
    link_fmt = wb.add_format({"font_color": "blue", "underline": 1})
    cell_fmt = wb.add_format({"border": 1})
    int_fmt = wb.add_format({"border": 1, "num_format": "0"})
    qty_fmt = wb.add_format({"border": 1, "num_format": "0.0"})
    money_fmt = wb.add_format({"border": 1, "num_format": "0.000"})
    total_fmt = wb.add_format({"bold": True, "border": 1, "bg_color": "#C6E0B4", "num_format": "0.000"})
    section_title_fmt = wb.add_format({"bold": True, "align": "left", "valign": "vcenter"})
    remarque_fmt = wb.add_format({"bold": True, "align": "left", "valign": "vcenter", "border": 1, "bg_color": "#C6E0B4", "text_wrap": True})

    def get_fmt_for_col(col_name: str):
        c = str(col_name).lower().strip()
        if "doss" in c:
            return int_fmt
        if c in ["n°", "no", "n"]:
            return int_fmt
        if "qt" in c:
            return qty_fmt
        if "montant" in c or "total htva" in c or c == "total":
            return money_fmt
        return cell_fmt

    # Vehicle List sheet
    summary_name = f"{brand} Vehicle List"[:31]
    summary_ws = wb.add_worksheet(summary_name)
    summary_ws.write_row(0, 0, ["Véhicule ID", "Total HTVA", "Open"], header_fmt)

    vehicules = sorted(pd.Series(full_df[veh_col]).dropna().astype(str).unique())
    total_vehicles = 0.0
    missing_rows: List[Dict] = []

    def should_flag_missing(record_dict: dict) -> bool:
        if des_col is None:
            return False
        tot = safe_float(record_dict.get(total_col, 0))
        des = str(record_dict.get(des_col, "")).strip()
        qte = safe_float(record_dict.get(qte_col, 0)) if qte_col else 0.0
        mnt = safe_float(record_dict.get(mnt_col, 0)) if mnt_col else 0.0
        if des == "":
            return False
        if abs(tot) > 1e-12:
            return False
        return (abs(qte) > 1e-12) or (abs(mnt) > 1e-12)

    for i, veh in enumerate(vehicules):
        df_veh = full_df[full_df[veh_col].astype(str) == str(veh)].copy()
        veh_total = round(safe_float(df_veh[total_col].sum()), 3)
        total_vehicles += veh_total

        sheet_name = str(veh)[:31]

        summary_ws.write(i + 1, 0, veh, cell_fmt)
        summary_ws.write_number(i + 1, 1, veh_total, money_fmt)
        summary_ws.write_url(i + 1, 2, f"internal:'{sheet_name}'!A1", link_fmt, string="Go to Sheet")

        ws = wb.add_worksheet(sheet_name)
        ws.write_url(0, 0, f"internal:'{summary_name}'!A1", link_fmt, string="Back To List")

        df_veh["Type"] = df_veh["Type"].astype(str).str.strip()
        df_main_part = df_veh[df_veh["Type"].str.contains("Main", case=False, na=False)].copy()
        df_piece_part = df_veh[df_veh["Type"].str.contains("Pi", case=False, na=False)].copy()

        show_cols = [c for c in df_veh.columns.tolist() if c != "Type"]
        row_cursor = 2

        def write_block(title: str, part: pd.DataFrame):
            nonlocal row_cursor
            if part.empty:
                return
            ws.write(row_cursor, 0, title, section_title_fmt)
            row_cursor += 1
            ws.write_row(row_cursor, 0, show_cols, header_fmt)
            row_cursor += 1

            for rec in part[show_cols].itertuples(index=False, name=None):
                rec_dict = dict(zip(show_cols, rec))
                excel_row_1based = row_cursor + 1

                if should_flag_missing(rec_dict):
                    missing_rows.append({
                        "Véhicule": sheet_name,
                        "Type": title,
                        "Excel Row": excel_row_1based,
                        "N° Doss": rec_dict.get(doss_col, "") if doss_col else "",
                        "N°": rec_dict.get(no_col, "") if no_col else "",
                        "Designation": str(rec_dict.get(des_col, "")).strip() if des_col else "",
                        "Qté": safe_float(rec_dict.get(qte_col, 0)) if qte_col else 0.0,
                        "Montant": safe_float(rec_dict.get(mnt_col, 0)) if mnt_col else 0.0,
                        "Total HTVA": safe_float(rec_dict.get(total_col, 0)),
                    })

                for c, (col_name, val) in enumerate(zip(show_cols, rec)):
                    fmt = get_fmt_for_col(col_name)
                    if pd.isna(val) or val == "":
                        ws.write_blank(row_cursor, c, None, fmt)
                    elif fmt == cell_fmt:
                        ws.write(row_cursor, c, str(val), fmt)
                    else:
                        ws.write_number(row_cursor, c, safe_float(val), fmt)

                row_cursor += 1

            row_cursor += 2

        write_block("Main d'œuvre", df_main_part)
        write_block("Pièce", df_piece_part)

        ws.write(row_cursor, 0, "Total HTVA", section_title_fmt)
        ws.write_number(row_cursor, 1, veh_total, total_fmt)
        ws.set_column(0, len(show_cols) - 1, 18)

    # Summary totals section
    last = len(vehicules) + 2
    summary_ws.write(last, 0, "TOTAL of VEHICLES", header_fmt)
    summary_ws.write_number(last, 1, round(total_vehicles, 3), total_fmt)

    summary_ws.write(last + 1, 0, "Différence (FGB) - GLOBAL", header_fmt)
    summary_ws.write_number(last + 1, 1, round(safe_float(fgb_global), 3), total_fmt)

    summary_ws.write(last + 2, 0, f"TOTAL Décompte {brand}", header_fmt)
    summary_ws.write_number(last + 2, 1, round(safe_float(decompte_total), 3), total_fmt)

    # Links row: Missing HTVA
    missing_sheet = "Missing HTVA"
    missing_row = last + 3
    summary_ws.write(missing_row, 0, "Missing data", header_fmt)
    summary_ws.write_blank(missing_row, 1, None, total_fmt)
    summary_ws.write_url(missing_row, 2, f"internal:'{missing_sheet}'!A1", link_fmt, string="Go to Missing HTVA")

    # KPI links
    design_main_sheet = "Designation Count - Main"
    design_piece_sheet = "Designation Count - Piece"

    summary_ws.write(missing_row + 1, 0, "Designation Count - Main d'œuvre", header_fmt)
    summary_ws.write_blank(missing_row + 1, 1, None, total_fmt)
    summary_ws.write_url(missing_row + 1, 2, f"internal:'{design_main_sheet}'!A1", link_fmt, string="Go to Main d'oeuvre")

    summary_ws.write(missing_row + 2, 0, "Designation Count - Pièce", header_fmt)
    summary_ws.write_blank(missing_row + 2, 1, None, total_fmt)
    summary_ws.write_url(missing_row + 2, 2, f"internal:'{design_piece_sheet}'!A1", link_fmt, string="Go to Pièce")

    # Remarque block
    summary_ws.write(missing_row + 3, 0, "Remarque :", remarque_fmt)
    summary_ws.write(missing_row + 3, 1, f"TOTAL Décompte {brand} = TOTAL des VEHICULES + FGB", remarque_fmt)
    summary_ws.write(missing_row + 4, 0, "Valeur FGB :", remarque_fmt)
    summary_ws.write(missing_row + 4, 1, "Différence entre TOTAL Décompte et TOTAL des VEHICULES", remarque_fmt)

    summary_ws.set_column("A:A", 35)
    summary_ws.set_column("B:B", 35)
    summary_ws.set_column("C:C", 20)

    # Missing HTVA sheet
    miss_ws = wb.add_worksheet(missing_sheet)
    miss_ws.write_url(0, 0, f"internal:'{summary_name}'!A1", link_fmt, string="Back To List")

    miss_headers = ["Véhicule", "Type", "Excel Row", "N° Doss", "N°", "Designation", "Qté", "Montant", "Total HTVA", "Open"]
    miss_ws.write_row(2, 0, miss_headers, header_fmt)

    if missing_rows:
        for r, item in enumerate(missing_rows, start=3):
            miss_ws.write(r, 0, item["Véhicule"], cell_fmt)
            miss_ws.write(r, 1, item["Type"], cell_fmt)
            miss_ws.write_number(r, 2, float(item["Excel Row"]), int_fmt)

            if item["N° Doss"] not in ["", None] and not pd.isna(item["N° Doss"]):
                miss_ws.write_number(r, 3, safe_float(item["N° Doss"]), int_fmt)
            else:
                miss_ws.write(r, 3, "", cell_fmt)

            if item["N°"] not in ["", None] and not pd.isna(item["N°"]):
                miss_ws.write_number(r, 4, safe_float(item["N°"]), int_fmt)
            else:
                miss_ws.write(r, 4, "", cell_fmt)

            miss_ws.write(r, 5, item["Designation"], cell_fmt)
            miss_ws.write_number(r, 6, safe_float(item["Qté"]), qty_fmt)
            miss_ws.write_number(r, 7, safe_float(item["Montant"]), money_fmt)
            miss_ws.write_number(r, 8, safe_float(item["Total HTVA"]), money_fmt)

            target_sheet = item["Véhicule"]
            target_row = int(item["Excel Row"])
            miss_ws.write_url(r, 9, f"internal:'{target_sheet}'!A{target_row}", link_fmt, string="Go to line")
    else:
        miss_ws.write(3, 0, "✅ No lines with Total HTVA = 0.000 found.", cell_fmt)

    miss_ws.set_column("A:A", 14)
    miss_ws.set_column("B:B", 12)
    miss_ws.set_column("C:C", 10)
    miss_ws.set_column("D:E", 10)
    miss_ws.set_column("F:F", 45)
    miss_ws.set_column("G:I", 12)
    miss_ws.set_column("J:J", 12)

    # KPI sheets
    main_kpi_ws = wb.add_worksheet(design_main_sheet[:31])
    main_kpi_ws.write_url(0, 0, f"internal:'{summary_name}'!A1", link_fmt, string="Back To List")
    main_kpi_ws.write_row(2, 0, ["Designation", "Main Count"], header_fmt)
    if not df_kpi_main.empty:
        for r, row in enumerate(df_kpi_main.itertuples(index=False, name=None), start=3):
            main_kpi_ws.write(r, 0, "" if pd.isna(row[0]) else str(row[0]), cell_fmt)
            main_kpi_ws.write_number(r, 1, int(row[1]), int_fmt)
    main_kpi_ws.set_column("A:A", 60)
    main_kpi_ws.set_column("B:B", 18)

    piece_kpi_ws = wb.add_worksheet(design_piece_sheet[:31])
    piece_kpi_ws.write_url(0, 0, f"internal:'{summary_name}'!A1", link_fmt, string="Back To List")
    piece_kpi_ws.write_row(2, 0, ["Designation", "Piece Count"], header_fmt)
    if not df_kpi_piece.empty:
        for r, row in enumerate(df_kpi_piece.itertuples(index=False, name=None), start=3):
            piece_kpi_ws.write(r, 0, "" if pd.isna(row[0]) else str(row[0]), cell_fmt)
            piece_kpi_ws.write_number(r, 1, int(row[1]), int_fmt)
    piece_kpi_ws.set_column("A:A", 60)
    piece_kpi_ws.set_column("B:B", 18)

    # Decompte sheet (summary)
    decompte_name = f"Decompte {brand}"[:31]
    de_ws = wb.add_worksheet(decompte_name)
    if df_decompte_sum is None or df_decompte_sum.empty:
        de_ws.write(0, 0, "No decompte summary", header_fmt)
    else:
        de_ws.write_row(0, 0, df_decompte_sum.columns.tolist(), header_fmt)
        for r, rec in enumerate(df_decompte_sum.itertuples(index=False, name=None), start=1):
            for c, val in enumerate(rec):
                if isinstance(val, (int, np.integer)):
                    de_ws.write_number(r, c, int(val), int_fmt)
                elif isinstance(val, (float, np.floating)):
                    de_ws.write_number(r, c, float(val), money_fmt)
                else:
                    de_ws.write(r, c, "" if val is None else val, cell_fmt)
        de_ws.set_column("A:E", 20)

    # link from summary to decompte
    summary_ws.write_url(last + 2, 2, f"internal:'{decompte_name}'!A1", link_fmt, string="Go to Decompte")

    wb.close()
    return bio.getvalue()