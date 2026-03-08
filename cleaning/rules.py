#rules.py

import numpy as np
import pandas as pd

def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df

def safe_float(x):
    try:
        if x is None:
            return 0.0
        v = float(x)
        if np.isnan(v) or np.isinf(v):
            return 0.0
        return v
    except Exception:
        return 0.0

def find_col(df: pd.DataFrame, candidates):
    cols = {str(c).lower().strip(): c for c in df.columns}
    for cand in candidates:
        k = cand.lower().strip()
        if k in cols:
            return cols[k]

    # contains match
    for c in df.columns:
        lc = str(c).lower()
        if any(k.lower() in lc for k in candidates):
            return c
    return None

def count_designations(df: pd.DataFrame) -> int:
    col = find_col(df, ["designation", "désignation", "libelle", "libellé"])
    if not col:
        return 0
    s = df[col].astype(str).str.strip()
    s = s[(s != "") & (s.str.lower() != "nan")]
    return int(s.nunique())
