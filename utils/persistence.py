from __future__ import annotations

import sqlite3
from pathlib import Path
from datetime import datetime
from io import BytesIO
from typing import Optional, List, Dict, Any

import pandas as pd
import streamlit as st

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

APP_DIR = Path(__file__).resolve().parent.parent


def _resolve_db_path() -> Path:
    """
    Single source of truth: secrets.toml -> [db].path
    Fallback: data/app.sqlite
    """
    try:
        cfg = st.secrets.get("db", {})
        p = cfg.get("path")
        if p:
            return (APP_DIR / p).resolve()
    except Exception:
        pass
    return (APP_DIR / "data" / "app.sqlite").resolve()


DB_PATH = _resolve_db_path()
DB_PATH.parent.mkdir(parents=True, exist_ok=True)


def _conn() -> sqlite3.Connection:
    return sqlite3.connect(str(DB_PATH), check_same_thread=False)


def _has_column(conn: sqlite3.Connection, table: str, column: str) -> bool:
    cur = conn.cursor()
    cur.execute(f"PRAGMA table_info({table});")
    cols = [r[1] for r in cur.fetchall()]  # r[1] is column name
    return column in cols


def _init_db() -> None:
    with _conn() as conn:
        cur = conn.cursor()

        cur.execute("""
        CREATE TABLE IF NOT EXISTS runs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            scope TEXT,
            run_name TEXT,
            created_at TEXT
        );
        """)

        cur.execute("""
        CREATE TABLE IF NOT EXISTS artifacts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            run_id INTEGER,
            scope TEXT,
            name TEXT,
            content_type TEXT,
            data BLOB,
            created_at TEXT
        );
        """)

        cur.execute("""
        CREATE TABLE IF NOT EXISTS dfs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            run_id INTEGER,
            scope TEXT,
            name TEXT,
            data BLOB,
            created_at TEXT
        );
        """)

        cur.execute("CREATE INDEX IF NOT EXISTS idx_runs_created_at ON runs(created_at);")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_artifacts_run ON artifacts(run_id);")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_dfs_run ON dfs(run_id);")

        conn.commit()

        # âœ… Migration safety: if an older DB exists without run_name, add it.
        if not _has_column(conn, "runs", "run_name"):
            cur.execute("ALTER TABLE runs ADD COLUMN run_name TEXT;")
            conn.commit()


_init_db()

# ======================================================
# Runs
# ======================================================
def create_run(scope: str, run_name: str = "") -> int:
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    run_name = (run_name or "").strip()
    with _conn() as conn:
        cur = conn.cursor()
        cur.execute(
            "INSERT INTO runs(scope, run_name, created_at) VALUES(?,?,?)",
            (scope, run_name, now),
        )
        conn.commit()
        return int(cur.lastrowid)


def rename_run(run_id: int, run_name: str) -> None:
    """
    Update the display name of a session/run.
    Pass empty string to clear the name.
    """
    run_name = (run_name or "").strip()
    # Optional length cap to keep UI tidy
    if len(run_name) > 80:
        run_name = run_name[:80]

    with _conn() as conn:
        cur = conn.cursor()
        cur.execute("UPDATE runs SET run_name=? WHERE id=?", (run_name, run_id))
        conn.commit()


def get_run(run_id: int) -> Optional[Dict[str, Any]]:
    with _conn() as conn:
        cur = conn.cursor()
        cur.execute(
            "SELECT id, scope, run_name, created_at FROM runs WHERE id=?",
            (run_id,),
        )
        row = cur.fetchone()
    if not row:
        return None
    rid, scope, run_name, created_at = row
    return {"id": rid, "scope": scope, "run_name": run_name, "created_at": created_at}


def list_runs(limit: int = 200) -> List[Dict[str, Any]]:
    with _conn() as conn:
        cur = conn.cursor()
        cur.execute(
            "SELECT id, scope, run_name, created_at FROM runs ORDER BY id DESC LIMIT ?",
            (limit,),
        )
        rows = cur.fetchall()

    return [
        {"id": rid, "scope": scope, "run_name": run_name or "", "created_at": created_at}
        for (rid, scope, run_name, created_at) in rows
    ]


# ======================================================
# Artifacts
# ======================================================
def save_artifact(run_id: int, scope: str, name: str, content_type: str, data: bytes) -> None:
    """
    âœ… Overwrite behavior:
    - If an artifact with same (run_id, scope, name) exists, delete it first.
    - Then insert the new one.
    """
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with _conn() as conn:
        cur = conn.cursor()

        # âœ… overwrite old artifact with same name
        cur.execute(
            "DELETE FROM artifacts WHERE run_id=? AND scope=? AND name=?",
            (run_id, scope, name),
        )

        cur.execute(
            """
            INSERT INTO artifacts(run_id, scope, name, content_type, data, created_at)
            VALUES(?,?,?,?,?,?)
            """,
            (run_id, scope, name, content_type, sqlite3.Binary(data), now),
        )
        conn.commit()


def list_artifacts(run_id: int) -> List[Dict[str, Any]]:
    with _conn() as conn:
        cur = conn.cursor()
        cur.execute(
            """
            SELECT id, scope, name, content_type, created_at
            FROM artifacts
            WHERE run_id=?
            ORDER BY id ASC
            """,
            (run_id,),
        )
        rows = cur.fetchall()

    return [
        {"id": i, "scope": s, "name": n, "content_type": ct, "created_at": ca}
        for (i, s, n, ct, ca) in rows
    ]


def get_artifact_bytes(artifact_id: int) -> Optional[bytes]:
    with _conn() as conn:
        cur = conn.cursor()
        cur.execute("SELECT data FROM artifacts WHERE id=?", (artifact_id,))
        row = cur.fetchone()
    return None if not row else row[0]


# ======================================================
# DataFrames
# ======================================================
def save_df_parquet(run_id: int, scope: str, name: str, df: pd.DataFrame) -> None:
    bio = BytesIO()
    payload: bytes

    try:
        df.to_parquet(bio, index=False)
        payload = bio.getvalue()
    except Exception:
        bio = BytesIO()
        df.to_pickle(bio)
        payload = bio.getvalue()

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with _conn() as conn:
        cur = conn.cursor()
        cur.execute(
            """
            INSERT INTO dfs(run_id, scope, name, data, created_at)
            VALUES(?,?,?,?,?)
            """,
            (run_id, scope, name, sqlite3.Binary(payload), now),
        )
        conn.commit()


def load_df(run_id: int, scope: str, name: str) -> Optional[pd.DataFrame]:
    with _conn() as conn:
        cur = conn.cursor()
        cur.execute(
            """
            SELECT data
            FROM dfs
            WHERE run_id=? AND scope=? AND name=?
            ORDER BY id DESC
            LIMIT 1
            """,
            (run_id, scope, name),
        )
        row = cur.fetchone()

    if not row:
        return None

    b = row[0]
    bio = BytesIO(b)
    try:
        return pd.read_parquet(bio)
    except Exception:
        bio = BytesIO(b)
        try:
            return pd.read_pickle(bio)
        except Exception:
            return None


# ======================================================
# Delete session
# ======================================================
def delete_run(run_id: int) -> None:
    with _conn() as conn:
        cur = conn.cursor()
        cur.execute("DELETE FROM artifacts WHERE run_id=?", (run_id,))
        cur.execute("DELETE FROM dfs WHERE run_id=?", (run_id,))
        cur.execute("DELETE FROM runs WHERE id=?", (run_id,))
        conn.commit()
