
# init_db.py

from sqlalchemy import text
from utils.db import get_engine

def init_db():
    schema = """
    PRAGMA foreign_keys = ON;

    CREATE TABLE IF NOT EXISTS app_run (
      id TEXT PRIMARY KEY,
      created_at TEXT NOT NULL,
      run_type TEXT NOT NULL,
      run_name TEXT,
      status TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS app_artifact (
      id TEXT PRIMARY KEY,
      run_id TEXT NOT NULL,
      scope TEXT NOT NULL,
      name TEXT NOT NULL,
      content_type TEXT NOT NULL,
      sha256 TEXT,
      byte_size INTEGER,
      data BLOB,
      created_at TEXT NOT NULL,
      FOREIGN KEY(run_id) REFERENCES app_run(id) ON DELETE CASCADE
    );

    CREATE UNIQUE INDEX IF NOT EXISTS uq_artifact ON app_artifact(run_id, scope, name);
    CREATE INDEX IF NOT EXISTS idx_artifact_run ON app_artifact(run_id);
    """
    eng = get_engine()
    with eng.begin() as con:
        for stmt in schema.split(";"):
            s = stmt.strip()
            if s:
                con.execute(text(s))