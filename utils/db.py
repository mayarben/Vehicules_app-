# db.py
import streamlit as st
from sqlalchemy import create_engine
from sqlalchemy.engine import Engine

@st.cache_resource
def get_engine() -> Engine:
    cfg = st.secrets["db"]
    if cfg.get("type") != "sqlite":
        raise ValueError("Only sqlite is configured right now.")
    return create_engine(f"sqlite:///{cfg['path']}", future=True)