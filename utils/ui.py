import streamlit as st

CSS = """
<style>
/* ====== Layout ====== */
.block-container {
  max-width: 1200px;
  padding-top: 1.2rem;
  padding-bottom: 2rem;
}
div[data-testid="stAppViewContainer"] > .main { padding-top: 0.5rem; }

/* ====== Header banner ====== */
.app-hero {
  background: linear-gradient(90deg, #1F4E79 0%, #2F5597 55%, #3B82F6 100%);
  border-radius: 18px;
  padding: 18px 20px;
  color: #fff;
  margin-bottom: 14px;
  box-shadow: 0 14px 30px rgba(15, 23, 42, 0.14);
  border: 1px solid rgba(255,255,255,0.12);
}
.app-hero-title {
  font-size: 30px;
  font-weight: 900;
  margin: 0;
  letter-spacing: 0.2px;
}
.app-hero-sub {
  margin: 6px 0 0;
  font-size: 14px;
  font-weight: 700;
  opacity: 0.92;
}

/* ====== Section title ====== */
.section-title {
  margin: 16px 0 8px;
  font-size: 15px;
  font-weight: 900;
  color: #0F172A;
}

/* ====== Card container ====== */
.card {
  background: #fff;
  border-radius: 16px;
  padding: 14px 14px;
  border: 1px solid rgba(15, 23, 42, 0.08);
  box-shadow: 0 10px 24px rgba(15, 23, 42, 0.07);
}

/* ====== Buttons ====== */
.stButton > button {
  border-radius: 12px !important;
  padding: 0.55rem 1.05rem !important;
  font-weight: 900 !important;
}

/* ====== Dataframes ====== */
div[data-testid="stDataFrame"] {
  border-radius: 14px;
  overflow: hidden;
  border: 1px solid rgba(15,23,42,.08);
}

/* Alerts */
div[data-testid="stAlert"] { border-radius: 12px; }

/* Sidebar separator */
section[data-testid="stSidebar"] {
  border-right: 1px solid rgba(15,23,42,.08);
}

.muted { color:#64748b; font-size: 12px; }
</style>
"""

def inject_css():
    st.markdown(CSS, unsafe_allow_html=True)

def hero(title: str, subtitle: str = ""):
    st.markdown(
        f"""
        <div class="app-hero">
          <div class="app-hero-title">{title}</div>
          <div class="app-hero-sub">{subtitle}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

def section(title: str):
    st.markdown(f"<div class='section-title'>{title}</div>", unsafe_allow_html=True)

def card_open():
    st.markdown("<div class='card'>", unsafe_allow_html=True)

def card_close():
    st.markdown("</div>", unsafe_allow_html=True)

