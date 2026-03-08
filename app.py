import streamlit as st

st.set_page_config(page_title="Vehicle Maintenance Automation", layout="wide")

st.title("Vehicle Maintenance — Automation App")
st.write("Use the sidebar pages: Sessions → Cleaning → Results of cleaning → Vehicle Dates Extraction → Dashboards.")

st.markdown(
    """
Use the sidebar pages:
1) **Sessions**: Select a saved session, restore it, and download saved outputs (cleaned files, global merge, vehicle dates) + open dashboards.
2) **Cleaning**: Upload TAS / Peugeot / Citroen files (Main d'œuvre + Pièces + Décompte) and generate cleaned outputs.
3) **Results of cleaning** : Download the cleaned brand workbooks, preview cleaned tables, and build the global merged workbook (Dataset_Complet.xlsx).
4) **Vehicle Dates Extraction**: Extract/prepare vehicle-related date fields for analysis and reporting.
5) **Dashboards**: KPIs and charts.
"""
)





