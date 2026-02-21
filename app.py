import streamlit as st

st.set_page_config(
    page_title="Repair App",
    page_icon="🛠️",
    layout="wide",
)

st.title("Repair App")

st.markdown(
    """
Use the sidebar pages:

1) **Cleaning**: upload brand files and generate cleaned outputs  
2) **Results**: download brand cleaned files + global merged workbook  
3) **Dashboard**: KPIs and charts
"""
)