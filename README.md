# Vehicules_app

Streamlit multipage application to automate vehicle maintenance Excel processing
(cleaning, validation, per-brand outputs, and a consolidated dataset).

## Prerequisites
- Python 3.10+ (recommended)
- pip

## Install
pip install -r requirements.txt

## Run
streamlit run app.py

## Inputs
The app expects 9 Excel files (3 brands × 3 datasets):
- Main d'œuvre (Labor)
- Pièces (Parts)
- Décompte (Financial summary)

## Outputs
After processing, the app generates:
- Per-brand cleaned/audit workbooks (e.g., TAS_cleaned.xlsx, Peugeot_cleaned.xlsx, Citroen_cleaned.xlsx)
- A consolidated workbook: Dataset_Complet.xlsx
- Optional anomaly sheets (e.g., Missing HTVA) depending on the data

## Notes
- Run commands from the project root (where `requirements.txt` and `app.py` exist).
- If you get a dependency error, recreate a fresh virtual environment and reinstall requirements.
