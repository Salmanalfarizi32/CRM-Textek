import os
import pandas as pd

# --- Path file relatif terhadap lokasi file ini ---
file_path = os.path.join(os.path.dirname(__file__), "CRM Analyst.xlsx")

# Cek apakah file-nya bener-bener ketemu
if not os.path.exists(file_path):
    st.error(f"File Excel tidak ditemukan di path: {file_path}")
else:
    df = pd.read_excel(file_path, sheet_name="VIP BUYER", engine="openpyxl")
