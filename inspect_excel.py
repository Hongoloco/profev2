import pandas as pd

try:
    print("--- Profesionales.XLSX ---")
    df_prof = pd.read_excel('c:/Users/e555044.NTDOM1/Desktop/planilla/Profesionales.XLSX')
    print("Columns:", df_prof.columns.tolist())
    print(df_prof.head())
except Exception as e:
    print("Error reading Profesionales:", e)

try:
    print("\n--- FUNCIONARIOS POR GRUPO OCUPACIONAL dic 2025 v1.ods ---")
    df_func = pd.read_excel('c:/Users/e555044.NTDOM1/Desktop/planilla/FUNCIONARIOS POR GRUPO OCUPACIONAL dic 2025 v1.ods', engine='odf')
    print("Columns:", df_func.columns.tolist())
    print(df_func.head())
except Exception as e:
    print("Error reading FUNCIONARIOS:", e)
