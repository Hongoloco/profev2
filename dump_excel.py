import pandas as pd
pd.set_option('display.max_rows', 100)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', 1000)

print("--- Profesionales.XLSX ---")
df_prof = pd.read_excel('c:/Users/e555044.NTDOM1/Desktop/planilla/Profesionales.XLSX')
print(df_prof.head(50))

print("\n--- FUNCIONARIOS POR GRUPO OCUPACIONAL dic 2025 v1.ods ---")
df_func = pd.read_excel('c:/Users/e555044.NTDOM1/Desktop/planilla/FUNCIONARIOS POR GRUPO OCUPACIONAL dic 2025 v1.ods', engine='odf')
df_func_dropna = df_func.dropna(how='all')
print(df_func_dropna.head(50))
