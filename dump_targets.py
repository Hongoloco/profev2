import pandas as pd
import ezodf
doc = ezodf.opendoc('c:/Users/e555044.NTDOM1/Desktop/planilla/FUNCIONARIOS POR GRUPO OCUPACIONAL dic 2025 v1.ods')
sheet = doc.sheets[0]
for i, row in enumerate(list(sheet.rows())):
    prof = str(row[4].value).strip()
    cargo = str(row[5].value).strip()
    if cargo.startswith('P.'):
        print(f"Row {i}: prof='{prof}', cargo='{cargo}'")
