import pandas as pd
import ezodf
from difflib import SequenceMatcher

def similar(a, b):
    # Normalize strings a bit for better matching
    a = a.replace('', 'O').replace('', 'I').replace('', 'A').replace('', 'E').replace('', 'U')
    return SequenceMatcher(None, a, b).ratio()

df_prof = pd.read_excel('c:/Users/e555044.NTDOM1/Desktop/planilla/Profesionales.XLSX')
data = {}
current_cargo = None
current_prof = None

for index, row in df_prof.iterrows():
    val = str(row['Unnamed: 0']).strip()
    count = row['Unnamed: 1']
    if pd.isna(row['Unnamed: 0']) or val == 'nan' or val == 'Etiquetas de fila' or val == 'Total general':
        continue
    
    if val.startswith('P.') and len(val.split('.')) == 3:
        current_cargo = val
    elif val in ['HOMBRE', 'MUJER']:
        if current_cargo and current_prof:
            if current_cargo not in data:
                data[current_cargo] = {}
            if current_prof not in data[current_cargo]:
                data[current_cargo][current_prof] = {}
            data[current_cargo][current_prof][val] = int(count)
    else:
        current_prof = val

source_profs = set()
for cargo in data:
    for prof in data[cargo]:
        source_profs.add(prof)

doc = ezodf.opendoc('c:/Users/e555044.NTDOM1/Desktop/planilla/FUNCIONARIOS POR GRUPO OCUPACIONAL dic 2025 v1.ods')
sheet = doc.sheets[0]

target_profs = set()
for row in sheet.rows():
    prof = str(row[4].value).strip()
    if prof and not prof.startswith('TOTAL') and not prof.startswith('PROFESIONALES') and prof != 'DENOMINACIÓN' and prof != 'None':
        target_profs.add(prof)

prof_map = {}
for sp in source_profs:
    best_match = None
    best_score = 0
    for tp in target_profs:
        score = similar(sp.upper(), tp.upper())
        if score > best_score:
            best_score = score
            best_match = tp
    prof_map[sp] = best_match

print("Mapping:")
for sp in sorted(source_profs):
    print(f"'{sp}' -> '{prof_map[sp]}'")
