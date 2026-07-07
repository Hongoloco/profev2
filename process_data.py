import pandas as pd
import ezodf
import unicodedata
from difflib import SequenceMatcher

def normalize_text(text):
    if pd.isna(text) or text == 'None':
        return ""
    text = str(text).strip().upper()
    # Remove accents
    text = ''.join(c for c in unicodedata.normalize('NFD', text) if unicodedata.category(c) != 'Mn')
    text = text.replace('.', '')
    return text

def similar(a, b):
    return SequenceMatcher(None, a, b).ratio()

# 1. Read source data
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
            if current_prof not in data:
                data[current_prof] = {}
            if current_cargo not in data[current_prof]:
                data[current_prof][current_cargo] = {}
            data[current_prof][current_cargo][val] = int(count)
    else:
        current_prof = val

# 2. Open ODS file
doc_path = 'c:/Users/e555044.NTDOM1/Desktop/planilla/FUNCIONARIOS POR GRUPO OCUPACIONAL dic 2025 v1.ods'
doc = ezodf.opendoc(doc_path)
sheet = doc.sheets[0]

# 3. Extract target professions and their row indices
target_rows = []
all_rows = list(sheet.rows())
for i, row in enumerate(all_rows):
    prof = str(row[4].value).strip()
    cargo = str(row[5].value).strip()
    
    if cargo.startswith('P.') and len(cargo.split('.')) == 3:
        # If prof is empty, inherit from previous
        if not prof or prof == 'None':
            # look back to find the last valid profession
            for j in range(i-1, -1, -1):
                prev_prof = str(all_rows[j][4].value).strip()
                if prev_prof and prev_prof != 'None':
                    prof = prev_prof
                    break
        target_rows.append({
            'row_idx': i,
            'prof': prof,
            'norm_prof': normalize_text(prof),
            'cargo': cargo
        })

# 4. Map source professions to target professions
source_profs = list(data.keys())
prof_map = {}
for sp in source_profs:
    norm_sp = normalize_text(sp)
    # Custom mapping for known hard cases or abbreviations
    if 'ADMINISTRACION Y CONTABIL' in norm_sp:
        norm_sp = 'TECNOLOGO EN ADMINISTRACION Y CONTABILIDAD'
    if 'DIS DE COM VISUAL/GRAF' in norm_sp:
        norm_sp = 'LICENCIADO EN DISENO DE COMUNICACION'
    if 'RE LAB/GEST HUMANA/RRHH' in norm_sp:
        norm_sp = 'LICENCIADO EN RELACIONES LABORALES'
    if 'CIENCIAS DE LA COMUNIC' in norm_sp:
        norm_sp = 'LIC EN CIENCIAS DE LA COMUNICACION'
    if 'DISEO INDUSTR/APLICADO' in norm_sp or 'DISENO INDUSTR/APLICADO' in norm_sp:
        norm_sp = 'LICENCIADO EN DISENO INDUSTRIAL'
    if 'TEC UNIVERS EN ADMINISTRACION' in norm_sp:
        norm_sp = 'TECNICO UNIVERSITARIO EN ADMINISTRACION'

    best_match = None
    best_score = 0
    # Search against target rows norm_prof
    for tr in target_rows:
        score = similar(norm_sp, tr['norm_prof'])
        if score > best_score:
            best_score = score
            best_match = tr['prof']
            
    if best_score < 0.6:
        print(f"Skipping '{sp}' (best match '{best_match}' score {best_score})")
        prof_map[sp] = None
    else:
        prof_map[sp] = best_match

print("Mapped Professions:")
for sp, tp in prof_map.items():
    print(f"'{sp}' -> '{tp}'")

# 5. Fill the data
# Columns:
# M (Mujeres): index 6 (FUNC.), index 8 (CFP) -> wait, we'll put all in FUNC. (6 and 7)? Let's just put it in FUNC. for now.
# H (Hombres): index 7 (FUNC.)

changes_made = 0
for sp in data:
    tp = prof_map.get(sp)
    if not tp:
        continue
    for cargo in data[sp]:
        # find the target row
        target_row_idx = None
        for tr in target_rows:
            if tr['prof'] == tp and tr['cargo'] == cargo:
                target_row_idx = tr['row_idx']
                break
        
        if target_row_idx is not None:
            # write values
            row_cells = all_rows[target_row_idx]
            mujeres = data[sp][cargo].get('MUJER', 0)
            hombres = data[sp][cargo].get('HOMBRE', 0)
            
            # FUNC. M
            row_cells[6].set_value(mujeres)
            # FUNC. H
            row_cells[7].set_value(hombres)
            changes_made += 1

# Save as a new file
out_path = 'c:/Users/e555044.NTDOM1/Desktop/planilla/FUNCIONARIOS POR GRUPO OCUPACIONAL dic 2025 v1_actualizado.ods'
doc.saveas(out_path)
print(f"Saved to {out_path} with {changes_made} changes.")
