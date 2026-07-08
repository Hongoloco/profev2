import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import ezodf
import unicodedata
from difflib import SequenceMatcher
import os

def normalize_text(text):
    if pd.isna(text) or text == 'None':
        return ""
    text = str(text).strip().upper()
    text = ''.join(c for c in unicodedata.normalize('NFD', text) if unicodedata.category(c) != 'Mn')
    text = text.replace('.', '')
    return text

def similar(a, b):
    return SequenceMatcher(None, a, b).ratio()

def detect_employment_status(value):
    norm = normalize_text(value)
    if not norm:
        return None
    if 'PROVIS' in norm:
        return 'PROVISORIO'
    if 'PERMAN' in norm:
        return 'PERMANENTE'
    return None

def process_files(source_path, target_path):
    try:
        # 1. Read source data
        df_prof = pd.read_excel(source_path)
        data = {}
        current_cargo = None
        current_prof = None
        current_status = 'PERMANENTE'

        for index, row in df_prof.iterrows():
            val = str(row['Unnamed: 0']).strip()
            count = row['Unnamed: 1']
            if pd.isna(row['Unnamed: 0']) or val == 'nan' or val == 'Etiquetas de fila' or val == 'Total general':
                continue

            status = detect_employment_status(val)
            if status:
                current_status = status
                continue
            
            if val.startswith('P.') and len(val.split('.')) == 3:
                current_cargo = val
            elif val in ['HOMBRE', 'MUJER']:
                if current_cargo and current_prof:
                    if current_prof not in data:
                        data[current_prof] = {}
                    if current_cargo not in data[current_prof]:
                        data[current_prof][current_cargo] = {}
                    data[current_prof][current_cargo].setdefault(val, 0)
                    # Regla de negocio: provisorio se suma al mismo total que permanente.
                    data[current_prof][current_cargo][val] += int(count)
            else:
                current_prof = val

        # 2. Open ODS file
        doc = ezodf.opendoc(target_path)
        sheet = doc.sheets[0]

        # 3. Extract target professions and their row indices
        target_rows = []
        all_rows = list(sheet.rows())
        for i, row in enumerate(all_rows):
            prof = str(row[4].value).strip()
            cargo = str(row[5].value).strip()
            
            if cargo.startswith('P.') and len(cargo.split('.')) == 3:
                if not prof or prof == 'None':
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
            for tr in target_rows:
                score = similar(norm_sp, tr['norm_prof'])
                if score > best_score:
                    best_score = score
                    best_match = tr['prof']
                    
            if best_score < 0.6:
                prof_map[sp] = None
            else:
                prof_map[sp] = best_match

        # 5. Fill the data
        changes_made = 0
        for sp in data:
            tp = prof_map.get(sp)
            if not tp:
                continue
            for cargo in data[sp]:
                target_row_idx = None
                for tr in target_rows:
                    if tr['prof'] == tp and tr['cargo'] == cargo:
                        target_row_idx = tr['row_idx']
                        break
                
                if target_row_idx is not None:
                    row_cells = all_rows[target_row_idx]
                    mujeres = data[sp][cargo].get('MUJER', 0)
                    hombres = data[sp][cargo].get('HOMBRE', 0)
                    
                    if mujeres > 0:
                        row_cells[6].set_value(mujeres)
                    else:
                        row_cells[6].set_value('')
                        
                    if hombres > 0:
                        row_cells[7].set_value(hombres)
                    else:
                        row_cells[7].set_value('')
                    changes_made += 1

        # 6. Save as a new file
        base, ext = os.path.splitext(target_path)
        out_path = f"{base}_actualizado{ext}"
        doc.saveas(out_path)
        
        return True, f"Se actualizó exitosamente y se guardó como:\n{out_path}\n\nCambios realizados: {changes_made}"
    except Exception as e:
        return False, str(e)

def run_app():
    root = tk.Tk()
    root.withdraw() # Hide the main window

    messagebox.showinfo("Actualizador de Planillas", "Paso 1: Selecciona el archivo de 'Profesionales' (Excel)")
    source_path = filedialog.askopenfilename(
        title="Seleccionar archivo de Profesionales (Excel)",
        filetypes=[("Archivos Excel", "*.xlsx *.xls")]
    )
    
    if not source_path:
        messagebox.showwarning("Cancelado", "No se seleccionó el archivo de origen.")
        return

    messagebox.showinfo("Actualizador de Planillas", "Paso 2: Selecciona la planilla de 'Funcionarios por Grupo Ocupacional' (.ods)")
    target_path = filedialog.askopenfilename(
        title="Seleccionar planilla de Funcionarios (ODS)",
        filetypes=[("Archivos ODS", "*.ods")]
    )

    if not target_path:
        messagebox.showwarning("Cancelado", "No se seleccionó el archivo destino.")
        return

    success, msg = process_files(source_path, target_path)
    
    if success:
        messagebox.showinfo("Éxito", msg)
    else:
        messagebox.showerror("Error", f"Ocurrió un error al procesar los archivos:\n{msg}")

if __name__ == '__main__':
    run_app()
