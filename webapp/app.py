import os
from flask import Flask, render_template, request, send_file, flash, redirect, url_for
from werkzeug.utils import secure_filename
import pandas as pd
import ezodf
import unicodedata
from difflib import SequenceMatcher
import tempfile

app = Flask(__name__)
app.secret_key = 'super_secret_key_for_flash_messages'

def normalize_text(text):
    if pd.isna(text) or text == 'None':
        return ""
    text = str(text).strip().upper()
    text = ''.join(c for c in unicodedata.normalize('NFD', text) if unicodedata.category(c) != 'Mn')
    text = text.replace('.', '')
    return text

def similar(a, b):
    return SequenceMatcher(None, a, b).ratio()

def process_files(source_path, target_path, output_path):
    # 1. Read source data
    df_prof = pd.read_excel(source_path)
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
                    row_cells[6].set_value(None)
                    
                if hombres > 0:
                    row_cells[7].set_value(hombres)
                else:
                    row_cells[7].set_value(None)
                changes_made += 1

    # 6. Save as a new file
    doc.saveas(output_path)
    return changes_made

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # check if the post request has the file part
        if 'source_file' not in request.files or 'target_file' not in request.files:
            flash('Faltan archivos', 'error')
            return redirect(request.url)
            
        source_file = request.files['source_file']
        target_file = request.files['target_file']
        
        if source_file.filename == '' or target_file.filename == '':
            flash('No seleccionaste ningún archivo', 'error')
            return redirect(request.url)
            
        if source_file and target_file:
            with tempfile.TemporaryDirectory() as temp_dir:
                source_path = os.path.join(temp_dir, secure_filename(source_file.filename))
                target_path = os.path.join(temp_dir, secure_filename(target_file.filename))
                
                # Keep original extension but add _actualizado
                base, ext = os.path.splitext(target_file.filename)
                output_filename = f"{base}_actualizado{ext}"
                output_path = os.path.join(temp_dir, output_filename)
                
                source_file.save(source_path)
                target_file.save(target_path)
                
                try:
                    changes_made = process_files(source_path, target_path, output_path)
                    # We need to send the file back to the user
                    return send_file(output_path, as_attachment=True, download_name=output_filename)
                except Exception as e:
                    flash(f'Ocurrió un error: {str(e)}', 'error')
                    return redirect(request.url)
                    
    return render_template('index.html')

if __name__ == '__main__':
    # Start web browser automatically
    import threading
    import webbrowser
    def open_browser():
        webbrowser.open_new("http://127.0.0.1:5000/")
    threading.Timer(1, open_browser).start()
    app.run(debug=True, use_reloader=False)
