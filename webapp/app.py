import os
from flask import Flask, render_template, request, send_file, flash, redirect, url_for
from werkzeug.utils import secure_filename
import pandas as pd
import ezodf
import unicodedata
from difflib import SequenceMatcher
import tempfile
import re
import io
import shutil
import time
from uuid import uuid4
from collections import defaultdict

TABLE_FORMULA_ATTR = '{urn:oasis:names:tc:opendocument:xmlns:table:1.0}formula'

app = Flask(__name__)
app.secret_key = 'super_secret_key_for_flash_messages'

CARGO_PATTERN = re.compile(r'^[A-Z]\.\d{2}\.[A-Z]$')
TOTAL_GROUP_PATTERN = re.compile(r'TOTAL\s+GRUPO\s+([A-Z])')
CHART_GROUP_PATTERN = re.compile(r'^\s*GRUPO\s+([A-Z])\b')
DOWNLOAD_CACHE = {}
CACHE_TTL_SECONDS = 20 * 60

# Abreviaturas y nombres alternativos frecuentes del origen.
PROF_ALIASES = {
    'ADMINISTRACION Y CONTABIL': 'TECNOLOGO EN ADMINISTRACION Y CONTABILIDAD',
    'DIS DE COM VISUAL/GRAF': 'LICENCIADO EN DISENO DE COMUNICACION',
    'RE LAB/GEST HUMANA/RRHH': 'LICENCIADO EN RELACIONES LABORALES',
    'CIENCIAS DE LA COMUNIC': 'LIC EN CIENCIAS DE LA COMUNICACION',
    'DISEO INDUSTR/APLICADO': 'LICENCIADO EN DISENO INDUSTRIAL',
    'DISENO INDUSTR/APLICADO': 'LICENCIADO EN DISENO INDUSTRIAL',
    'TEC UNIVERS EN ADMINISTRACION': 'TECNICO UNIVERSITARIO EN ADMINISTRACION',
}

def normalize_text(text):
    if pd.isna(text) or text == 'None':
        return ""
    text = str(text).strip().upper()
    text = ''.join(c for c in unicodedata.normalize('NFD', text) if unicodedata.category(c) != 'Mn')
    text = text.replace('.', '')
    return text

def similar(a, b):
    return SequenceMatcher(None, a, b).ratio()

def is_cargo_code(value):
    if not value:
        return False
    return bool(CARGO_PATTERN.match(str(value).strip().upper()))

def source_prof_key(value):
    norm = normalize_text(value)
    for alias_key, alias_target in PROF_ALIASES.items():
        if alias_key in norm:
            return normalize_text(alias_target)
    return norm

def detect_employment_status(value):
    norm = normalize_text(value)
    if not norm:
        return None
    if 'PROVIS' in norm:
        return 'PROVISORIO'
    if 'PERMAN' in norm:
        return 'PERMANENTE'
    return None

def parse_source(source_path):
    df_prof = pd.read_excel(source_path)
    if len(df_prof.columns) < 2:
        raise ValueError('La planilla de Profesionales no tiene el formato esperado (faltan columnas).')

    first_col = df_prof.columns[0]
    second_col = df_prof.columns[1]

    data = {}
    current_cargo = None
    current_prof = None
    current_status = 'PERMANENTE'
    raw_status_totals = {
        'PERMANENTE': {'HOMBRE': 0, 'MUJER': 0},
        'PROVISORIO': {'HOMBRE': 0, 'MUJER': 0},
    }
    effective_status_totals = {
        'PERMANENTE': {'HOMBRE': 0, 'MUJER': 0},
    }

    for _, row in df_prof.iterrows():
        raw_val = row[first_col]
        count = row[second_col]
        val = '' if pd.isna(raw_val) else str(raw_val).strip()
        upper_val = val.upper()

        if not val or upper_val in {'ETIQUETAS DE FILA', 'TOTAL GENERAL', 'NAN'}:
            continue

        status = detect_employment_status(upper_val)
        if status:
            current_status = status
            continue

        if is_cargo_code(upper_val):
            current_cargo = upper_val
            continue

        if upper_val in {'HOMBRE', 'MUJER'}:
            if current_cargo and current_prof and not pd.isna(count):
                count_value = int(count)
                data.setdefault(current_prof, {}).setdefault(current_cargo, {}).setdefault(upper_val, 0)
                data[current_prof][current_cargo][upper_val] += count_value

                raw_status_totals.setdefault(current_status, {'HOMBRE': 0, 'MUJER': 0})
                raw_status_totals[current_status][upper_val] += count_value

                # Regla de negocio: provisorio computa como permanente.
                effective_status = 'PERMANENTE' if current_status == 'PROVISORIO' else current_status
                effective_status_totals.setdefault(effective_status, {'HOMBRE': 0, 'MUJER': 0})
                effective_status_totals[effective_status][upper_val] += count_value
            continue

        current_prof = val

    return data, {
        'raw_status_totals': raw_status_totals,
        'effective_status_totals': effective_status_totals,
    }

def parse_target(target_path):
    doc = ezodf.opendoc(target_path)
    sheet = doc.sheets[0]
    all_rows = list(sheet.rows())

    target_rows = []
    for i, row in enumerate(all_rows):
        prof = str(row[4].value).strip()
        cargo = str(row[5].value).strip().upper()

        if is_cargo_code(cargo):
            if not prof or prof == 'None':
                for j in range(i - 1, -1, -1):
                    prev_prof = str(all_rows[j][4].value).strip()
                    if prev_prof and prev_prof != 'None':
                        prof = prev_prof
                        break

            target_rows.append({
                'row_idx': i,
                'prof': prof,
                'norm_prof': normalize_text(prof),
                'cargo': cargo,
            })

    return doc, all_rows, target_rows

def build_profession_totals_from_target_rows(all_rows, target_rows):
    # Agrupa solo grupo P usando lo efectivamente escrito en la hoja principal.
    by_prof_level = defaultdict(lambda: {'A': 0, 'B': 0})
    by_level_series = defaultdict(lambda: defaultdict(int))

    for tr in target_rows:
        cargo = tr['cargo']
        parts = str(cargo).split('.')
        if len(parts) != 3:
            continue

        group_letter, serie, level = parts
        if group_letter != 'P':
            continue

        row_cells = all_rows[tr['row_idx']]
        total = parse_cell_int(row_cells[6].value) + parse_cell_int(row_cells[7].value) + parse_cell_int(row_cells[8].value) + parse_cell_int(row_cells[9].value)
        if total <= 0:
            continue

        target_prof = tr['prof']
        if level in ('A', 'B'):
            by_prof_level[target_prof][level] += total

        if level in ('A', 'B', 'C', 'D', 'E') and serie in ('01', '02', '03'):
            by_level_series[level][serie] += total

    return by_prof_level, by_level_series

def fill_sheet_prof_ab(doc, by_prof_level):
    if len(doc.sheets) < 2:
        return

    sheet = doc.sheets[1]
    rows = list(sheet.rows())

    total_a = 0
    total_b = 0

    for row_cells in rows:
        prof_name = row_cells[3].value
        if not isinstance(prof_name, str):
            continue

        prof_name_clean = prof_name.strip()
        if not prof_name_clean:
            continue

        if normalize_text(prof_name_clean) == 'TOTAL':
            continue

        norm_prof = normalize_text(prof_name_clean)
        a_value = 0
        b_value = 0
        for target_prof, totals in by_prof_level.items():
            if normalize_text(target_prof) == norm_prof:
                a_value = totals.get('A', 0)
                b_value = totals.get('B', 0)
                break

        set_numeric_or_blank_no_formula(row_cells[4], a_value)
        set_numeric_or_blank_no_formula(row_cells[5], b_value)
        set_numeric_or_blank_no_formula(row_cells[6], a_value + b_value)

        total_a += a_value
        total_b += b_value

    for row_cells in rows:
        label = row_cells[2].value
        if isinstance(label, str) and normalize_text(label) == 'TOTAL':
            set_numeric_or_blank_no_formula(row_cells[4], total_a)
            set_numeric_or_blank_no_formula(row_cells[5], total_b)
            set_numeric_or_blank_no_formula(row_cells[6], total_a + total_b)
            break

def fill_sheet_level_series(doc, by_level_series):
    if len(doc.sheets) < 3:
        return

    sheet = doc.sheets[2]
    rows = list(sheet.rows())
    levels = ['A', 'B', 'C', 'D', 'E']

    level_row_map = {}
    for row_cells in rows:
        label = row_cells[2].value
        if not isinstance(label, str):
            continue

        label_norm = normalize_text(label)
        for level in levels:
            marker = f'PROFESIONALES {level}'
            if marker in label_norm:
                level_row_map[level] = row_cells
                break

    total_01 = 0
    total_02 = 0
    total_03 = 0

    for level in levels:
        row_cells = level_row_map.get(level)
        if not row_cells:
            continue

        v01 = by_level_series[level].get('01', 0)
        v02 = by_level_series[level].get('02', 0)
        v03 = by_level_series[level].get('03', 0)
        row_total = v01 + v02 + v03

        set_numeric_or_blank_no_formula(row_cells[3], v01)
        set_numeric_or_blank_no_formula(row_cells[4], v02)
        set_numeric_or_blank_no_formula(row_cells[5], v03)
        set_numeric_or_blank_no_formula(row_cells[6], row_total)

        total_01 += v01
        total_02 += v02
        total_03 += v03

    for row_cells in rows:
        label = row_cells[2].value
        if isinstance(label, str) and normalize_text(label) == 'TOTAL':
            set_numeric_or_blank_no_formula(row_cells[3], total_01)
            set_numeric_or_blank_no_formula(row_cells[4], total_02)
            set_numeric_or_blank_no_formula(row_cells[5], total_03)
            set_numeric_or_blank_no_formula(row_cells[6], total_01 + total_02 + total_03)
            break

def build_profession_map(source_data, target_rows):
    source_profs = list(source_data.keys())
    target_norm_to_prof = {}
    for tr in target_rows:
        target_norm_to_prof[tr['norm_prof']] = tr['prof']
    unique_target_norms = list(target_norm_to_prof.keys())

    prof_map = {}
    unmapped = []

    for sp in source_profs:
        norm_sp = source_prof_key(sp)

        # Solo se permite match exacto normalizado (incluyendo alias conocidos).
        # Evitamos mapeos difusos para no cargar cargos/perfiles que no existan realmente en fuente.
        if norm_sp in target_norm_to_prof:
            prof_map[sp] = target_norm_to_prof[norm_sp]
            continue

        best_match = None
        best_score = 0
        for norm_target in unique_target_norms:
            score = similar(norm_sp, norm_target)
            if score > best_score:
                best_score = score
                best_match = target_norm_to_prof[norm_target]

        prof_map[sp] = None
        unmapped.append({'source_prof': sp, 'best_match': best_match, 'score': round(best_score, 3)})

    return prof_map, unmapped

def set_numeric_or_blank(cell, value):
    if value and int(value) > 0:
        cell.set_value(int(value))
    else:
        cell.set_value('')

def parse_cell_int(value):
    if value in (None, ''):
        return 0
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return 0

def set_numeric_value(cell, value):
    cell.xmlnode.attrib.pop(TABLE_FORMULA_ATTR, None)
    if isinstance(value, float):
        cell.set_value(float(value))
    else:
        cell.set_value(int(value))

def set_numeric_or_blank_no_formula(cell, value):
    cell.xmlnode.attrib.pop(TABLE_FORMULA_ATTR, None)
    if value and float(value) > 0:
        if isinstance(value, float) and not float(value).is_integer():
            cell.set_value(float(value))
        else:
            cell.set_value(int(value))
    else:
        cell.set_value('')

def cleanup_temp_dir(path):
    for _ in range(5):
        try:
            shutil.rmtree(path)
            return
        except PermissionError:
            time.sleep(0.2)
        except FileNotFoundError:
            return

    shutil.rmtree(path, ignore_errors=True)

def clean_broken_ref_formulas(doc):
    for sheet in doc.sheets:
        all_rows = list(sheet.rows())
        for row in all_rows:
            for cell in row:
                formula = getattr(cell, 'formula', None)
                if formula and '#REF' in str(formula).upper():
                    # Eliminamos formulas rotas para que no aparezcan #REF! tras exportar.
                    cell.xmlnode.attrib.pop(TABLE_FORMULA_ATTR, None)
                    cell.set_value('')

def reset_target_data_columns(all_rows, target_rows):
    for tr in target_rows:
        row_cells = all_rows[tr['row_idx']]
        row_cells[6].set_value('')
        row_cells[7].set_value('')
        row_cells[8].set_value('')
        row_cells[9].set_value('')

def rebuild_totals_and_chart_data(all_rows, target_rows):
    group_totals = defaultdict(lambda: {'M': 0, 'H': 0, 'T': 0})

    for tr in target_rows:
        row_cells = all_rows[tr['row_idx']]
        total_m = parse_cell_int(row_cells[6].value) + parse_cell_int(row_cells[8].value)
        total_h = parse_cell_int(row_cells[7].value) + parse_cell_int(row_cells[9].value)
        total_t = total_m + total_h

        set_numeric_or_blank_no_formula(row_cells[11], total_m)
        set_numeric_or_blank_no_formula(row_cells[12], total_h)
        set_numeric_or_blank_no_formula(row_cells[13], total_t)

        group_letter = tr['cargo'][:1]
        if group_letter:
            group_totals[group_letter]['M'] += total_m
            group_totals[group_letter]['H'] += total_h
            group_totals[group_letter]['T'] += total_t

    for row_cells in all_rows:
        label = row_cells[4].value
        if isinstance(label, str):
            match = TOTAL_GROUP_PATTERN.search(label.upper())
            if match:
                letter = match.group(1)
                totals = group_totals.get(letter, {'M': 0, 'H': 0, 'T': 0})
                set_numeric_or_blank_no_formula(row_cells[11], totals['M'])
                set_numeric_or_blank_no_formula(row_cells[12], totals['H'])
                set_numeric_or_blank_no_formula(row_cells[13], totals['T'])

    chart_rows = []
    for row_cells in all_rows:
        label = row_cells[15].value
        if isinstance(label, str):
            match = CHART_GROUP_PATTERN.search(label.upper())
            if match:
                chart_rows.append((row_cells, match.group(1)))

    present_groups = []
    seen_letters = set()
    for tr in target_rows:
        letter = tr['cargo'][:1]
        if letter in seen_letters:
            continue
        if group_totals.get(letter, {'T': 0})['T'] > 0:
            present_groups.append(letter)
            seen_letters.add(letter)

    chart_total = sum(group_totals.get(letter, {'T': 0})['T'] for letter in present_groups)

    for idx, (row_cells, _) in enumerate(chart_rows):
        if idx < len(present_groups):
            letter = present_groups[idx]
            group_total = group_totals.get(letter, {'T': 0})['T']
            percentage = (float(group_total) / float(chart_total)) if chart_total > 0 else 0.0
            row_cells[15].set_value(f'GRUPO {letter} ')
            set_numeric_or_blank_no_formula(row_cells[16], group_total)
            set_numeric_or_blank_no_formula(row_cells[17], percentage)
        else:
            row_cells[15].set_value('')
            set_numeric_or_blank_no_formula(row_cells[16], 0)
            set_numeric_or_blank_no_formula(row_cells[17], 0)

    for row_cells in all_rows:
        f16 = str(getattr(row_cells[16], 'formula', '')).upper()
        f17 = str(getattr(row_cells[17], 'formula', '')).upper()
        if 'SUM([.Q' in f16:
            set_numeric_or_blank_no_formula(row_cells[16], chart_total)
        if 'SUM([.R' in f17:
            set_numeric_or_blank_no_formula(row_cells[17], 1.0 if chart_total > 0 else 0.0)

def build_processing_summary(source_data, source_stats, changes_made, unmapped, missing_combinations, sheet_names):
    cargo_totals = defaultdict(int)
    profession_totals = defaultdict(int)

    for prof, cargos in source_data.items():
        for cargo, by_gender in cargos.items():
            total = int(by_gender.get('MUJER', 0)) + int(by_gender.get('HOMBRE', 0))
            cargo_totals[cargo] += total
            profession_totals[prof] += total

    top_cargos = [
        {'label': cargo, 'total': total}
        for cargo, total in sorted(cargo_totals.items(), key=lambda item: item[1], reverse=True)[:10]
    ]
    top_professions = [
        {'label': prof, 'total': total}
        for prof, total in sorted(profession_totals.items(), key=lambda item: item[1], reverse=True)[:10]
    ]

    permanent_totals = source_stats['effective_status_totals'].get('PERMANENTE', {'HOMBRE': 0, 'MUJER': 0})
    provisorio_raw = source_stats['raw_status_totals'].get('PROVISORIO', {'HOMBRE': 0, 'MUJER': 0})

    return {
        'total_profesiones': len(source_data),
        'total_cargos': len(cargo_totals),
        'total_hombres': int(permanent_totals.get('HOMBRE', 0)),
        'total_mujeres': int(permanent_totals.get('MUJER', 0)),
        'total_personas': int(permanent_totals.get('HOMBRE', 0)) + int(permanent_totals.get('MUJER', 0)),
        'provisorio_detectado': int(provisorio_raw.get('HOMBRE', 0)) + int(provisorio_raw.get('MUJER', 0)),
        'top_cargos': top_cargos,
        'top_profesiones': top_professions,
        'changes_made': changes_made,
        'unmapped_count': len(unmapped),
        'missing_count': len(missing_combinations),
        'sheet_names': sheet_names,
    }

def cleanup_download_cache():
    now = time.time()
    expired = [token for token, item in DOWNLOAD_CACHE.items() if now - item['created_at'] > CACHE_TTL_SECONDS]
    for token in expired:
        DOWNLOAD_CACHE.pop(token, None)

def store_download(output_filename, output_bytes):
    cleanup_download_cache()
    token = uuid4().hex
    DOWNLOAD_CACHE[token] = {
        'filename': output_filename,
        'bytes': output_bytes,
        'created_at': time.time(),
    }
    return token

def process_files(source_path, target_path, output_path):
    source_data, source_stats = parse_source(source_path)
    doc, all_rows, target_rows = parse_target(target_path)
    sheet_names = [sheet.name for sheet in doc.sheets]
    clean_broken_ref_formulas(doc)
    reset_target_data_columns(all_rows, target_rows)
    prof_map, unmapped = build_profession_map(source_data, target_rows)

    target_index = {}
    for tr in target_rows:
        target_index[(tr['prof'], tr['cargo'])] = tr['row_idx']

    changes_made = 0
    missing_combinations = []
    for sp in source_data:
        tp = prof_map.get(sp)
        if not tp:
            continue
        for cargo in source_data[sp]:
            target_row_idx = target_index.get((tp, cargo))
            
            if target_row_idx is not None:
                row_cells = all_rows[target_row_idx]
                mujeres = source_data[sp][cargo].get('MUJER', 0)
                hombres = source_data[sp][cargo].get('HOMBRE', 0)

                set_numeric_or_blank(row_cells[6], mujeres)
                set_numeric_or_blank(row_cells[7], hombres)
                changes_made += 1
            else:
                missing_combinations.append({'source_prof': sp, 'target_prof': tp, 'cargo': cargo})

    rebuild_totals_and_chart_data(all_rows, target_rows)
    by_prof_level, by_level_series = build_profession_totals_from_target_rows(all_rows, target_rows)
    fill_sheet_prof_ab(doc, by_prof_level)
    fill_sheet_level_series(doc, by_level_series)

    doc.saveas(output_path)
    summary = build_processing_summary(
        source_data=source_data,
        source_stats=source_stats,
        changes_made=changes_made,
        unmapped=unmapped,
        missing_combinations=missing_combinations,
        sheet_names=sheet_names,
    )
    return {
        'changes_made': changes_made,
        'unmapped': unmapped,
        'missing_combinations': missing_combinations,
        'summary': summary,
    }

@app.route('/download/<token>', methods=['GET'])
def download_processed_file(token):
    item = DOWNLOAD_CACHE.pop(token, None)
    if not item:
        flash('El archivo ya no está disponible. Vuelve a procesar las planillas.', 'error')
        return redirect(url_for('index'))

    return send_file(
        io.BytesIO(item['bytes']),
        as_attachment=True,
        download_name=item['filename'],
        mimetype='application/vnd.oasis.opendocument.spreadsheet',
    )

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
            temp_dir = tempfile.mkdtemp(prefix='planilla_')
            try:
                source_path = os.path.join(temp_dir, secure_filename(source_file.filename))
                target_path = os.path.join(temp_dir, secure_filename(target_file.filename))

                # Keep original extension but add _actualizado
                base, ext = os.path.splitext(target_file.filename)
                output_filename = f"{base}_actualizado{ext}"
                output_path = os.path.join(temp_dir, output_filename)

                source_file.save(source_path)
                target_file.save(target_path)

                result = process_files(source_path, target_path, output_path)
                print('Resultado de procesamiento:', result)
                with open(output_path, 'rb') as f:
                    output_bytes = f.read()

                download_token = store_download(output_filename, output_bytes)

                return render_template(
                    'index.html',
                    result_summary=result['summary'],
                    download_token=download_token,
                    output_filename=output_filename,
                )
            except Exception as e:
                flash(f'Ocurrió un error: {str(e)}', 'error')
                return redirect(request.url)
            finally:
                cleanup_temp_dir(temp_dir)
                    
    return render_template('index.html')

if __name__ == '__main__':
    # Start web browser automatically
    import threading
    import webbrowser
    def open_browser():
        webbrowser.open_new("http://127.0.0.1:5000/")
    threading.Timer(1, open_browser).start()
    app.run(debug=False, use_reloader=False)
