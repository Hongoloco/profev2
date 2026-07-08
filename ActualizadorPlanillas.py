import tkinter as tk
from tkinter import filedialog, messagebox
import os
from webapp.app import process_files as process_files_web

def process_files(source_path, target_path):
    try:
        base, ext = os.path.splitext(target_path)
        out_path = f"{base}_actualizado{ext}"
        result = process_files_web(source_path, target_path, out_path)
        
        total_personas = result['summary'].get('total_personas', 0)
        return True, (
            f"Se actualizó exitosamente y se guardó como:\n{out_path}\n\n"
            f"Cambios realizados: {result.get('changes_made', 0)}\n"
            f"Personas cargadas en destino: {total_personas}\n"
            f"Sin mapeo: {len(result.get('unmapped', []))}\n"
            f"Combinaciones faltantes de cargo: {len(result.get('missing_combinations', []))}"
        )
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
