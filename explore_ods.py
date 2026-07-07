import ezodf

doc = ezodf.opendoc('c:/Users/e555044.NTDOM1/Desktop/planilla/FUNCIONARIOS POR GRUPO OCUPACIONAL dic 2025 v1.ods')
sheet = doc.sheets[0]

for i, row in enumerate(sheet.rows()):
    row_values = []
    for cell in row:
        val = cell.value
        row_values.append(str(val) if val is not None else '')
    if any(row_values):
        print(f"Row {i}: {row_values}")
    if i > 60:
        break
