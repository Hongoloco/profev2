import ezodf
doc = ezodf.opendoc('c:/Users/e555044.NTDOM1/Desktop/planilla/FUNCIONARIOS POR GRUPO OCUPACIONAL dic 2025 v1_actualizado.ods')
sheet = doc.sheets[0]
for i, row in enumerate(list(sheet.rows())):
    if i > 50:
        break
    m = row[6].value
    h = row[7].value
    if m or h:
        print(f"Row {i}: M={m}, H={h}")
