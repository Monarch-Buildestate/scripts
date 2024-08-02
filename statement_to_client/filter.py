import openpyxl


wb1 = openpyxl.open("statement.xlsx")
wb2 = openpyxl.open("statement2.xlsx")
unit_names_allowed = []
for row in wb1.active.rows:
    unit_names_allowed.append(row[0].value)

filtered_rows = []
for row in wb2.active.rows:
    if row[3].value not in unit_names_allowed:
        print(row[3].value)
    else:
        filtered_rows.append(row)

wb3 = openpyxl.Workbook()
ws = wb3.active
for row in filtered_rows:
    ws.append([cell.value for cell in row])

wb3.save("filtered.xlsx")