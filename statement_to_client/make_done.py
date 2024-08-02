import openpyxl
wb = openpyxl.open("statement.xlsx")
import json

done = []
count = 0

for row in wb.active.rows:
    if count > 353:
        break
    count+=1
    done.append(f"{row[4].value} - {row[3].value} - {row[8].value}")

with open("done.json", "w") as f:
    json.dump(done, f, indent=4)
print("done")