import re
from openpyxl import load_workbook

s = input('Enter number of column with emails to extract [1]: ')
colum = 1 if s == '' else int(s)

# open workbook from file
wb = load_workbook('ss.xlsx')
# grab the active worksheet
ws = wb.active

# go through cells
for i in range(1, ws.max_row+1):
    content = ""
    cc = ws.cell(row=i, column=colum).value
    if type(cc) is str:
        match = re.search(r'[\w\.-]+@[\w\.-]+', cc)
        if match:
            content=match.group(0)
        else:
            content=""
    ws.cell(row=i, column=colum, value=content)

# Save the file
wb.save("ss.xlsx")

print("Done!")
