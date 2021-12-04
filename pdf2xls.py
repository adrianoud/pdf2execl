import pdfplumber
import openpyxl
wb = openpyxl.load_workbook("thermalData.xlsx")


cells = []
pdf = pdfplumber.open("111.pdf")
page0 = pdf.pages[2]
table = page0.extract_text()
rows = table.split("\n")
for i in range(len(rows)):
    cells.append([])
    cols = rows[i].split(" ")
    for j in range(len(cols)):
        cells[i].append(cols[j])
print(cells[12][0])

if cells[1][3] in wb.sheetnames:
    sheet = wb[cells[1][3]]
else:
    sheet = wb.create_sheet(cells[1][3])

sheet.cell(2, 1).value = "工况："
sheet.cell(3, 1).value = cells[0][0]
sheet.cell(1, 2).value = "烟气侧"
sheet.cell(1, 12).value = "工质侧"
for i in range(2, 12):
    sheet.cell(2, i).value = cells[i+1][0]

wb.save("thermalData.xlsx")

