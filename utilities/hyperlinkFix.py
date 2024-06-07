import openpyxl
from openpyxl.styles import Font

workbookPath = r"\\ucclerk\pgmdoc\Veterans\Veterans.xlsx" 
workbook = openpyxl.load_workbook(workbookPath)
worksheet = workbook.active
linkText = "PDF Image"
linkFont = Font(underline="single", color="0563C1") 
dynamicPart = "B"
for rowIndex in range(128, 568):
    cell = worksheet.cell(row=rowIndex, column=15)
    redactedFile = f"\\\\ucclerk\\pgmdoc\\Veterans\\Cemetery - Redacted\\Evergreen - Redacted\\{dynamicPart}\\Evergreen{dynamicPart}{str(rowIndex-1).zfill(5)} redacted.pdf"
    cell.value = linkText
    cell.font = linkFont
    cell.hyperlink = redactedFile

workbook.save(r"\\ucclerk\pgmdoc\Veterans\Veterans.xlsx" )
