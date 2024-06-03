import os
import re
from openpyxl.worksheet.hyperlink import Hyperlink
import openpyxl

def cleanDelete(badWorkSheet, badID, badRow):
    # baseRedacPath = r"\\ucclerk\pgmdoc\Veterans\Cemetery - Redacted"
    # for dirpath, dirnames, filenames in os.walk(baseRedacPath):
    #     for filename in filenames:
    #         if f"{badID:05d}" in filename:
    #             file_path = os.path.join(dirpath, filename)
    #             os.remove(file_path)
    #             print(f"\nRedacted file, {filename}, deleted successfully.")
    badWorkSheet.delete_rows(badRow)
    print(f"Row {badRow} deleted successfully.")
    for row in range(badRow, badWorkSheet.max_row + 1):
        next_row = row + 1
        cell_ref = f'O{row}'
        next_cell_ref = f'O{next_row}'
        if next_row <= badWorkSheet.max_row and badWorkSheet[next_cell_ref].hyperlink:
            temp_hyperlink = badWorkSheet[next_cell_ref].hyperlink
            badWorkSheet[cell_ref].hyperlink = Hyperlink(ref=cell_ref, target=temp_hyperlink.target, display=temp_hyperlink.display)
            badWorkSheet[cell_ref].value = badWorkSheet[next_cell_ref].value
            print(f"Hyperlink and value from row {next_row} moved to row {row}.")
    print(badWorkSheet.max_row - 1)
    if badWorkSheet[f'O{badWorkSheet.max_row - 1}'].hyperlink:
        last_row = badWorkSheet.max_row
        print(f"Adjusting the hyperlink in the last row, row {last_row}.")
        second_last_hyperlink = badWorkSheet[f'O{last_row - 1}'].hyperlink
        new_target = re.sub(r'(\d+)(?=%\d+)', lambda m: f"{int(m.group(1)) + 1:05d}", second_last_hyperlink.target)
        badWorkSheet[f'O{last_row}'].hyperlink = Hyperlink(ref=f'O{last_row}', target=new_target, display=second_last_hyperlink.display)
        print(f"Updated hyperlink in row {last_row} to new target, {new_target}.\n")

excelFilePath = r"\\ucclerk\pgmdoc\Veterans\Veterans2.xlsx"
workbook = openpyxl.load_workbook(excelFilePath)
sheet = workbook["Graceland"]
cleanDelete(sheet, 10370, 2130)
workbook.save(excelFilePath)