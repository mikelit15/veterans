import openpyxl
import re
import os


'''
Adjusts hyperlinks in an Excel file that contain numerical IDs, decrementing those IDs by 1
starting from a specified index. This function is used when IDs in a series are modified and 
corresponding hyperlinks need to reflect these changes. Hyperlinks with IDs greater than or 
equal to startIndex will have their ID decremented by 1.

@param filePath (str) - The path to the Excel file to be modified.
@param columnLetter (str) - The letter of the column containing hyperlinks to be adjusted.
@param startIndex (int) - The row number from which to start adjusting hyperlinks.

@author Mike
'''     
def cleanHyperlink(id, cemetery):
    def decrementWithRollover(path):
        pattern = re.compile(fr"({cemetery})([A-Z])(\d+)(\s*redacted\.pdf)")
        def handleDecrement(match):
            prefix, letter, num, suffix = match.groups()
            newNum = int(num) - 1
            if newNum < 0:
                newNum = 9999  
            newNumberFormatted = f"{newNum:0{len(num)}d}"
            return f"{prefix}{letter}{newNumberFormatted}{suffix}"
        return pattern.sub(handleDecrement, path)
    excelFilePath = r"\\ucclerk\pgmdoc\Veterans\Veterans.xlsx"
    columnWithHyperlinks = 'O'
    startIndex = id 
    workbook = openpyxl.load_workbook(excelFilePath)
    sheet = workbook[cemetery]
    for row in sheet.iter_rows(min_col=openpyxl.utils.column_index_from_string(columnWithHyperlinks),
                               max_col=openpyxl.utils.column_index_from_string(columnWithHyperlinks),
                               min_row=startIndex):
        cell = row[0]
        if cell.hyperlink:
            url = cell.hyperlink.target
            url = url.replace("%20", " ")
            newURL = decrementWithRollover(url)
            if newURL != url:
                cell.hyperlink.target = newURL
                print(f"Updated hyperlink from {os.path.basename(url)} to {os.path.basename(newURL)}")
    workbook.save(excelFilePath)
    print("Hyperlinks updated.")


if __name__ == "__main__":
    id = 1800 # The hyperlink gets decremented by 1 at this row #
    cemetery = "Graceland"
    cleanHyperlink(id, cemetery)