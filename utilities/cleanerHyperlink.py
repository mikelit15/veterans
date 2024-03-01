import openpyxl
import re

'''
Adjusts hyperlinks in an Excel file that contain numerical IDs, decrementing those IDs by 1
starting from a specified index. This function is used when IDs in a series are modified and 
corresponding hyperlinks need to reflect these changes. Hyperlinks with IDs greater than or 
equal to start_index will have their ID decremented by 1.

@param filePath (str) - The path to the Excel file to be modified.
@param columnLetter (str) - The letter of the column containing hyperlinks to be adjusted.
@param startIndex (int) - The row number from which to start adjusting hyperlinks.

@author Mike
'''     
def decrement_hyperlinks_in_excel(filePath, columnLetter, startIndex):
    workbook = openpyxl.load_workbook(filePath)
    sheet = workbook.active
    for row in sheet.iter_rows(min_col=openpyxl.utils.column_index_from_string(columnLetter),
                               max_col=openpyxl.utils.column_index_from_string(columnLetter),
                               min_row=startIndex): 
        cell = row[0] 
        if cell.hyperlink: 
            url = cell.hyperlink.target
            new_url = re.sub(r'(\d+)', lambda x: f"{int(x.group()) - 1:05d}" if int(x.group()) >= startIndex else x.group(), url)
            if new_url != url:
                cell.hyperlink.target = new_url
                print(f"Updated hyperlink: {url} to {new_url}")
    workbook.save(filePath)
    print(f"Hyperlinks updated.")


'''
Initiates the process of decrementing hyperlink IDs in an Excel document, starting from a 
specified ID. This is particularly useful in maintaining accurate references after modifications 
to the document IDs they point to. Hyperlinks in the specified column that contain an ID greater 
than or equal to the specified start ID will be adjusted to reflect decremented ID values.

@param id (int) - The ID at which to start decrementing hyperlink references in the document.

@author Mike
'''
def cleanHyperlink(id):
    excelFilePath = r"\\ucclerk\pgmdoc\Veterans\Veterans.xlsx" 
    columnWithHyperlinks = 'O'  
    startIndex = id
    decrement_hyperlinks_in_excel(excelFilePath, columnWithHyperlinks, startIndex)


if __name__ == "__main__":
    id = 58 # The hyperlink gets decremented by 1 at this cell ID number
    cleanHyperlink(id)