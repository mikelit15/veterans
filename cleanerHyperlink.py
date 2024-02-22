import openpyxl
import re

def decrement_hyperlinks_in_excel(file_path, column_letter, start_index):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active
    for row in sheet.iter_rows(min_col=openpyxl.utils.column_index_from_string(column_letter),
                               max_col=openpyxl.utils.column_index_from_string(column_letter),
                               min_row=start_index): 
        cell = row[0] 
        if cell.hyperlink: 
            url = cell.hyperlink.target
            new_url = re.sub(r'(\d+)', lambda x: f"{int(x.group())-1:05d}" if int(x.group()) >= start_index else x.group(), url)
            if new_url != url:
                cell.hyperlink.target = new_url
                print(f"Updated hyperlink: {url} to {new_url}")
    workbook.save(file_path)
    print(f"Hyperlinks updated.")

excel_file_path = r"\\ucclerk\pgmdoc\Veterans\Veterans.xlsx" 
column_with_hyperlinks = 'O'  
start_index = 3015 # ID number
decrement_hyperlinks_in_excel(excel_file_path, column_with_hyperlinks, start_index)
