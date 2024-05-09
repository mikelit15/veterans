import os
import re
import openpyxl
from openpyxl.worksheet.hyperlink import Hyperlink

# path = r"\\ucclerk\pgmdoc\Veterans\pngFiles"

# for dirpath, dirnames, filenames in os.walk(path):
#         for filename in filenames:
#             if "Page_1" in filename:
#                 os.rename(os.path.join(path, filename), os.path.join(path, filename.replace("_Page_1", "")))
#             elif "Page_2" in filename:
#                 os.remove(os.path.join(path,filename))

def cleanHyperlinks(cemetery, startIndex):
    base_path = fr'\\ucclerk\pgmdoc\Veterans\Cemetery\{cemetery}'
    file_directory_map = {}
    for dirpath, dirnames, filenames in os.walk(base_path):
        for filename in filenames:
            file_id = ''.join(filter(str.isdigit, filename))[:5]
            if file_id:
                file_directory_map[file_id] = dirpath
    new_id = worksheet[f'A{startIndex}'].value
    for row in range(startIndex, worksheet.max_row + 1):
        worksheet[f'A{row}'].value = new_id
        cell_ref = f'O{row}'
        if worksheet[cell_ref].hyperlink:
            orig_target = worksheet[cell_ref].hyperlink.target.replace("%20", " ")
            formatted_id = f"{new_id:05d}"
            folder_name = file_directory_map.get(formatted_id, "UnknownFolder")[-1]
            modified_string = re.sub(fr'(\\)([A-Z])(\\)', f'\\1{folder_name}\\3', orig_target)
            modified_string = re.sub(fr'({cemetery}[A-Z])\d{{5}}', f'{cemetery}{folder_name}{formatted_id}', modified_string)
            # worksheet[cell_ref].hyperlink = Hyperlink(ref=cell_ref, target=modified_string, display="PDF Image")
            # worksheet[cell_ref].value = "PDF Image"
            print(f"Updated hyperlink from {orig_target} to {modified_string} in row {row}.")
        new_id += 1
        
cemetery = "Graceland"
excelFilePath = r"\\ucclerk\pgmdoc\Veterans\Veterans2.xlsx"
workbook = openpyxl.load_workbook(excelFilePath)
worksheet = workbook[cemetery]
cleanHyperlinks(cemetery, 63)