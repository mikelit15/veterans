import os
import re
import openpyxl

def process_cemetery(cemetery_name, base_path, initial_count, start, startFlag):
    uppercase_alphabet = [chr(i) for i in range(ord('A'), ord('Z') + 1)]
    first_file_flag = True
    for letter in uppercase_alphabet:
        name_path = os.path.join(base_path, cemetery_name, letter)
        try:
            print(f"Processing {cemetery_name} - {letter}")
            initial_count, startFlag = clean(cemetery_name, letter, name_path, os.path.join(base_path, cemetery_name), \
                initial_count, first_file_flag, start, startFlag)
            first_file_flag = False
        except FileNotFoundError:
            continue
    return initial_count, startFlag


def clean(cemetery, letter, name_path, cem_path, counterA, is_first_file, start, startFlag):
    pdf_files = sorted(os.listdir(name_path))
    letters = sorted([folder for folder in os.listdir(cem_path) if os.path.isdir(os.path.join(cem_path, folder))])
    try:
        current_letter_index = letters.index(letter)
    except ValueError:
        return
    folder_before_index = max(0, current_letter_index - 1)
    folder_before = letters[folder_before_index]
    pdf_files_before = sorted([file for file in os.listdir(os.path.join(cem_path, folder_before)) if file.lower().endswith('.pdf')])
    max_counter = 0
    for file in pdf_files_before:
        match = re.search(r'\d+', file)
        if match:
            number = int(match.group())
            max_counter = max(max_counter, number)
    counter = max_counter + 1
    for x in pdf_files:
        if counter != start and startFlag:
            is_first_file = False
            if "a" in x[-5:]:
                pass
            else:
                counter += 1
            continue
        startFlag = False
        if is_first_file:
            counter = counterA 
        if "a.pdf" in x:
            newName = f"{cemetery}{letter}{counter:05d}a.pdf"
            os.rename(os.path.join(name_path, x), os.path.join(name_path, newName))
            counter -= 1
        elif "b.pdf"  in x:
            newName = f"{cemetery}{letter}{counter:05d}b.pdf"
            os.rename(os.path.join(name_path, x), os.path.join(name_path, newName))
        else:    
            newName = f"{cemetery}{letter}{counter:05d}.pdf"
            os.rename(os.path.join(name_path, x), os.path.join(name_path, newName))
        counter += 1
        is_first_file = False
    return counter, startFlag


def cleanImages(id):
    base_cem_path = r"\\ucclerk\pgmdoc\Veterans\Cemetery"
    cemeterys = [d for d in os.listdir(base_cem_path) if os.path.isdir(os.path.join(base_cem_path, d))]
    initial_count = 1
    start = id - 200
    startFlag = True
    for cemetery in cemeterys:
        if cemetery in ["Jewish", "Misc"]:
            sub_cem_path = os.path.join(base_cem_path, cemetery)
            sub_cemeteries = [d for d in os.listdir(sub_cem_path) if os.path.isdir(os.path.join(sub_cem_path, d))]
            for sub_cemetery in sub_cemeteries:
                print(f"Processing {cemetery} - {sub_cemetery}")
                initial_count, startFlag = process_cemetery(sub_cemetery, sub_cem_path, initial_count, start, startFlag)
        else:
            print(f"Processing {cemetery}")
            initial_count, startFlag = process_cemetery(cemetery, base_cem_path, initial_count, start, startFlag)
        
        
def decrement_file_numbers(base_path, start_index):
    for root, dirs, files in os.walk(base_path):
        for file in sorted(files):
            match = re.search(r'(\d+)', file)
            if match:
                current_index = int(match.group())
                if current_index >= start_index:
                    new_index = current_index - 1
                    new_file_name = re.sub(r'\d+', f"{new_index:05}", file, 1)  
                    os.rename(os.path.join(root, file), os.path.join(root, new_file_name))
                    print(f"Renamed '{file}' to '{new_file_name}'")


def cleanRedacted(id):
    base_cem_path = r"\\ucclerk\pgmdoc\Veterans\Cemetery - Redacted"
    start_index = id + 1 # The image file name gets decremented starting at this ID number
    decrement_file_numbers(base_cem_path, start_index)
       
       
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


def cleanHyperlink(id):
    excel_file_path = r"\\ucclerk\pgmdoc\Veterans\Veterans.xlsx" 
    column_with_hyperlinks = 'O'  
    start_index = id # The hyperlink gets decremented by 1 at this cell ID number
    decrement_hyperlinks_in_excel(excel_file_path, column_with_hyperlinks, start_index)
       
       
if __name__ == "__main__":
    id = 3094 #proper ID number
    cleanRedacted(id)
    cleanHyperlink(id)
    cleanImages(id)