import os
import re
    

def process_cemetery(cemetery_name, base_path, initial_count):
    uppercase_alphabet = [chr(i) for i in range(ord('A'), ord('Z') + 1)]
    first_file_flag = True
    for letter in uppercase_alphabet:
        name_path = os.path.join(base_path, cemetery_name, letter)
        try:
            print(f"Processing {cemetery_name} - {letter}")
            initial_count = clean(cemetery_name, letter, name_path, os.path.join(base_path, cemetery_name), initial_count, first_file_flag)
            first_file_flag = False
        except FileNotFoundError:
            continue
    return initial_count


'''
Cleans and renames file names in a given letter and cemetery for inconsistencies 
when scanning. It standardizes the naming convention of files and ensures correct ordering.

@param cemetery (str) - The name of the cemetery
@param letter (str) - The letter or section of the cemetery being processed
@param name_path (str) - The path where the files are located
@param cem_path (str) - The path to the cemetery directory
@param counterA (int) - The starting counter for naming files
@param is_file_first (bool) - Flag indicating if the current file is the first in the series

@author Mike
'''   
def clean(cemetery, letter, name_path, cem_path, counterA, is_first_file):
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
        if is_first_file:
            counter = counterA 
        if "a.pdf" in x:
            newName = f"{cemetery}{letter}{counter:05d}a.pdf"
            os.rename(os.path.join(name_path, x), os.path.join(name_path, newName))
            counter -= 1
        elif "b.pdf"  in x:
            newName = f"{cemetery}{letter}{counter:05d}b.pdf"
            os.rename(os.path.join(name_path, x), os.path.join(name_path, newName))
        elif "a redacted" in x:
            newName = f"{cemetery}{letter}{counter:05d}a redacted.pdf"
            os.rename(os.path.join(name_path, x), os.path.join(name_path, newName))
            counter -= 1
        elif "b redacted" in x:
            newName = f"{cemetery}{letter}{counter:05d}b redacted.pdf"
            os.rename(os.path.join(name_path, x), os.path.join(name_path, newName))
            counter -= 1
        elif "redacted" in x:
            newName = f"{cemetery}{letter}{counter:05d} redacted.pdf"
            os.rename(os.path.join(name_path, x), os.path.join(name_path, newName))
            counter -= 1
        else:    
            newName = f"{cemetery}{letter}{counter:05d}.pdf"
            os.rename(os.path.join(name_path, x), os.path.join(name_path, newName))
        counter += 1
        is_first_file = False
    return counter

def mainLoop():
    base_cem_path = r"\\ucclerk\pgmdoc\Veterans\Cemetery"
    cemeterys = [d for d in os.listdir(base_cem_path) if os.path.isdir(os.path.join(base_cem_path, d))]
    initial_count = 1
    for cemetery in cemeterys:
        if cemetery in ["Jewish", "Misc"]:
            sub_cem_path = os.path.join(base_cem_path, cemetery)
            sub_cemeteries = [d for d in os.listdir(sub_cem_path) if os.path.isdir(os.path.join(sub_cem_path, d))]
            for sub_cemetery in sub_cemeteries:
                print(f"Processing {cemetery} - {sub_cemetery}")
                initial_count = process_cemetery(sub_cemetery, sub_cem_path, initial_count)
        else:
            print(f"Processing {cemetery}")
            initial_count = process_cemetery(cemetery, base_cem_path, initial_count)

def mainTarget():
    network_folder = r"\\ucclerk\pgmdoc\Veterans"
    cemetery = "Fairview"
    os.chdir(network_folder)
    cem_path = os.path.join(network_folder, fr"Cemetery\{cemetery}")
    initialCount = 7000
    uppercase_alphabet = [chr(i) for i in range(ord('A'), ord('Z') + 1)]
    firstFileFlag = True
    for letter in uppercase_alphabet:
        try:
            name_path = os.path.join(cem_path, letter)
            print(letter)
            clean(cemetery, letter, name_path, cem_path, initialCount, firstFileFlag)
            firstFileFlag = False
        except FileNotFoundError:
            continue
        
if __name__ == "__main__":
    mainLoop()