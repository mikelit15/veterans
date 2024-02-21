import os
import re

def process_cemetery(cemetery_name, base_path, initial_count, start_index):
    uppercase_alphabet = [chr(i) for i in range(ord('A'), ord('Z') + 1)]
    first_file_flag = True
    for letter in uppercase_alphabet:
        name_path = os.path.join(base_path, cemetery_name, letter)
        try:
            print(f"Processing {cemetery_name} - {letter}")
            initial_count = clean(cemetery_name, letter, name_path, os.path.join(base_path, cemetery_name), initial_count, first_file_flag, start_index)
            first_file_flag = False
        except FileNotFoundError:
            continue
    return initial_count

def clean(cemetery, letter, name_path, cem_path, counterA, is_first_file, start_index):
    pdf_files = sorted(os.listdir(name_path), key=lambda x: int(re.search(r'\d+', x).group()) if re.search(r'\d+', x) else 0)
    sequence_started = False  # Flag to indicate when to start renumbering

    for x in pdf_files:
        match = re.search(r'(\d+)', x)
        if match:
            file_number = int(match.group())
            if file_number >= start_index:
                if not sequence_started:
                    sequence_started = True  # Start renumbering from this file onwards
                    counter = start_index  # Initialize counter at start_index
                else:
                    counter += 1  # Increment counter for subsequent files
                
                # Directly construct the new filename using the cemetery name, letter, and new counter value
                new_name_suffix = x[re.search(r'(\d+)', x).end():]  # Extract suffix from original filename
                newName = f"{cemetery}{letter}{counter:05d}{new_name_suffix}"
                os.rename(os.path.join(name_path, x), os.path.join(name_path, newName))
                print(f"Renamed '{x}' to '{newName}'")
            # Files numbered less than start_index are not renamed, hence no else clause

    return counterA  # You might want to update this return value based on your specific needs


def mainLoop(start_index):  
    base_cem_path = r"\\ucclerk\pgmdoc\Veterans\Cemetery1"
    cemeterys = [d for d in os.listdir(base_cem_path) if os.path.isdir(os.path.join(base_cem_path, d))]
    initial_count = 1
    for cemetery in cemeterys:
        if cemetery in ["Jewish", "Misc"]:
            sub_cem_path = os.path.join(base_cem_path, cemetery)
            sub_cemeteries = [d for d in os.listdir(sub_cem_path) if os.path.isdir(os.path.join(sub_cem_path, d))]
            for sub_cemetery in sub_cemeteries:
                print(f"Processing {cemetery} - {sub_cemetery}")
                initial_count = process_cemetery(sub_cemetery, sub_cem_path, initial_count, start_index)
        else:
            print(f"Processing {cemetery}")
            initial_count = process_cemetery(cemetery, base_cem_path, initial_count, start_index)

if __name__ == "__main__":
    start_index = 135 
    mainLoop(start_index)
