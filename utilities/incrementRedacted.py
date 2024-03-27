import os
import re

def increment_filename(filename):
    match = re.match(r"(.+?)(\d+)(\sredacted\.pdf)$", filename)
    if not match:
        return filename

    base, number, ext = match.groups()
    new_number = int(number) + 1
    new_filename = f"{base}{new_number:05d}{ext}" 
    return new_filename

def get_sorted_files(directory):
    files = os.listdir(directory)
    files.sort(key=lambda f: int(re.search(r"(\d+)", f).group()) if re.search(r"(\d+)", f) else 0, reverse=True)
    return files

directory_path = r"\\ucclerk\pgmdoc\Veterans\Cemetery - Redacted\Fairview - Redacted\P"
files = get_sorted_files(directory_path)

start_index = 0 # 0 for the last file in the sorted list
target_index = 149  # Example: stop after processing 10 files

for i in range(start_index, min(target_index, len(files))):
    file = files[i]
    old_path = os.path.join(directory_path, file)
    new_file = increment_filename(file)
    new_path = os.path.join(directory_path, new_file)
    if not os.path.exists(new_path):
        os.rename(old_path, new_path)
        print(f"Renamed '{file}' to '{new_file}'")
    else:
        print(f"Cannot rename '{file}' to '{new_file}' as it already exists.")
