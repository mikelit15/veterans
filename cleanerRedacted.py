import os
import re

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

def main():
    base_cem_path = r"\\ucclerk\pgmdoc\Veterans\Cemetery - Redacted"
    start_index = 2638 # this file gets decremented
    decrement_file_numbers(base_cem_path, start_index)

if __name__ == "__main__":
    main()
