import os

directory = r"\\ucclerk\pgmdoc\Veterans\Cemetery\Rosehill Crematory\O"
for filename in os.listdir(directory):
    if filename.startswith("Rosehill CrematoryN") and filename.endswith(".pdf"):
        # Construct the new filename by replacing 'N' with 'O'
        new_filename = filename.replace("Rosehill CrematoryN", "Rosehill CrematoryO")
        
        # Construct the full file paths
        old_file = os.path.join(directory, filename)
        new_file = os.path.join(directory, new_filename)
        
        # Rename the file
        os.rename(old_file, new_file)
        print(f"Renamed '{filename}' to '{new_filename}'")
