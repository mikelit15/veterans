import os 
import shutil


def clean_and_rename_folders(parent_folder):
    # Iterate over all the directories in the parent folder
    for folder_name in os.listdir(parent_folder):
        folder_path = os.path.join(parent_folder, folder_name)
        
        # Check if it's a directory
        if os.path.isdir(folder_path):
            # Iterate over all the directories in each alphabet folder
            for subfolder_name in os.listdir(folder_path):
                subfolder_path = os.path.join(folder_path, subfolder_name)
                
                # Check if it's a directory
                if os.path.isdir(subfolder_path):
                    # Iterate over all the files in the subfolder and delete them
                    for file_name in os.listdir(subfolder_path):
                        file_path = os.path.join(subfolder_path, file_name)
                        
                        # Check if it's a file and delete it
                        if os.path.isfile(file_path):
                            os.remove(file_path)
                            print(f"Deleted {file_path}")

# Set your parent folder path
parent_folder = r"\\ucclerk\pgmdoc\Veterans\Cemetery - Redacted\Misc - Redacted"
clean_and_rename_folders(parent_folder)