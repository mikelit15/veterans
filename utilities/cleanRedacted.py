import os
import re

'''
Decrements the numerical part of file names in a directory and its subdirectories, 
starting from a specified index. This is used to maintain a sequential order after 
removing files. Files with a number greater than or equal to start_index will have 
their number decremented by 1.

@param base_path (str) - The path to the directory containing the files to be renamed.
@param start_index (int) - The file number from which to start decrementing.

@author Mike
'''      
def decrementFileNumbers(basePath, startIndex):
    for root, dirs, files in os.walk(basePath):
        for file in sorted(files):
            match = re.search(r'(\d+)', file)
            if match:
                currentIndex = int(match.group())
                if currentIndex >= startIndex:
                    newIndex = currentIndex - 1
                    newFileName = re.sub(r'\d+', f"{newIndex:05}", file, 1)  
                    os.rename(os.path.join(root, file), os.path.join(root, newFileName))
                    print(f"Renamed '{file}' to '{newFileName}'")


'''
Initiates the file renaming process for redacted documents within a specific directory,
starting from a given file ID. This is part of managing file sequences after some documents 
are removed or redacted. Files in the directory and subdirectories with ID numbers greater 
than this will be decremented.

@param id (int) - The ID from which to start the cleaning process, affecting files with ID 
                  numbers greater by adjusting their names to maintain sequence.

@author Mike
'''
def cleanRedacted(id):
    baseCemPath = r"\\ucclerk\pgmdoc\Veterans\Cemetery - Redacted"
    startIndex = id + 1 # The image file name gets decremented starting at this ID number
    decrementFileNumbers(baseCemPath, startIndex)
    
    
if __name__ == "__main__":
    id = 58 # Proper ID number
    cleanRedacted(id)