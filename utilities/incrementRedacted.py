import os
import re

'''
Increments the numeric part of a filename that ends with " redacted.pdf". If the filename 
matches the expected pattern, the numeric part is incremented by 1, and the filename is 
reformatted with the incremented number while maintaining leading zeros and the original 
suffix.

@param fileName (str): The filename to be incremented.

@return newFileName (str): The filename with its numeric part incremented by 1. 
                            Returns the original filename if it does not match the expected 
                            pattern.

- Uses regex to identify the base name, numeric part, and extension of the filename.
- Increments the numeric part and reconstructs the filename with the new number.

@author Mike
'''
def incrementFilename(fileName):
    match = re.match(r"(.+?)(\d+)(\sredacted\.pdf)$", fileName)
    if not match:
        return fileName
    base, number, ext = match.groups()
    newNumber = int(number) + 1
    newFileName = f"{base}{newNumber:05d}{ext}" 
    return newFileName


'''
Retrieves and sorts the filenames in a specified directory based on the numeric part of 
each filename. Filenames are sorted in descending order to prioritize higher numbers, with 
files lacking a numeric part sorted to the beginning.

@param directory (str): The path of the directory from which files are listed and sorted.

@return files (list): A list of filenames sorted based on their numeric parts in descending 
                      order.

- Uses regex to extract the numeric part of each filename for sorting purposes.
- Handles filenames without a numeric part by defaulting their sort value to 0.

@author Mike
'''
def getSortedFiles(directory):
    files = os.listdir(directory)
    files.sort(key=lambda f: int(re.search(r"(\d+)", f).group()) if re.search(r"(\d+)", f) else 0, reverse=True)
    return files


'''
Main function to increment and rename files within a specified directory according to a 
predefined naming convention. It utilizes a sorting and renaming strategy to process files 
in descending numeric order, incrementing each file's numeric identifier.

- Retrieves a sorted list of files from the specified directory, focusing on those with 
  numeric identifiers.
- Iterates over a subset of these files, from the highest (or last) in the sorted list up 
  to a target index, to rename each.
- Renames files by incrementing their numeric part, ensuring no file name conflicts occur 
  in the process.
- Reports the outcome of each renaming attempt, indicating success or failure due to 
  potential name conflicts.

The process is designed to accommodate files ending with " redacted.pdf", incrementing 
their numeric part while avoiding overwriting existing files.

@author Mike
'''
def main():
    directoryPath = r"\\ucclerk\pgmdoc\Veterans\Cemetery - Redacted\Fairview - Redacted\P"
    files = getSortedFiles(directoryPath)
    startIndex = 0 # 0 for the last file in the sorted list
    targetIndex = 149  # Example: stop after processing 10 files
    for i in range(startIndex, min(targetIndex, len(files))):
        file = files[i]
        oldPath = os.path.join(directoryPath, file)
        newFile = incrementFilename(file)
        newPath = os.path.join(directoryPath, newFile)
        if not os.path.exists(newPath):
            os.rename(oldPath, newPath)
            print(f"Renamed '{file}' to '{newFile}'")
        else:
            print(f"Cannot rename '{file}' to '{newFile}' as it already exists.")


if __name__ == "__main__":
    main()
