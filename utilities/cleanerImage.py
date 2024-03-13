import os
import re

'''
Processes each cemetery by iterating through alphabetically named subdirectories and renaming PDF files within.
It starts renaming files from a specified point, controlled by 'start' and 'startFlag'.

@param cemeteryName (str) - The name of the cemetery being processed.
@param basePath (str) - The base directory path where cemeteries are located.
@param initialCount (int) - The initial counter for renaming files.
@param start (int) - The starting file number from which renaming begins.
@param startFlag (bool) - A flag indicating whether the start point has been reached.

@return initialCount (int) - Updated counter after processing.
@return startFlag (bool) - Updated flag indicating if the start point was reached and surpassed.

@author Mike
'''
def process_cemetery(cemeteryName, basePath, initialCount, start, startFlag):
    uppercase_alphabet = [chr(i) for i in range(ord('A'), ord('Z') + 1)]
    firstFileFlag = True
    for letter in uppercase_alphabet:
        name_path = os.path.join(basePath, cemeteryName, letter)
        try:
            print(f"Processing {cemeteryName} - {letter}")
            initialCount, startFlag = clean(cemeteryName, letter, name_path, os.path.join(basePath, cemeteryName), \
                initialCount, firstFileFlag, start, startFlag)
            firstFileFlag = False
        except FileNotFoundError:
            continue
    return initialCount, startFlag


'''
Renames PDF files within a specified directory, starting from a given file number. It skips renaming until the start condition is met.

@param cemetery (str) - The cemetery name.
@param letter (str) - The letter indicating the subdirectory within the cemetery directory.
@param namePath (str) - The path to the subdirectory being processed.
@param cemPath (str) - The path to the cemetery directory.
@param counterA (int) - The initial renaming counter.
@param isFirstFile (bool) - Flag indicating if the current file is the first to be processed.
@param start (int) - The file number from which to start renaming.
@param startFlag (bool) - Flag indicating if the start file number has been reached.

@return counter (int) - The updated file counter after renaming.
@return startFlag (bool) - Updated flag indicating if the start condition was met.

@author Mike
'''
def clean(cemetery, letter, namePath, cemPath, counterA, isFirstFile, start, startFlag):
    pdfFiles = sorted(os.listdir(namePath))
    letters = sorted([folder for folder in os.listdir(cemPath) if os.path.isdir(os.path.join(cemPath, folder))])
    try:
        currentLetterIndex = letters.index(letter)
    except ValueError:
        return
    folderBeforeIndex = max(0, currentLetterIndex - 1)
    folderBefore = letters[folderBeforeIndex]
    pdfFilesBefore = sorted([file for file in os.listdir(os.path.join(cemPath, folderBefore)) if file.lower().endswith('.pdf')])
    maxCounter = 0
    for file in pdfFilesBefore:
        match = re.search(r'\d+', file)
        if match:
            number = int(match.group())
            maxCounter = max(maxCounter, number)
    counter = maxCounter + 1
    for x in pdfFiles:
        if counter != start and startFlag:
            isFirstFile = False
            if "a" in x[-5:]:
                pass
            else:
                counter += 1
            continue
        startFlag = False
        if isFirstFile:
            counter = counterA 
        if "a.pdf" in x:
            newName = f"{cemetery}{letter}{counter:05d}a.pdf"
            os.rename(os.path.join(namePath, x), os.path.join(namePath, newName))
            counter -= 1
        elif "b.pdf"  in x:
            newName = f"{cemetery}{letter}{counter:05d}b.pdf"
            os.rename(os.path.join(namePath, x), os.path.join(namePath, newName))
        else:    
            newName = f"{cemetery}{letter}{counter:05d}.pdf"
            os.rename(os.path.join(namePath, x), os.path.join(namePath, newName))
        counter += 1
        isFirstFile = False
    return counter, startFlag


'''
Initiates the cleaning process for image files in all cemeteries, starting from a specific file ID.
Adjusts the starting point for renaming based on the given ID and processes each cemetery found in the base path.

@param id (int) - The file ID from which to start the cleaning process.

@author Mike
'''
def cleanImages(id):
    baseCemPath = r"\\ucclerk\pgmdoc\Veterans\Cemetery"
    cemeterys = [d for d in os.listdir(baseCemPath) if os.path.isdir(os.path.join(baseCemPath, d))]
    initialCount = 1
    start = id
    startFlag = False
    for cemetery in cemeterys:
        if cemetery in ["Jewish", "Misc"]:
            subCemPath = os.path.join(baseCemPath, cemetery)
            subCemeteries = [d for d in os.listdir(subCemPath) if os.path.isdir(os.path.join(subCemPath, d))]
            for subCemetery in subCemeteries:
                print(f"Processing {cemetery} - {subCemetery}")
                initialCount, startFlag = process_cemetery(subCemetery, subCemPath, initialCount, start, startFlag)
        else:
            print(f"Processing {cemetery}")
            initialCount, startFlag = process_cemetery(cemetery, baseCemPath, initialCount, start, startFlag)
            
            
if __name__ == "__main__":
    id = 2 # Proper ID number
    cleanImages(id)