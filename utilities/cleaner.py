import os
import re
import openpyxl
import subprocess
import time
import sys
sys.path.append(r'C:\workspace\veterans')
import microsoftOCR


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
    start = id - 200
    startFlag = True
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
      
        
'''
Decrements the numerical part of file names in a directory and its subdirectories, 
starting from a specified index. This is used to maintain a sequential order after 
removing files. Files with a number greater than or equal to start_index will have 
their number decremented by 1.

@param base_path (str) - The path to the directory containing the files to be renamed.
@param start_index (int) - The file number from which to start decrementing.

@author Mike
'''      
def cleanRedacted(id):
    baseCemPath = r"\\ucclerk\pgmdoc\Veterans\Cemetery - Redacted"
    startIndex = id + 1 # The image file name gets decremented starting at this ID number
    for root, dirs, files in os.walk(baseCemPath):
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
Adjusts hyperlinks in an Excel file that contain numerical IDs, decrementing those IDs by 1
starting from a specified index. This function is used when IDs in a series are modified and 
corresponding hyperlinks need to reflect these changes. Hyperlinks with IDs greater than or 
equal to start_index will have their ID decremented by 1.

@param filePath (str) - The path to the Excel file to be modified.
@param columnLetter (str) - The letter of the column containing hyperlinks to be adjusted.
@param startIndex (int) - The row number from which to start adjusting hyperlinks.

@author Mike
'''     
def cleanHyperlink(id, cemetery):
    excelFilePath = r"\\ucclerk\pgmdoc\Veterans\Veterans.xlsx" 
    columnWithHyperlinks = 'O'  
    startIndex = id # The hyperlink gets decremented by 1 at this cell ID number
    workbook = openpyxl.load_workbook(excelFilePath)
    sheet = workbook[cemetery]
    for row in sheet.iter_rows(min_col=openpyxl.utils.column_index_from_string(columnWithHyperlinks),
                               max_col=openpyxl.utils.column_index_from_string(columnWithHyperlinks),
                               min_row=startIndex): 
        cell = row[0] 
        if cell.hyperlink: 
            url = cell.hyperlink.target
            new_url = re.sub(r'(\d+)', lambda x: f"{int(x.group())-1:05d}" if int(x.group()) >= startIndex else x.group(), url)
            if new_url != url:
                cell.hyperlink.target = new_url
                print(f"Updated hyperlink: {os.path.basename(url)} to {os.path.basename(new_url)}")
    workbook.save(excelFilePath)
    print(f"Hyperlinks updated.")


def cleanDelete(id):
    # baseCemPath = r"\\ucclerk\pgmdoc\Veterans\Cemetery"
    # for dirpath, dirnames, filenames in os.walk(baseCemPath):
    #     for filename in filenames:
    #         if f"{id:05d}" in filename:
    #             file_path = os.path.join(dirpath, filename)
    #             os.remove(file_path)
    #             print(f"{filename} deleted successfully.")
    baseRedacPath = r"\\ucclerk\pgmdoc\Veterans\Cemetery - Redacted"
    for dirpath, dirnames, filenames in os.walk(baseRedacPath):
        for filename in filenames:
            if f"{id:05d}" in filename:
                file_path = os.path.join(dirpath, filename)
                os.remove(file_path)
                print(f"{filename} deleted successfully.")


def open_excel_and_wait():
    filePath = r"\\ucclerk\pgmdoc\Veterans\Veterans.xlsx" 
    try:
        os.startfile(filePath)
    except Exception as e:
        print(f"Error opening file: {e}")
        return
    print("Excel file opened.")
    while True:
        tasklist = subprocess.Popen('tasklist', stdout=subprocess.PIPE)
        output = tasklist.communicate()[0].decode('utf-8')
        if 'EXCEL.EXE' not in output.upper():
            print("Excel file closed.")
            break
        time.sleep(2)
       
       
def adjustImageName(goodID, badID):
    baseCemPath = r"\\ucclerk\pgmdoc\Veterans\Cemetery"
    for dirpath, dirnames, filenames in os.walk(baseCemPath):
        for filename in filenames:
            if f"{goodID:05d}" in filename:
                file_path = os.path.join(dirpath, filename)
                os.rename(file_path, file_path.replace(".pdf", "a.pdf"))
                print(f"{filename} renamed to {filename.replace(".pdf", "a.pdf")}")
            if f"{badID:05d}" in filename:
                file_path = os.path.join(dirpath, filename)
                os.rename(file_path, file_path.replace(f"{badID:05d}.pdf", f"{goodID:05d}b.pdf"))
                print(f"{filename} renamed to {filename.replace(".pdf", "b.pdf")}")
                

       
if __name__ == "__main__":
    cemetery = "Evergreen"
    goodIDs = []
    badIDs = []
    count = 0
    for id in badIDs:
        cleanDelete(id)
        adjustImageName(goodIDs[count], id)
        open_excel_and_wait()
        microsoftOCR.main(True)
        cleanRedacted(id)
        cleanHyperlink(id, cemetery)
        count += 1
    cleanImages(goodIDs[0]-200)
