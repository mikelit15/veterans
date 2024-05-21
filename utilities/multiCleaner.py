import openpyxl
import os
import shutil
import re
import sys
from openpyxl.worksheet.hyperlink import Hyperlink
import duplicates
sys.path.append(r'C:\workspace\veterans')
from microsoftOCR import microsoftOCR


'''
Initiates the file name cleaning process. This function goes through each cemetery folder
and then calls process_cemetery to handle the letter subfolders. The function starts calling 
process_cemetery when the ID matches "start." This is used to significantly speed up
the process of this function, especially for further cemeteries. Special handling is 
required if a cemetery folder contains more cemetery subfolders. This is the case for 
Jewish and Misc folders as they contain more, much smaller, cemeteries.

@param goodID (int) - A file ID used to determine when to start processing
@param redac (str) - The redacted folder name extension
@param redac2 (str) - The redacted file name extension

- Processes all cemeteries in the veternas cemetery folder, or the cemetery redacted
  folder.
- Starts processing file names at goodID - 600
@author Mike
'''
def cleanImages(goodID, redac, redac2):
    baseCemPath = fr"\\ucclerk\pgmdoc\Veterans\Cemetery{redac}"
    cemeterys = [d for d in os.listdir(baseCemPath) if os.path.isdir(os.path.join(baseCemPath, d))]
    initialCount = 1
    start = goodID - 600
    startFlag = True
    for cemetery in cemeterys:
        if cemetery in ["Jewish", "Misc"]:
            subCemPath = os.path.join(baseCemPath, cemetery)
            subCemeteries = [d for d in os.listdir(subCemPath) if os.path.isdir(os.path.join(subCemPath, d))]
            for subCemetery in subCemeteries:
                print(f"Processing {cemetery} - {subCemetery}")
                initialCount, startFlag = process_cemetery(subCemetery, subCemPath, initialCount, start, startFlag, redac2)
        else:
            print(f"Processing {cemetery}")
            initialCount, startFlag = process_cemetery(cemetery, baseCemPath, initialCount, start, startFlag, redac2)
      

'''
Iterates through alphabetically-based letter subfolders within a specified cemetery 
directory, calling clean() to adjust image names for each letter subfolder.

@param cemeteryName (str) - The name of the cemetery being processed.
@param basePath (str) - The root directory path where cemetery folders are located.
@param initialCount (int) - The starting point for file numbering within each alphabetical 
                            folder.
@param start (int) - A specific starting counter value used when deciding whether to process 
                     a file.
@param startFlag (bool) - A flag indicating whether the process should start renaming from 
                          the 'start' parameter.
@param redac2 (str) - The redacted file name extension

@return initialCount (int) - The starting point for file numbering within each alphabetical 
                             folder.
@return startFlag (bool) - A flag indicating whether the process should start renaming from 
                           the 'start' parameter.

- Iterates through A-Z folder, calling clean() for each letter
- Returns updated initialCount and startFlag

@author Mike
'''
def process_cemetery(cemeteryName, basePath, initialCount, start, startFlag, redac2):
    uppercase_alphabet = [chr(i) for i in range(ord('A'), ord('Z') + 1)]
    isFirstFile = True
    for letter in uppercase_alphabet:
        name_path = os.path.join(basePath, cemeteryName, letter)
        try:
            print(f"Processing {cemeteryName} - {letter}")
            initialCount, startFlag = clean(letter, name_path, os.path.join(basePath, cemeteryName), \
                initialCount, isFirstFile, start, startFlag, redac2)
            isFirstFile = False
        except FileNotFoundError:
            continue
    return initialCount, startFlag

'''
Iterates through every PDF file in the given letter subfolder, renaming them to ensure 
sequential ordering and removes extra "a" and "b" redacted files. The function maintaines
continuity and handles special cases, such as the redacted files should not increment the 
count for the numbering sequence.

@param letter (str) - The current letter folder being processed.
@param namePath (str) - The path to the current folder where PDF files are located.
@param cemPath (str) - The path to the cemetery's root directory containing alphabetical 
                       folders.
@param initialCount (int) - The starting numbering for files in the current alphabetical 
                            sequence.
@param isFirstFile (bool) - Indicates if the current file is the first in its alphabetical 
                            sequence.
@param start (int) - The numbering value at which to begin processing files, if startFlag 
                     is True.
@param startFlag (bool) - A flag that, if True, delays file processing until the 'start' 
                          value is reached.
@param redac2 (str) - The redacted file name extension

@return counter (int) - The next starting count for continued processing.
@return startFlag (bool) - State of startFlag for continued processing.

- Creates a list for every PDF in the subfolder
- Finds the index of the last PDF file, which is in the previous letter subfolder
- Counter is set to that last index
- Adjusts the image file name to match the current index

@author Mike
'''
def clean(letter, namePath, cemPath, initialCount, isFirstFile, start, startFlag, redac2):
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
            counter = initialCount 
        if redac2:
            if "a redacted.pdf" in x:
                counter -= 1
                os.remove(os.path.join(namePath, x))
                print(f"Removed {x}")
            elif "b redacted.pdf" in x:
                counter -= 1
                os.remove(os.path.join(namePath, x))
                print(f"Removed {x}")
            else:
                newName = re.sub(r'\d+', f"{counter:05d}", x)
                os.rename(os.path.join(namePath, x), os.path.join(namePath, newName))
        else:
            if "a.pdf" in x:
                newName = re.sub(r'\d+', f"{counter:05d}", x)
                os.rename(os.path.join(namePath, x), os.path.join(namePath, newName))
                counter -= 1
            elif "b.pdf"  in x:
                newName = re.sub(r'\d+', f"{counter:05d}", x)
                os.rename(os.path.join(namePath, x), os.path.join(namePath, newName))
            else:   
                newName = re.sub(r'\d+', f"{counter:05d}", x)
                os.rename(os.path.join(namePath, x), os.path.join(namePath, newName))
        counter += 1
        isFirstFile = False
    return counter, startFlag


'''
Adjusts the filenames and paths for image files associated with goodID and badID. GoodID 
receives a "a.pdf" suffix and badID receives a "b.pdf" suffix. badID is also adjusted to 
match the same letter and folder as the goodID if needed. The data in goodRow is then cleared
and the hyperlink forcefully cleared to allow for microsoftOCR to locate and append the new
merged data.

@param goodID (int): The correct identifier associated with the proper record.
@param badID (int): The incorrect identifier that needs to be adjusted to match the good ID.
@param goodRow (int): The row in the Excel spreadsheet where the good ID is located.

- Searches for files matching goodID and badID
- Adds the suffix, "a.pdf", to goodID
- Adds the suffix, "b.pdf", to badID. Changes letter in file name and moves to correct
  letter subfolder is necessary.
- Clears the data in Excel row associated with the goodID

@author Mike
'''
def adjustImageName(goodID, badID, goodRow):
    baseCemPath = r"\\ucclerk\pgmdoc\Veterans\Cemetery"
    goodIDFound = False
    badIDFound = False
    goodIDFilePath = ""
    badIDFilePath = ""
    for dirpath, dirnames, filenames in os.walk(baseCemPath):
        for filename in filenames:
            if f"{goodID:05d}" in filename:
                goodIDFilePath = os.path.join(dirpath, filename)
                goodIDFound = True
            elif f"{badID:05d}" in filename:
                badIDFilePath = os.path.join(dirpath, filename)
                badIDFound = True
            if goodIDFound and badIDFound:
                break
        if goodIDFound and badIDFound:
            break
    if goodIDFound and badIDFound:
        goodIDPrefix = os.path.basename(os.path.dirname(goodIDFilePath))
        badIDPrefix = os.path.basename(os.path.dirname(badIDFilePath))
        newBadIDFilename = os.path.basename(badIDFilePath).replace(f"{badID:05d}", f"{goodID:05d}b").replace(badIDPrefix, goodIDPrefix)
        newBadIDFilePath = os.path.join(os.path.dirname(goodIDFilePath), newBadIDFilename)
        if not "a.pdf" in os.path.basename(goodIDFilePath):
            newGoodIDFilename = os.path.basename(goodIDFilePath).replace(".pdf", "a.pdf")
            newGoodIDFilePath = os.path.join(os.path.dirname(goodIDFilePath), newGoodIDFilename)
            os.rename(goodIDFilePath, newGoodIDFilePath)
            print(f"{os.path.basename(goodIDFilePath)} renamed to {newGoodIDFilename}")
        shutil.move(badIDFilePath, newBadIDFilePath)
        print(f"{os.path.basename(badIDFilePath)} moved and renamed to {newBadIDFilename}")
    else:
        print("GoodID or BadID file not found. Please check the IDs and directories.")
        quit()
    start_column = 'A'
    end_column = 'O'
    start_column_index = openpyxl.utils.column_index_from_string(start_column)
    end_column_index = openpyxl.utils.column_index_from_string(end_column)
    for col_index in range(start_column_index, end_column_index + 1):
        worksheet.cell(row=goodRow, column=col_index).value = None
    worksheet[f'O{goodRow}'].hyperlink = None
    print(f"Record {goodID} data from row {goodRow} cleared successfully.")


'''
Deletes redacted image file associated with a badID and updates an Excel spreadsheet to 
remove the corresponding row. Due to a bug with the openpyxl module, hyperlinks do not 
move up a row with the rest of the cell data when calling .delete_rows(). So the hyperlinks 
are forced up one row and the hyperlink id and ref are adjusted to make the hyperlinks 
active in order to adjust them automatically later in the process. 

@param cemetery (str): The name of the cemetery where the file is located.
@param badID (int): The ID of the file to be deleted.
@param badRow (int): The row in the Excel spreadsheet corresponding to badID.

- Deletes the redacted file with badID ID number
- Deletes the Excel row corresponding to the bad ID
- Hyperlinks are adjusted starting from badRow

@author Mike
'''
def cleanDelete(cemetery, badID, badRow):
    baseRedacPath = r"\\ucclerk\pgmdoc\Veterans\Cemetery - Redacted"
    for dirpath, dirnames, filenames in os.walk(baseRedacPath):
        for filename in filenames:
            if f"{badID:05d}" in filename:
                file_path = os.path.join(dirpath, filename)
                os.remove(file_path)
                print(f"\nRedacted file, {filename}, deleted successfully.")
    worksheet = workbook[cemetery]
    worksheet.delete_rows(badRow)
    print(f"Row {badRow} deleted successfully.")
    for row in range(badRow, worksheet.max_row + 1):
        next_row = row + 1
        cell_ref = f'O{row}'
        next_cell_ref = f'O{next_row}'
        if next_row <= worksheet.max_row and worksheet[next_cell_ref].hyperlink:
            temp_hyperlink = worksheet[next_cell_ref].hyperlink
            worksheet[cell_ref].hyperlink = Hyperlink(ref=cell_ref, target=temp_hyperlink.target, display=temp_hyperlink.display)
            worksheet[cell_ref].value = worksheet[next_cell_ref].value
            print(f"Hyperlink and value from row {next_row} moved to row {row}.")
    if worksheet[f'O{worksheet.max_row - 1}'].hyperlink:
        last_row = worksheet.max_row
        print(f"Adjusting the hyperlink in the last row, row {last_row}.")
        second_last_hyperlink = worksheet[f'O{last_row - 1}'].hyperlink
        new_target = re.sub(r'(\d+)(?=%\d+)', lambda m: f"{int(m.group(1)) + 1:05d}", second_last_hyperlink.target)
        worksheet[f'O{last_row}'].hyperlink = Hyperlink(ref=f'O{last_row}', target=new_target, display=second_last_hyperlink.display)
        print(f"Updated hyperlink in row {last_row} to new target, {new_target}.\n")


'''
Updates the record ID numbers in column A to fix fill any gaps and continue going in 
sequential order. Updates hyperlink references in an Excel spreadsheet starting from a 
specified index. This function adjusts the numeric ID within each hyperlink to ensure 
they accurately reflect the current record IDs following modifications.

@param startIndex (int): The row index from which to start updating hyperlinks in the 
                         spreadsheet.

- Adjusts record ID in column A 
- Takes the hyperlink in the cell and changes the ID portion of the file name to match 
  the ID in column A for each row starting at startIndex

@author Mike
'''
def cleanHyperlinks(cemetery, startIndex):
    base_path = fr'\\ucclerk\pgmdoc\Veterans\Cemetery\{cemetery}'
    file_directory_map = {}
    for dirpath, dirnames, filenames in os.walk(base_path):
        for filename in filenames:
            file_id = ''.join(filter(str.isdigit, filename))[:5]
            if file_id:
                file_directory_map[file_id] = dirpath
    new_id = worksheet[f'A{startIndex}'].value
    for row in range(startIndex, worksheet.max_row + 1):
        worksheet[f'A{row}'].value = new_id
        cell_ref = f'O{row}'
        if worksheet[cell_ref].hyperlink:
            orig_target = worksheet[cell_ref].hyperlink.target.replace("%20", " ")
            formatted_id = f"{new_id:05d}"
            folder_name = file_directory_map.get(formatted_id, "UnknownFolder")[-1]
            modified_string = re.sub(fr'(\\)([A-Z])(\\)', f'\\1{folder_name}\\3', orig_target)
            modified_string = re.sub(fr'({cemetery}[A-Z])\d{{5}}', f'{cemetery}{folder_name}{formatted_id}', modified_string)
            worksheet[cell_ref].hyperlink = Hyperlink(ref=cell_ref, target=modified_string, display="PDF Image")
            worksheet[cell_ref].value = "PDF Image"
            print(f"Updated hyperlink from {orig_target} to {modified_string} in row {row}.")
        new_id += 1
        
def main(cemetery, goodIDs, badIDs):
    excelFilePath = r"\\ucclerk\pgmdoc\Veterans\Veterans.xlsx"
    global workbook
    workbook = openpyxl.load_workbook(excelFilePath)
    global worksheet 
    worksheet = workbook[cemetery]
    for x in range(0, len(goodIDs)):
        workbook = openpyxl.load_workbook(excelFilePath)
        worksheet = workbook[cemetery]
        for row in range(1, worksheet.max_row + 1):
            if worksheet[f'A{row}'].value == goodIDs[x]:
                goodRow = row
            if worksheet[f'A{row}'].value == badIDs[x]:
                badRow = row
        letter = worksheet[f'B{goodRow}'].value[0]
        adjustImageName(goodIDs[x], badIDs[x], goodRow)
        workbook.save(excelFilePath)
        microsoftOCR.main(True, cemetery, letter)
        workbook = openpyxl.load_workbook(excelFilePath)
        worksheet = workbook[cemetery]
        cleanDelete(cemetery, badIDs[x], badRow)
        workbook.save(excelFilePath)
    mini = min(goodIDs)
    if min(badIDs) < mini:
        mini = min(badIDs) - 1
    cleanImages(mini, "", "")
    cleanImages(mini, " - Redacted", " redacted")
    workbook = openpyxl.load_workbook(excelFilePath)
    worksheet = workbook[cemetery]
    for row in range(1, worksheet.max_row + 1):
        if worksheet[f'A{row}'].value == mini:
            startIndex = row
    cleanHyperlinks(cemetery, startIndex)
    workbook.save(excelFilePath)
    print("\n")
    duplicates.main(cemetery)

if __name__ == "__main__": 
    cemetery = "Graceland"
    goodIDs = [9382, 9579, 9558, 9761, 9754, 9768, 10023, 10041, 9696] 
    badIDs = [9405, 9583, 9588, 9764, 9767, 9780, 10030, 10045, 10194]
    main(cemetery, goodIDs, badIDs)