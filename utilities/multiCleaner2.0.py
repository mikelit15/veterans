import openpyxl
import os
import shutil
import re
import time
from PyPDF2 import PdfMerger
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.styles import PatternFill

'''
Initiates the file name cleaning process. This function goes through each cemetery folder
and then calls processCemetery to handle the letter subfolders. The function starts calling 
processCemetery when the ID matches "start." This is used to significantly speed up
the process of this function, especially for further cemeteries. Special handling is 
required if a cemetery folder contains more cemetery subfolders. This is the case for 
Jewish and Misc folders as they contain more, much smaller, cemeteries.

@param fullClean (bool) - A flag indicating whether the full cleaning process should be done
@param goodID (int) - A file ID used to determine when to start processing
@param redac (str) - The redacted folder name extension
@param redac2 (str) - The redacted file name extension

- Processes all cemeteries in the veterans' cemetery folder or the cemetery redacted
  folder.
- Starts processing file names at goodID - 600

@author Mike
'''
def cleanImages(fullClean, goodID, redac, redac2):
    baseCemPath = fr"\\ucclerk\pgmdoc\Veterans\Cemetery{redac}"
    cemeterys = [d for d in os.listdir(baseCemPath) if os.path.isdir(os.path.join(baseCemPath, d))]
    initialCount = 1
    start = goodID
    startFlag = not fullClean
    for cemetery in cemeterys:
        if cemetery in [f"Jewish{redac}", f"Misc{redac}"]:
            subCemPath = os.path.join(baseCemPath, cemetery)
            subCemeteries = [d for d in os.listdir(subCemPath) if os.path.isdir(os.path.join(subCemPath, d))]
            for subCemetery in subCemeteries:
                print(f"Processing {cemetery} - {subCemetery}")
                initialCount, startFlag = processCemetery(subCemetery, subCemPath, initialCount, start, startFlag, redac, redac2)
        else:
            print(f"Processing {cemetery}")
            initialCount, startFlag = processCemetery(cemetery, baseCemPath, initialCount, start, startFlag, redac, redac2)
      
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
@param redac (str) - The redacted folder name extension
@param redac2 (str) - The redacted file name extension

@return initialCount (int) - The starting point for file numbering within each alphabetical 
                             folder.
@return startFlag (bool) - A flag indicating whether the process should start renaming from 
                           the 'start' parameter.

- Iterates through A-Z folders, calling clean() for each letter
- Returns updated initialCount and startFlag

@author Mike
'''
def processCemetery(cemeteryName, basePath, initialCount, start, startFlag, redac, redac2):
    uppercaseAlphabet = [chr(i) for i in range(ord('A'), ord('Z') + 1)]
    isFirstFile = True
    for letter in uppercaseAlphabet:
        namePath = os.path.join(basePath, cemeteryName, letter)
        try:
            print(f"Processing {cemeteryName} - {letter}")
            initialCount, startFlag = clean(letter, namePath, os.path.join(basePath, cemeteryName), \
                initialCount, isFirstFile, start, startFlag, redac, redac2)
            isFirstFile = False
        except FileNotFoundError:
            continue
    return initialCount, startFlag

'''
Iterates through every PDF file in the given letter subfolder, renaming them to ensure 
sequential ordering and removes extra "a" and "b" redacted files. The function maintains
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
@param redac (str) - The redacted folder name extension
@param redac2 (str) - The redacted file name extension

@return counter (int) - The next starting count for continued processing.
@return startFlag (bool) - State of startFlag for continued processing.

- Creates a list for every PDF in the subfolder
- Finds the index of the last PDF file, which is in the previous letter subfolder
- Counter is set to that last index
- Adjusts the image file name to match the current index

@author Mike
'''
def clean(letter, namePath, cemPath, initialCount, isFirstFile, start, startFlag, redac, redac2):
    global cemSet, miscSet, jewishSet
    pdfFiles = sorted(os.listdir(namePath))
    letters = sorted([folder for folder in os.listdir(cemPath) if os.path.isdir(os.path.join(cemPath, folder))])
    try:
        currentLetterIndex = letters.index(letter)
    except ValueError:
        return initialCount, startFlag
    folderBeforeIndex = max(0, currentLetterIndex - 1)
    folderBefore = letters[folderBeforeIndex]
    
    if currentLetterIndex == 0 or folderBefore == letter: 
        pdfFilesBefore = []
    else:
        pdfFilesBefore = sorted([file for file in os.listdir(os.path.join(cemPath, folderBefore)) if file.lower().endswith('.pdf')])

    if letter == "A" and folderBefore == "A":
        if os.path.basename(os.path.normpath(cemPath)).replace(" - Redacted", "") in cemSet:
            cemeteryFolders = sorted(os.listdir(os.path.dirname(cemPath)))
            currentCemeteryIndex = cemeteryFolders.index(os.path.basename(os.path.normpath(cemPath)))
        elif os.path.basename(os.path.normpath(cemPath)).replace(" - Redacted", "") in jewishSet:
            cemPath = os.path.dirname(cemPath)
            cemeteryFolders = sorted(os.listdir(os.path.dirname(cemPath)))
            currentCemeteryIndex = cemeteryFolders.index(f"Jewish{redac}")
        elif os.path.basename(os.path.normpath(cemPath)).replace(" - Redacted", "") in miscSet:
            cemPath = os.path.dirname(cemPath)
            cemeteryFolders = sorted(os.listdir(os.path.dirname(cemPath)))
            currentCemeteryIndex = cemeteryFolders.index(f"Misc{redac}")
        if currentCemeteryIndex > 0:
            previousCemeteryIndex = currentCemeteryIndex - 1
            previousCemeteryFolder = cemeteryFolders[previousCemeteryIndex]
            previousCemeteryPath = os.path.join(os.path.dirname(cemPath), previousCemeteryFolder)
            previousCemeteryLetters = sorted([folder for folder in os.listdir(previousCemeteryPath) if os.path.isdir(os.path.join(previousCemeteryPath, folder))])
            folderBefore = previousCemeteryLetters[-1]
            pdfFilesBefore = sorted([file for file in os.listdir(os.path.join(previousCemeteryPath, folderBefore)) if file.lower().endswith('.pdf')])
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
            elif "b.pdf" in x:
                newName = re.sub(r'\d+', f"{counter:05d}", x)
                os.rename(os.path.join(namePath, x), os.path.join(namePath, newName))
            else:
                newName = re.sub(r'\d+', f"{counter:05d}", x)
                os.rename(os.path.join(namePath, x), os.path.join(namePath, newName))
        counter += 1
        isFirstFile = False

    return counter, startFlag

'''
Updates the record ID numbers in column A to fix fill any gaps and continue going in 
sequential order. Updates hyperlink references in an Excel spreadsheet starting from a 
specified index. This function adjusts the numeric ID within each hyperlink to ensure 
they accurately reflect the current record IDs following modifications.

@param startIndex (int): The row index from which to start updating hyperlinks in the 
                         spreadsheet.
@param newID (int): The new record ID to assign starting from startIndex.
@param fileDirectoryMap (dict): A dictionary mapping file IDs to their respective directories.

- Adjusts record ID in column A 
- Takes the hyperlink in the cell and changes the ID portion of the file name to match 
  the ID in column A for each row starting at startIndex

@return newID (int): The updated record ID after processing.

@author Mike
'''
def cleanHyperlinks(badWorksheet, startIndex, newID, fileDirectoryMap):
    for row in range(startIndex, badWorksheet.max_row + 1):
        badWorksheet[f'A{row}'].value = newID
        cellRef = f'O{row}'
        if badWorksheet[cellRef].hyperlink:
            origTarget = badWorksheet[cellRef].hyperlink.target.replace("%20", " ")
            formattedID = f"{newID:05d}"
            try:
                folderName = fileDirectoryMap.get(formattedID)[-1]
                modifiedString = re.sub(r'([\\/])[A-Z]([\\/])', f'\\1{folderName}\\2', origTarget)
                if badWorksheet.title in ["Jewish", "Misc"]:
                    match = re.search(r'[\\/][A-Z][\\/](.*?)([A-Z])\d+', origTarget)
                    extracted_value = match.group(1)  
                    modifiedString = re.sub(fr'({extracted_value}[A-Z])\d{{5}}', f'{extracted_value}{folderName}{formattedID}', modifiedString)
                else:
                    modifiedString = re.sub(fr'({badWorksheet.title}[A-Z])\d{{5}}', f'{badWorksheet.title}{folderName}{formattedID}', modifiedString)
                badWorksheet[cellRef].hyperlink = Hyperlink(ref=cellRef, target=modifiedString, display="PDF Image")
                badWorksheet[cellRef].value = "PDF Image"
            except Exception:
                pass
            print(f"Updated hyperlink from {origTarget} to {modifiedString} in row {row}.")
        newID += 1
    return newID


'''
Compares the letters in the folder path and file name of hyperlinks within an Excel 
workbook to ensure they match. The function identifies rows where the folder letter 
and file letter in the hyperlink do not match and prints these discrepancies.

@param workbookPath (str): The path to the Excel workbook to be processed.

- Iterates through each worksheet and each row starting from row 2.
- Extracts the hyperlink from column O (15th column) and compares the letters in the 
  folder path and file name.
- Prints hyperlinks where the letters do not match.

@author Mike
'''
def compareHyperlinkLetters(workbookPath):
    print("Checking hyperlinks...")
    workbook = openpyxl.load_workbook(workbookPath)
    pattern = re.compile(
        r'Cemetery - Redacted[\\/][^\\/]+ - Redacted[\\/](?P<folderLetter>[A-Z])[\\/][^\\/]+(?P<fileLetter>[A-Z])\d+ redacted\.pdf'
    )
    for sheetName in workbook.sheetnames:
        worksheet = workbook[sheetName]
        flag = False
        for sheetName in workbook.sheetnames:
            worksheet = workbook[sheetName]
            for row in worksheet.iter_rows(min_row=2, max_col=worksheet.max_column):
                cell = row[14]  
                if cell.hyperlink:
                    hyperlink = cell.hyperlink.target.replace("%20", " ")
                    match = pattern.search(hyperlink)
                    if match:
                        folderLetter = match.group(1)
                        fileLetter = match.group(2)
                        if folderLetter == fileLetter:
                            pass
                        else:
                            flag = True
                            print(f"Not matching: {hyperlink} in row {cell.row}\n")
        if not flag:
            print("All hyperlinks are correct.")

'''
Main function to process and update Excel sheets containing good and bad IDs. It adjusts
image names, updates hyperlinks, cleans up records, and compares hyperlink letters to 
ensure consistency.

@param fullClean (bool): A flag indicating whether the full cleaning process should be done
@param goodIDs (list): List of good IDs to be processed.
@param badIDs (list): List of bad IDs to be processed.

- Loads the Excel workbook and identifies sheets and rows containing good and bad IDs.
- Adjusts image names and updates records based on good and bad IDs.
- Cleans and deletes records with bad IDs.
- Processes images to update records and hyperlinks.
- Compares hyperlink letters to ensure consistency.

The function uses various helper functions such as `adjustImageName`, `microsoftOCR.main`,
`cleanDelete`, `cleanImages`, `cleanHyperlinks`, and `compareHyperlinkLetters`.

The Excel file is saved after each major operation to ensure data consistency.

@author Mike
'''      
def main(fullClean, goodIDs, badIDs):
    cemeterys = [x for x in os.listdir(r"\\ucclerk\pgmdoc\Veterans\Cemetery")]
    miscs = [x for x in os.listdir(r"\\ucclerk\pgmdoc\Veterans\Cemetery\Misc")]
    jewishs = [x for x in os.listdir(r"\\ucclerk\pgmdoc\Veterans\Cemetery\Jewish")]
    global cemSet, miscSet, jewishSet
    cemSet, miscSet, jewishSet = set(cemeterys), set(miscs), set(jewishs)
    excelFilePath = r"\\ucclerk\pgmdoc\Veterans\Veterans.xlsx"
    basePath = fr'\\ucclerk\pgmdoc\Veterans\Cemetery'
    redactedBasePath = fr'\\ucclerk\pgmdoc\Veterans\Cemetery - redacted'
    workbook = openpyxl.load_workbook(excelFilePath)
    for goodID, badID in zip(goodIDs, badIDs):
        print(f"Processing Good ID: {goodID}, Bad ID: {badID}...")
        time.sleep(5)
        goodRow, badRow = None, None
        goodSheet, badSheet = None, None
        goodWorkSheet, badWorksheet = None, None

        # Find goodID and badID in sheets
        for sheet in workbook.sheetnames:
            worksheet = workbook[sheet]
            for row in range(2, worksheet.max_row + 1):
                if worksheet[f'A{row}'].value == goodID:
                    goodRow = row
                    goodSheet = sheet
                    print(f"Good ID {goodID} found in sheet: {goodSheet}, row: {goodRow}")
                if worksheet[f'A{row}'].value == badID:
                    badRow = row
                    badSheet = sheet
                    print(f"Bad ID {badID} found in sheet: {badSheet}, row: {badRow}")
                if goodRow and badRow:
                    break
            if goodRow and badRow:
                break
        if goodRow is None or badRow is None:
            print(f"Failed to find either Good ID {goodID} or Bad ID {badID}, skipping...")
            continue
        
        # Assign worksheets
        goodWorkSheet = workbook[goodSheet]
        badWorksheet = workbook[badSheet]

        # Get row data for both goodID and badID
        vals1 = [goodWorkSheet.cell(row=goodRow, column=col).value for col in range(2, 16)]
        vals2 = [badWorksheet.cell(row=badRow, column=col).value for col in range(2, 16)]
        print(f"Retrieved values from Good ID row {goodRow}:\n", end="")
        print(' | '.join(str(value) for value in vals1 if value is not None))
        print(f"Retrieved values from Bad ID row {badRow}:\n", end="")
        print(' | '.join(str(value) for value in vals2 if value is not None))
        time.sleep(8)
        
        # Clear the goodRow data (starting from column B) and delete badRow
        for col in range(2, goodWorkSheet.max_column + 1):
            goodWorkSheet.cell(row=goodRow, column=col).value = None
        print(f"Cleared data from Good ID row {goodRow}.")
        badWorksheet.delete_rows(badRow)
        print(f"Row {badRow} deleted successfully.")
        for row in range(badRow, badWorksheet.max_row + 1):
            nextRow = row + 1
            cellRef = f'O{row}'
            nextCellRef = f'O{nextRow}'
            if nextRow <= badWorksheet.max_row and badWorksheet[nextCellRef].hyperlink:
                tempHyperlink = badWorksheet[nextCellRef].hyperlink
                badWorksheet[cellRef].hyperlink = Hyperlink(ref=cellRef, target=tempHyperlink.target, display=tempHyperlink.display)
                badWorksheet[cellRef].value = badWorksheet[nextCellRef].value
                print(f"Hyperlink and value from row {nextRow} moved to row {row}.")
        if badWorksheet[f'O{badWorksheet.max_row - 1}'].hyperlink:
            lastRow = badWorksheet.max_row
            print(f"Adjusting the hyperlink in the last row, row {lastRow}.")
            secondLastHyperlink = badWorksheet[f'O{lastRow - 1}'].hyperlink
            newTarget = re.sub(r'(\d+)(?=%\d+)', lambda m: f"{int(m.group(1)) + 1:05d}", secondLastHyperlink.target)
            badWorksheet[f'O{lastRow}'].hyperlink = Hyperlink(ref=f'O{lastRow}', target=newTarget, display=secondLastHyperlink.display)
            print(f"Updated hyperlink in row {lastRow} to new target, {newTarget}.\n")
        
        # Merge values but keep goodID's hyperlink and other values for columns A and O
        mergedArray = []
        for col in range(len(vals1)):
            if col == 0:  # Column A (1st column, Good ID)
                mergedArray.append(vals1[col])
            elif col == 14:  # Column O (15th column, hyperlink)
                mergedArray.append(vals1[col])
            else:
                mergedArray.append(max(vals1[col], vals2[col], key=lambda x: len(str(x)) if x else 0))
        print(f"Merged values into Good ID row {goodRow}.")        
        counter = 2
        for value in mergedArray:
            goodWorkSheet.cell(row=goodRow, column=counter, value=value)
            counter += 1

        # Highlight cells in GoodRow, columns B - N and P
        highlight_fill = PatternFill(start_color="001AFF", end_color="001AFF", fill_type="solid")
        for col in range(2, 15):  # Columns B (2) to N (14)
            goodWorkSheet.cell(row=goodRow, column=col).fill = highlight_fill
        # Highlight column P (16)
        goodWorkSheet.cell(row=goodRow, column=16).fill = highlight_fill
        print(f"Highlighted cells in GoodRow {goodRow}, columns B - N and P with color #001AFF.")
        
        workbook.save(excelFilePath)
        
        # File handling: Apply the suffix "a" and "b" to goodFilePath and badFilePath and move if necessary
        goodIDFilePath, badIDFilePath = None, None
        for dirPath, _, fileNames in os.walk(basePath):
            for fileName in fileNames:
                if f"{goodID:05d}" in fileName:
                    goodIDFilePath = os.path.join(dirPath, fileName)
                if f"{badID:05d}" in fileName:
                    badIDFilePath = os.path.join(dirPath, fileName)
                if goodIDFilePath and badIDFilePath:
                    break
        if goodIDFilePath and badIDFilePath:
            goodIDPrefix = os.path.basename(os.path.dirname(goodIDFilePath))
            badIDPrefix = os.path.basename(os.path.dirname(badIDFilePath))
            newBadIDFilename = os.path.basename(goodIDFilePath).replace(f"{goodID:05d}", f"{goodID:05d}b").replace(badIDPrefix, goodIDPrefix)
            newBadIDFilePath = os.path.join(os.path.dirname(goodIDFilePath), newBadIDFilename)
            if not "a.pdf" in os.path.basename(goodIDFilePath):
                newGoodIDFilename = os.path.basename(goodIDFilePath).replace(".pdf", "a.pdf")
                newGoodIDFilePath = os.path.join(os.path.dirname(goodIDFilePath), newGoodIDFilename)
                os.rename(goodIDFilePath, newGoodIDFilePath)
                print(f"{os.path.basename(goodIDFilePath)} renamed to {newGoodIDFilename}")
            shutil.move(badIDFilePath, newBadIDFilePath)
            print(f"{os.path.basename(badIDFilePath)} moved and renamed to {newBadIDFilename}")

        # Redacted file handling: Combine redacted versions of the good and bad files
        redactedGoodFilePath, redactedBadFilePath = None, None
        for dirPath, _, fileNames in os.walk(redactedBasePath):
            for fileName in fileNames:
                if f"{goodID:05d}" in fileName:
                    redactedGoodFilePath = os.path.join(dirPath, fileName)
                if f"{badID:05d}" in fileName:
                    redactedBadFilePath = os.path.join(dirPath, fileName)
                if redactedGoodFilePath and redactedBadFilePath:
                    break
        if redactedGoodFilePath and redactedBadFilePath:
            print(f"Combining redacted files for Good ID {goodID} and Bad ID {badID}...")
            merger = PdfMerger()
            with open(redactedGoodFilePath, 'rb') as f1, open(redactedBadFilePath, 'rb') as f2:
                merger.append(f1)
                merger.append(f2)
                tempRedactedFilePath = redactedGoodFilePath.replace(".pdf", "_temp.pdf")
                with open(tempRedactedFilePath, 'wb') as tempFile:
                    merger.write(tempFile)
            merger.close()
            os.remove(redactedGoodFilePath)
            os.remove(redactedBadFilePath)
            os.rename(tempRedactedFilePath, redactedGoodFilePath)
            print(f"Redacted files combined and renamed back to {redactedGoodFilePath}\n")

    cleanImages(fullClean, 1, "", "")
    cleanImages(fullClean, 1, " - Redacted", " redacted")
    fileDirectoryMap = {}
    for dirPath, _, fileNames in os.walk(basePath):
        for fileName in fileNames:
            fileID = ''.join(filter(str.isdigit, fileName))[:5]
            if fileID:
                fileDirectoryMap[fileID] = dirPath
    newID = 1
    for sheetName in workbook.sheetnames:
        badWorksheet = workbook[sheetName]
        newID = cleanHyperlinks(badWorksheet, 2, newID, fileDirectoryMap)
    workbook.save(excelFilePath)
    compareHyperlinkLetters(excelFilePath)

if __name__ == "__main__": 
    fullClean = True
    # A's:
    # goodIDs = [11870, 11873, 15623, 11876, 11886, 16761, 22198, 31093, 27977, 15627, 35311, 35326, 16814, 11917, 11921, 27150, 35362, 16843, 27978, 23874,
    #            15636, 31138, 28034, 16874, 15006, 15651, 35607, 35646, 169, 35651, 31181, 11998, 11212, 42474, 35687] 
    # 
    # badIDs = [13331, 36, 16755, 27144, 54, 16764, 35462, 35296, 27999, 16791, 5101, 8103, 16820, 23873, 11923, 68, 28023, 16853, 28027, 14997, 16858, 14998,
    #           31082, 35389, 26152, 16918, 11197, 31249, 15008, 11205, 31255, 12005, 16969, 42477, 12016]
    # 
    # B's:
    # goodIDs = [42440, 41387, 16966, 15656, 28097, 214, 16999, 31241, 31246, 17010, 230, 17033, 31263, 35695, 244, 15670, 28101, 256, 35746, 31310, 35779, 
    #            35777, 288, 5289, 31332, 17140, 324, 28164, 25415, 31355, 41980, 35852, 26173, 12239, 23268, 42514, 26179, 35909, 15709, 28207, 42536, 35928, 
    #            31429, 28219, 418, 35938, 8459, 31460, 31435, 429, 430, 434, 31490, 15719, 31511, 28242, 17280, 23276, 12237, 28256, 17270, 26195, 26198, 
    #            17287, 12235, 31547, 12218, 31549, 31564, 26203, 493, 496, 31573]
    # 
    # badIDs = [42442, 14996, 16981, 16987, 35679, 14999, 17000, 35701, 15000, 17027, 15002, 17035, 15003, 5234, 15007, 17047, 28114, 15008, 4915, 31311, 28148,
    #           12079, 15011, 15012, 31551, 4893, 15013, 15015, 15014, 15016, 31368, 28179, 15019, 12240, 373, 42530, 15026, 41411, 17212, 15028, 15029, 445, 
    #           31431, 28222, 15036, 31456, 15038, 28225, 31461, 15039, 15040, 15041, 28233, 17246, 31553, 28246, 17307, 23978, 12238, 15044, 17284, 15045, 
    #           15046, 17329, 12236, 31559, 12219, 31562, 31565, 15049, 15050, 15052, 12234]
    # 
    # C & D's:
    # goodIDs = [12235, 26163, 27163, 31474, 15700, 17345, 42529, 42530, 42531, 26180, 28252, 31582, 624, 28267, 28273, 17431, 25413, 28293, 11274, 42535, 5508,
    #            15029, 42600, 23812, 5523, 682, 36148, 693, 28332, 15737, 704, 28333, 5545, 15738, 36204, 710, 17499, 36212, 31698, 42631, 31701, 753, 754, 755, 
    #            31725, 36261, 17580, 17581, 28379, 23814, 26229, 12475, 26238, 776, 790, 15758, 809, 11305, 41978, 36323, 31540, 36351, 12480, 15766, 26257, 17666,
    #            31837, 36422, 12507, 930, 17775, 17776, 15783, 12559, 15778, 27235, 15076, 36448, 26275, 17771, 42684, 12549, 36503, 36497, 15078, 17853, 12554,
    #            42753, 17882, 36523, 31912, 31917, 25461, 31927, 28565, 1052, 5802, 26326, 36601, 36603, 31951, 15806, 12630, 23301, 1075, 26331, 36644, 26340]
    # 
    # badIDs = [12237, 15014, 41359, 31531, 15710, 17351, 42548, 42550, 42551, 15019, 15020, 15022, 15021, 850, 28276, 17535, 15025, 15024, 31538, 42589, 36132, 
    #           23811, 42602, 24207, 15030, 15032, 36149, 15036, 31651, 17494, 15037, 15038, 15039, 17513, 31667, 15040, 17518, 5560, 28356, 42637, 32042, 15045,
    #           15046, 15047, 28374, 5581, 17584, 17585, 28389, 15050, 15051, 12423, 15055, 15056, 15058, 17627, 15060, 27220, 15061, 31781, 33333, 28436, 36367,
    #           17679, 15065, 17687, 28470, 31849, 26265, 15073, 17789, 17783, 17794, 35331, 17735, 15075, 23718, 31884, 15077, 17773, 42728, 12552, 5724, 993, 
    #           23721, 17855, 12591, 4977, 17888, 4849, 31914, 15080, 15082, 15084, 15088, 15091, 15093, 15094, 28578, 28579, 15096, 17946, 15097, 1070, 15098, 
    #           15099, 31973, 15106]
    #
    # E, F, G's:
    # goodIDs = [15617, 15633, 15650, 12652, 1152, 31910, 34388, 36599, 17982, 31944, 28557, 1216, 28562, 11373, 36642, 26276, 26282, 25424, 31992, 27200, 22668,
    #            18102, 42710, 41925, 32016, 11384, 26305, 32022, 32032, 32034, 26311, 36758, 12752, 36764, 32043, 1325, 12768, 41932, 26326, 12770, 6003, 6004,
    #            32088, 1382, 15825, 1392, 12800, 26341, 36866, 28675, 6033, 18282, 42762, 36892, 32128, 41353, 4831, 1455, 36887, 28710, 28716, 1476, 32158, 28724,
    #            36946, 32170, 15849, 6083, 1491, 26368, 18395, 15852, 22715, 26378, 18463, 15862, 1526, 37067, 32279, 18518, 9171, 12904]
    
    # badIDs = [17112, 17194, 842, 36564, 15054, 28100, 34428, 31938, 17984, 31946, 31947, 15057, 28564, 28577, 28579, 15060, 15061, 15063, 32045, 5938, 22671,
    #           18105, 42720, 15069, 15071, 24756, 15073, 32026, 15075, 28629, 15076, 28630, 4848, 32042, 32046, 15077, 15079, 41345, 15080, 12773, 15082, 15084, 
    #           15086, 15087, 18251, 15088, 12808, 15089, 32110, 28682, 15090, 18290, 28684, 1435, 28692, 41376, 5067, 15094, 36919, 28713, 28718, 15098, 15100,
    #           15101, 6075, 36953, 18379, 15104, 36966, 15105, 18401, 18410, 22716, 15109, 18562, 18465, 15110, 28773, 32289, 18340, 18526, 1574]
    #
    # H, I, J's:
    # goodIDs = [12763, 6043, 1648, 1650, 37087, 32299, 37102, 12944, 32329, 32332, 12958, 6229, 32365, 23694, 37185, 15862, 28805, 1752, 1757, 26354, 32444, 18674,
    #            13045, 18694, 18729, 24885, 18739, 32482, 1879, 41874, 26366, 37301, 18766, 37314, 1845, 13091, 13097, 26375, 6341, 13099, 32545, 32560, 37377, 
    #            32583, 32581, 32586, 37386, 15905, 13139, 37417, 27251, 13154, 26387, 37415, 2017, 2027, 37496, 2037, 32764, 2055, 32665, 32738, 32730, 37527, 37532,
    #            32712, 35208, 37573, 35203, 37572, 2128, 35216, 26398, 35223]
    
    # badIDs = [36687, 12807, 15081, 15082, 18563, 15084, 28762, 12948, 28776, 15086, 37148, 15088, 28370, 15091, 28798, 18623, 15094, 15095, 15096, 15097, 15099, 
    #           18693, 15102, 18732, 18734, 24956, 15103, 31493, 15104, 15105, 15106, 28863, 18769, 28883, 15107, 13096, 37318, 15108, 15109, 13104, 15112, 28904,
    #           1915, 15118, 32585, 15119, 28916, 18848, 15120, 28929, 27254, 13162, 15123, 37438, 15124, 15125, 32695, 15127, 28964, 15129, 32775, 28979, 28986, 
    #           6435, 28991, 15132, 15133, 35271, 35230, 29003, 15134, 28996, 15135, 29013]
    
    # K, L, M's:
    # goodIDs = [32718, 26353, 9519, 26359, 2205, 32742, 42753, 15911, 26378, 19035, 37646, 32818, 37672, 22451, 2272, 22667, 37684, 37703, 32851, 22133, 13366, 13367,
    #            19137, 22139, 37755, 37768, 29080, 41833, 37806, 41830, 42819, 29137, 19248, 15158, 15968, 37860, 11464, 23571, 32940, 25510, 37910, 37919, 26414, 4792,
    #            37948, 32988, 32992, 13491, 6705, 33005, 26423, 2540, 2541, 6720, 41333, 42870, 29240, 29241, 15150, 2573, 2567, 33042, 26441, 19473, 13571, 19503, 
    #            41367, 2635, 29334, 29337, 13581, 26523, 42957, 42958, 6784, 2647, 41372, 33139, 2659, 2664, 29409, 13657, 2693, 33187, 41379, 13660, 16056, 33127, 38274,
    #            29442, 29440, 13670, 33252, 2732, 43003, 2918, 33061, 19675, 33063, 2925, 26457, 19957, 26461, 26464, 26466, 2944, 41860, 42942, 26483, 3041, 29290, 29291,
    #            29300, 33093, 26547, 38345, 16074, 33273, 13714, 13716, 13724, 2775, 13735, 2790, 26563, 6892, 26564, 33315, 16099, 38437, 33119, 19826, 33161, 41895,
    #            29556, 33347, 38503, 29576, 2882, 38532, 29586, 19908, 19898, 13800, 26585, 33387, 33388, 29592, 38565, 26598, 11573]
    
    # badIDs = [15098, 15099, 15100, 15101, 15104, 15108, 4936, 19019, 15111, 2230, 32762, 15113, 6547, 22453, 15114, 15106, 32830, 32843, 32866, 22153, 13374, 13368, 
    #           2310, 22154, 32865, 29078, 29084, 41298, 29113, 15121, 42535, 15126, 19255, 15159, 15125, 29144, 32934, 15131, 32961, 25518, 9686, 29164, 15137, 13474, 
    #           29187, 37965, 15142, 4840, 15144, 33006, 15146, 15147, 15148, 33017, 41334, 42865, 29247, 29248, 33034, 15153, 15154, 29252, 15156, 19481, 13573, 19603,
    #           42962, 15181, 15182, 33233, 13617, 15184, 42983, 42985, 15187, 15189, 41373, 6812, 15191, 15190, 15197, 15199, 15200, 33202, 15201, 13662, 19647, 33250, 
    #           38321, 33130, 27348, 13676, 38329, 15202, 43007, 15160, 38048, 19963, 29263, 15162, 15165, 19960, 15167, 15168, 15170, 15169, 15171, 43011, 15172, 42909,
    #           27308, 27309, 29452, 33105, 15204, 9926, 3009, 29475, 13671, 13719, 13726, 15210, 26561, 15213, 15214, 15215, 15216, 15221, 19803, 33327, 15225, 21812,
    #           33337, 41896, 27377, 29557, 13795, 29579, 15232, 29301, 27387, 15235, 19903, 13802, 15236, 15237, 15238, 29605, 29606, 15242, 19936]
    
    # N, O, P, Q's:
    # goodIDs = [26506, 3050, 33293, 3068, 29524, 19975, 33321, 42963, 29538, 38558, 40710, 33359, 38597, 13869, 38616, 29580, 23703, 38627, 38633, 38638, 38645, 38648, 
    #            38651, 20099, 33402, 23871, 43045, 29628, 13905, 16086, 33428, 29656, 43055, 33433, 29673, 33444, 29677, 38771, 7279, 3251, 33540, 29711, 38819, 10248, 
    #            38850, 20296, 25584, 3310, 16123, 29767, 29766, 3366, 26584, 29780]
    
    # badIDs = [15162, 15165, 33297, 15170, 15171, 10095, 15173, 19985, 29545, 29553, 40715, 33363, 10139, 12166, 24971, 27322, 23989, 29589, 27324, 29591, 3151, 29595,
    #           29596, 20104, 33536, 15184, 43047, 15185, 13937, 20170, 15187, 38732, 43058, 33481, 27338, 33454, 27339, 29678, 15190, 15191, 29708, 15193, 33542, 15194,
    #           29724, 20312, 25585, 15198, 20349, 30320, 15200, 20367, 15202, 15203]
    
    # R, S, T's:
    goodIDs = [33533, 29749, 38886, 38890, 3385, 33557, 33567, 29767, 41749, 38926, 38936, 29790, 20452, 3444, 33596, 33613, 29807, 29818, 20483, 38990, 3470, 3502, 39020,
               14157, 41290, 39032, 14168, 7450, 20553, 39073, 33730, 20589, 14195, 16156, 3595, 39146, 39147, 14235, 20647, 16165, 39220, 16179, 16180, 20801, 3698, 20757,
               3707, 23262, 27047, 14292, 3723, 20768, 3730, 30010, 20773, 39300, 33786, 30026, 33790, 43247, 14348, 34093, 39354, 20780, 39352, 23990, 34085, 30082, 26639,
               41337, 16233, 16236, 3817, 16238, 3883, 33843, 39441, 20989, 24174, 33856, 26653, 39459, 30136, 14426, 3941, 39547, 3955, 21098, 43272, 33922, 41350, 39565,
               33810, 20781, 23298, 3993, 4039, 16283, 30198, 33978, 39644, 21205, 41801, 23490, 16297, 34116, 14531, 34120, 14532, 7787, 34137, 34140, 39718, 30238, 30245,
               34163, 34190, 39738, 4160, 23493, 34208, 34216, 7823, 34225, 27478, 34232, 7854, 4206, 4214, 30310, 23998]
    
    badIDs = [25557, 15185, 3380, 29751, 15187, 34558, 29766, 29774, 15190, 27345, 29787, 38940, 20479, 15193, 33604, 31415, 29811, 15197, 38978, 3473, 15198, 15200, 33703,
              14159, 15201, 29860, 14171, 15204, 20555, 33722, 29880, 20594, 12031, 20619, 15211, 10473, 29917, 26603, 20650, 20685, 29968, 20814, 20815, 20825, 15223, 20759,
              15215, 14285, 24576, 14296, 15218, 20853, 15219, 4835, 20858, 3867, 39304, 30049, 34056, 43322, 39341, 16226, 30064, 20929, 30065, 24478, 30074, 30083, 15231,
              43307, 20954, 20962, 15232, 20971, 15233, 33845, 33846, 39447, 24574, 15235, 15237, 4081, 33894, 15242, 15243, 30164, 15248, 21100, 43277, 33925, 27437, 24233,
              33933, 21137, 16275, 15251, 15252, 21169, 39626, 34314, 30208, 16293, 15257, 15259, 21223, 30224, 14538, 34128, 14539, 15264, 34139, 15265, 34142, 30197, 34147,
              33309, 34539, 11748, 15266, 15269, 15270, 30268, 15272, 7843, 23495, 39797, 15275, 16329, 15276, 30313, 24578]
    
    main(fullClean, goodIDs, badIDs)
    

# workbook_path = r"\\ucclerk\pgmdoc\Veterans\VeteransWeb.xlsx"
# wb = openpyxl.load_workbook(workbook_path)
# for sheet in wb.worksheets:
#     for row in range(2, sheet.max_row + 1):
#         cell = sheet[f'O{row}']
#         if cell.hyperlink:
#             cell.value = cell.hyperlink.target
# wb.save(workbook_path)

# workbook_path = r"\\ucclerk\pgmdoc\Veterans\VeteransWeb.xlsx"
# wb = openpyxl.load_workbook(workbook_path)

# old_prefix = r"VeteranCards"
# new_prefix = r"\\WEBDB914\VeteranCards"

# for sheet in wb.worksheets:
#     for row in range(2, sheet.max_row + 1):
#         cell = sheet[f'O{row}']
#         if cell.hyperlink:
#             old_link = cell.hyperlink.target
#             if old_link.startswith(old_prefix):
#                 new_link = old_link.replace(old_prefix, new_prefix, 1)
#                 cell.hyperlink.target = new_link
#                 cell.value = new_link

# wb.save(r"\\ucclerk\pgmdoc\Veterans\VeteransWeb.xlsx")
