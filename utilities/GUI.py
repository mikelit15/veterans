from PyQt6.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, \
     QHBoxLayout, QPushButton, QLineEdit, QLabel, QGroupBox, QFormLayout, QDialog, \
     QTextEdit, QScrollArea
from PyQt6.QtGui import QPixmap, QFont, QIcon
from PyQt6.QtCore import Qt, QThread, pyqtSignal
import os
import openpyxl
import re
from openpyxl.worksheet.hyperlink import Hyperlink
import shutil
import sys
import pandas as pd
import duplicates
sys.path.append(r'C:\workspace\veterans')
from microsoftOCR import microsoftOCR

class Worker(QThread):
    adjustImageName_signal = pyqtSignal(int, int, int)  # Signal to emit IDs
    cleanDelete_signal = pyqtSignal(str, int, int) # Signal to emit KVS
    cleanImages_signal = pyqtSignal(int, str, str) # Signal to emit Error Message
    cleanHyperlinks_signal = pyqtSignal(str, int) # Signal to emit Image Path
    
    def __init__(self, cemetery, goodIDs, badIDs):
        super().__init__()
        self.cemetery = cemetery
        self.goodIDs = goodIDs
        self.badIDs = badIDs
        self.paused = False

    def run(self):
        global excelFilePath
        excelFilePath = r"\\ucclerk\pgmdoc\Veterans\Veterans.xlsx"
        global workbook
        workbook = openpyxl.load_workbook(excelFilePath)
        global worksheet 
        worksheet = workbook[self.cemetery]
        for x in range(0, len(self.goodIDs)):
            workbook = openpyxl.load_workbook(excelFilePath)
            worksheet = workbook[self.cemetery]
            for row in range(1, worksheet.max_row + 1):
                if worksheet[f'A{row}'].value == self.goodIDs[x]:
                    goodRow = row
                if worksheet[f'A{row}'].value == self.badIDs[x]:
                    badRow = row
            letter = worksheet[f'B{goodRow}'].value[0]
            self.adjustImageName_signal.emit(self.goodIDs[x], self.badIDs[x], goodRow)
            workbook.save(excelFilePath)
            microsoftOCR.main(True, self.cemetery, letter)
            workbook = openpyxl.load_workbook(excelFilePath)
            worksheet = workbook[self.cemetery]
            self.cleanDelete_signal.emit(self.cemetery, self.badIDs[x], badRow)
            workbook.save(excelFilePath)
        mini = min(self.goodIDs)
        if min(self.badIDs) < mini:
            mini = min(self.badIDs) - 1
        self.cleanImages_signal.emit(mini, "", "")
        self.cleanImages_signal.emit(mini, " - Redacted", " redacted")
        workbook = openpyxl.load_workbook(excelFilePath)
        worksheet = workbook[self.cemetery]
        for row in range(1, worksheet.max_row + 1):
            if worksheet[f'A{row}'].value == mini:
                startIndex = row
        self.cleanHyperlinks_signal.emit(self.cemetery, startIndex)
        workbook.save(excelFilePath)
        print("\n")
        duplicates.main(self.cemetery)
                
class MainWindow(QMainWindow):
    updateLabelSignal = pyqtSignal(str)
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Veteran Duplicate Cleaner")
        self.setGeometry(400, 150, 1050, 150)
        self.worker = None  
        self.mainLayout()
        self.updateLabelSignal.connect(self.updateScroll) 

    def mainLayout(self):
        centralWidget = QWidget(self)
        centralWidget.setObjectName("mainCentralWidget")
        self.setCentralWidget(centralWidget)
        self.setStyleSheet("#mainCentralWidget { background-color: white; }")
        layout1 = QVBoxLayout(centralWidget)
        
        # Create the top container
        topContainer = QWidget()
        topLayout = QHBoxLayout()
        logoLabel = QLabel() 
        parent_path = os.path.dirname(os.getcwd())
        logoImage = QPixmap(f"{parent_path}/veteranData/ucLogo.png") 
        logoLabel.setPixmap(logoImage)
        programName = QLabel("Union County Clerk's Office\n\n Veteran Duplicate Cleaner")
        font = QFont()
        font.setBold(True)
        font.setPointSize(26)
        programName.setFont(font)
        topLayout.addWidget(logoLabel)
        topLayout.addWidget(programName)
        topContainer.setLayout(topLayout)

        # Create a group box for the middle top left
        middleTopLeftGroupBox = QGroupBox("Parameters")
        middleTopLeftGroupBox.setFixedSize(375, 375)
        middleTopLeftLayout = QFormLayout()
        self.cemeteryBox = QLineEdit()
        self.cemeteryBox.setFixedWidth(140)
        self.goodIDBox = QTextEdit()
        self.goodIDBox.setFixedWidth(250)
        self.goodIDBox.setFixedHeight(75)
        self.goodIDBox.setLineWrapColumnOrWidth(15)
        self.badIDBox = QTextEdit()
        self.badIDBox.setFixedWidth(250)
        self.badIDBox.setFixedHeight(75)
        self.badIDBox.setLineWrapColumnOrWidth(15)
        middleTopLeftLayout.addRow(" ", None)
        middleTopLeftLayout.addRow("Cemetery :   ", self.cemeteryBox)
        middleTopLeftLayout.addRow(None, QLabel("Note: This is the cemetery name based on its folder."))
        middleTopLeftLayout.addRow(" ", None)
        middleTopLeftLayout.addRow("Good IDs :", self.goodIDBox)
        middleTopLeftLayout.addRow(None, QLabel("Note: This is a list of all the original IDs."))
        middleTopLeftLayout.addRow(" ", None)
        middleTopLeftLayout.addRow("Bad IDs :", self.badIDBox)
        middleTopLeftLayout.addRow(None, QLabel("Note: This is a list of all the duplicate IDs."))
        middleTopLeftGroupBox.setLayout(middleTopLeftLayout)
        self.goodIDBox.setDisabled(True)
        self.badIDBox.setDisabled(True)
        
        # Create a group box for the middle top right
        self.middleTopRightGroupBox = QGroupBox("Results")
        self.middleTopRightGroupBox.setFixedSize(775, 375)
        self.middleTopRightLayout1 = QFormLayout()
        self.middleTopRightLayout2 = QFormLayout()
        
        self.duplicatesLabel = QLabel("")
        self.scroll_area1 = QScrollArea()
        self.scroll_area1.setWidgetResizable(True) 
        self.scroll_area1.setWidget(self.duplicatesLabel)
        self.middleTopRightLayout1.addRow(" ", None)
        self.middleTopRightLayout1.addRow("Duplicate IDs :", self.scroll_area1)
        self.middleTopRightLayout1.addRow(" ", None)
        
        self.duplicatesLabel2 = QLabel("")
        self.scroll_area2 = QScrollArea()
        self.scroll_area2.setWidgetResizable(True) 
        self.scroll_area2.setWidget(self.duplicatesLabel2)
        self.middleTopRightLayout2.addRow(" ", None)
        self.middleTopRightLayout2.addRow("Details :", self.scroll_area2)
        self.middleTopRightLayout2.addRow(" ", None)
        
        self.middleTopRightGroupBox.setLayout(self.middleTopRightLayout1)

        # Create container for middle top left and top right group boxes
        middleTopContainer = QWidget()
        middleTopLayout = QHBoxLayout()
        middleTopLayout.addWidget(middleTopLeftGroupBox)
        middleTopLayout.addWidget(self.middleTopRightGroupBox)
        middleTopContainer.setLayout(middleTopLayout)
        
        # Create the bottom container with a button
        bottomContainer = QWidget()
        bottomLayout = QHBoxLayout()
        self.checkButton = QPushButton("Check Duplicates")
        self.checkButton.clicked.connect(self.updateDuplicates)
        self.checkButton.setFixedWidth(150)
        self.pauseButton = QPushButton("Pause Code")
        self.pauseButton.clicked.connect(self.togglePauseResume)
        self.pauseButton.setFixedWidth(150)
        self.cleanButton = QPushButton("Clean Duplicates")
        self.cleanButton.clicked.connect(self.startProcessing)
        self.cleanButton.setFixedWidth(150)
        self.pauseButton.setDisabled(True)
        self.cleanButton.setDisabled(True)
        bottomLayout.addWidget(self.checkButton)
        bottomLayout.addWidget(self.pauseButton)
        bottomLayout.addWidget(self.cleanButton)
        bottomContainer.setLayout(bottomLayout)
        
        # Create a main layout to arrange the containers
        mainLayout = QVBoxLayout()
        mainLayout.addWidget(topContainer)
        mainLayout.addWidget(middleTopContainer)
        mainLayout.addWidget(bottomContainer)
        self.status = QLabel("              Status :  Idle\n")
        self.status.setFixedWidth(150)
        mainLayout.addWidget(self.status, 0, alignment=Qt.AlignmentFlag.AlignCenter)
        layout1.addLayout(mainLayout)
        
    def popupWindow(self, text):
        popup = QDialog()
        popup.setWindowTitle(" ")
        parent_path = os.path.dirname(os.getcwd())
        popup.setWindowIcon(QIcon(f"{parent_path}/veteranData/veteranLogo.png"))
        popup.setGeometry(850, 500, 200, 100)
        layout = QVBoxLayout()
        message_label = QLabel(text)
        message_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(message_label)
        close_button = QPushButton("Close")
        close_button.clicked.connect(popup.close)
        layout.addWidget(close_button)
        popup.setLayout(layout)
        popup.exec()
        self.checkButton.setDisabled(False)
        self.status.setText("              Status :  Idle\n")
    
    def togglePauseResume(self):
        if self.worker:
            if self.worker.paused:
                self.worker.paused = False
                self.pauseButton.setText("Pause")
                self.status.setText("              Status : Running...\n")
            else:
                self.worker.paused = True
                self.pauseButton.setText("Resume")
                self.status.setText("              Status : Paused\n" )
    
    def clearLayout(self, layout):
        while layout.count():
            item = layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
            elif item.layout():
                self.clearLayout(item.layout())
                
    def switchToLayout2(self):
        if self.middleTopRightGroupBox.layout() is not None:
            self.clearLayout(self.middleTopRightGroupBox.layout())
            QWidget().setLayout(self.middleTopRightGroupBox.layout()) 
        self.middleTopRightGroupBox.setLayout(self.middleTopRightLayout2)
        
    def updateScroll(self, new_text):
        current_text = self.duplicatesLabel2.text()
        self.duplicatesLabel2.setText(current_text + new_text + "\n")
        self.scroll_area2.verticalScrollBar().setValue(self.scroll_area2.verticalScrollBar().maximum())
    
    def updateDuplicates(self):
        column_widths = {
        "VID": 8,
        "VLNAME": 15,
        "VFNAME": 15,
        "VMNAME": 5,
        "VSUF": 5,
        "VDOB": 15,
        "VDOBY": 10,
        "VDOD": 15,
        "VDODY": 10
        }
        try:
            df = duplicates.main(self.cemeteryBox.text())
        except Exception:
            self.popupWindow("Cemetery field not filled \n out or mispelled.")
            return
        # Apply fixed width formatting to each column
        for column, width in column_widths.items():
            df[column] = df[column].apply(lambda x: f"{x: <{width}}" if pd.notna(x) else " " * width)

        string = df.to_string(index=False, header=True, formatters={key: f'{{:>{width}}}'.format for key, width in column_widths.items()})
        self.duplicatesLabel.setText(string)
        self.scroll_area1.verticalScrollBar().setValue(self.scroll_area1.verticalScrollBar().maximum())
        self.cleanButton.setDisabled(False)
        self.goodIDBox.setDisabled(False)
        self.badIDBox.setDisabled(False)

          
    def startProcessing(self):
        try:
            self.goodIDBox = [int(numeric_string) for numeric_string in self.goodIDBox.toPlainText().split(", ")]
        except Exception:
            self.popupWindow("Good ID field is empty.")
            return
        try:
            self.badIDBox = [int(numeric_string) for numeric_string in self.badIDBox.toPlainText().split(", ")]
        except Exception:
            self.popupWindow("Bad ID field is empty.")
            return
        self.worker = Worker(self.cemeteryBox.text(), self.goodIDBox, self.badIDBox)
        self.worker.adjustImageName_signal.connect(lambda goodID, badID, goodRow: self.adjustImageName(goodID, badID, goodRow))
        self.worker.cleanDelete_signal.connect(lambda cemetery, badID, badRow: self.cleanDelete(cemetery, badID, badRow))
        self.worker.cleanImages_signal.connect(lambda goodID, redac, redac2: self.cleanImages(goodID, redac, redac2))
        self.worker.cleanHyperlinks_signal.connect(lambda cemetery, startIndex: self.cleanHyperlinks(cemetery, startIndex))
        self.worker.start()
        self.status.setText("              Status:  Running...\n")
        self.cleanButton.setDisabled(True)
        self.switchToLayout2()
        
    def cleanImages(self, goodID, redac, redac2):
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
                        self.updateLabelSignal.emit(f"Processing {cemetery} - {subCemetery}")
                        print(f"Processing {cemetery} - {subCemetery}")
                        initialCount, startFlag = self.process_cemetery(subCemetery, subCemPath, initialCount, start, startFlag, redac2)
                else:
                    self.updateLabelSignal.emit(f"Processing {cemetery}")
                    print(f"Processing {cemetery}")
                    initialCount, startFlag = self.process_cemetery(cemetery, baseCemPath, initialCount, start, startFlag, redac2)
            
    def process_cemetery(self, cemeteryName, basePath, initialCount, start, startFlag, redac2):
        uppercase_alphabet = [chr(i) for i in range(ord('A'), ord('Z') + 1)]
        isFirstFile = True
        for letter in uppercase_alphabet:
            name_path = os.path.join(basePath, cemeteryName, letter)
            try:
                self.updateLabelSignal.emit(f"Processing {cemeteryName} - {letter}\n")
                print(f"Processing {cemeteryName} - {letter}")
                initialCount, startFlag = self.clean(letter, name_path, os.path.join(basePath, cemeteryName), \
                    initialCount, isFirstFile, start, startFlag, redac2)
                isFirstFile = False
            except FileNotFoundError:
                continue
        return initialCount, startFlag

    def clean(self, letter, namePath, cemPath, initialCount, isFirstFile, start, startFlag, redac2):
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
                    self.updateLabelSignal.emit(f"Removed {x}")
                    print(f"Removed {x}")
                elif "b redacted.pdf" in x:
                    counter -= 1
                    os.remove(os.path.join(namePath, x))
                    self.updateLabelSignal.emit(f"Removed {x}")
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

    def adjustImageName(self, goodID, badID, goodRow):
        global excelFilePath
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
                self.updateLabelSignal.emit(f"{os.path.basename(goodIDFilePath)} renamed to {newGoodIDFilename}")
                print(f"{os.path.basename(goodIDFilePath)} renamed to {newGoodIDFilename}")
            shutil.move(badIDFilePath, newBadIDFilePath)
            self.updateLabelSignal.emit(f"{os.path.basename(badIDFilePath)} moved and renamed to {newBadIDFilename}")
            print(f"{os.path.basename(badIDFilePath)} moved and renamed to {newBadIDFilename}")
        else:
            self.updateLabelSignal.emit("GoodID or BadID file not found. Please check the IDs and directories.")
            print("GoodID or BadID file not found. Please check the IDs and directories.")
            workbook.save(excelFilePath)
            quit()
        start_column = 'A'
        end_column = 'O'
        start_column_index = openpyxl.utils.column_index_from_string(start_column)
        end_column_index = openpyxl.utils.column_index_from_string(end_column)
        for col_index in range(start_column_index, end_column_index + 1):
            worksheet.cell(row=goodRow, column=col_index).value = None
        worksheet[f'O{goodRow}'].hyperlink = None
        self.updateLabelSignal.emit(f"Record {goodID} data from row {goodRow} cleared successfully.")
        print(f"Record {goodID} data from row {goodRow} cleared successfully.")

    def cleanDelete(self, cemetery, badID, badRow):
        baseRedacPath = r"\\ucclerk\pgmdoc\Veterans\Cemetery - Redacted"
        text = ""
        for dirpath, dirnames, filenames in os.walk(baseRedacPath):
            for filename in filenames:
                if f"{badID:05d}" in filename:
                    file_path = os.path.join(dirpath, filename)
                    os.remove(file_path)
                    text = f"\nRedacted file, {filename}, deleted successfully."
                    print(f"\nRedacted file, {filename}, deleted successfully.")
        self.updateLabelSignal.emit(text)
        worksheet = workbook[cemetery]
        worksheet.delete_rows(badRow)
        self.updateLabelSignal.emit(f"Row {badRow} deleted successfully.")
        print(f"Row {badRow} deleted successfully.")
        text = ""
        for row in range(badRow, worksheet.max_row + 1):
            next_row = row + 1
            cell_ref = f'O{row}'
            next_cell_ref = f'O{next_row}'
            if next_row <= worksheet.max_row and worksheet[next_cell_ref].hyperlink:
                temp_hyperlink = worksheet[next_cell_ref].hyperlink
                worksheet[cell_ref].hyperlink = Hyperlink(ref=cell_ref, target=temp_hyperlink.target, display=temp_hyperlink.display)
                worksheet[cell_ref].value = worksheet[next_cell_ref].value
                text = text + f"Hyperlink and value from row {next_row} moved to row {row}.\n"
                print(f"Hyperlink and value from row {next_row} moved to row {row}.")
        self.updateLabelSignal.emit(text)
        if worksheet[f'O{worksheet.max_row - 1}'].hyperlink:
            last_row = worksheet.max_row
            self.updateLabelSignal.emit(f"Adjusting the hyperlink in the last row, row {last_row}.")
            print(f"Adjusting the hyperlink in the last row, row {last_row}.")
            second_last_hyperlink = worksheet[f'O{last_row - 1}'].hyperlink
            new_target = re.sub(r'(\d+)(?=%\d+)', lambda m: f"{int(m.group(1)) + 1:05d}", second_last_hyperlink.target)
            worksheet[f'O{last_row}'].hyperlink = Hyperlink(ref=f'O{last_row}', target=new_target, display=second_last_hyperlink.display)
            self.updateLabelSignal.emit(f"Updated hyperlink in row {last_row} to new target, {new_target}.\n")
            print(f"Updated hyperlink in row {last_row} to new target, {new_target}.\n")

    def cleanHyperlinks(self, cemetery, startIndex):
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
                self.updateLabelSignal.emit(f"Updated hyperlink from {orig_target} to {modified_string} in row {row}.")
                print(f"Updated hyperlink from {orig_target} to {modified_string} in row {row}.")
            new_id += 1

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    parent_path = os.path.dirname(os.getcwd())
    window.setWindowIcon(QIcon(f"{parent_path}/veteranData/veteranLogo.png"))
    window.show()
    sys.exit(app.exec())