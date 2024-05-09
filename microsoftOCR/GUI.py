from PyQt6.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, \
     QHBoxLayout, QPushButton, QLineEdit, QLabel, QGroupBox, QFormLayout, QDialog
from PyQt6.QtGui import QPixmap, QFont, QIcon
from PyQt6.QtCore import Qt, QThread, pyqtSignal
import os
import openpyxl
from openpyxl.styles import PatternFill, Font
import traceback
import sys
import microsoftOCR


class Worker(QThread):
    id_signal = pyqtSignal(str)  # Signal to emit IDs
    kvs_signal = pyqtSignal(str) # Signal to emit KVS
    popup_signal = pyqtSignal(str) # Signal to emit popup
    def __init__(self, singleFlag, singleCem, singleLetter, initialID):
        super().__init__()
        self.singleFlag = singleFlag
        self.singleCem = singleCem
        self.singleLetter = singleLetter
        self.initialID = initialID
        self.paused = False
        self.stopped = False

    def run(self):
        global cemSet
        global miscSet
        global jewishSet
        networkFolder = r"\\ucclerk\pgmdoc\Veterans"
        os.chdir(networkFolder)
        cemeterys = []
        for x in os.listdir(r"Cemetery"):
            cemeterys.append(x)
        miscs = []
        for x in os.listdir(r"Cemetery\Misc"):
            miscs.append(x)
        jewishs = []
        for x in os.listdir(r"Cemetery\Jewish"):
            jewishs.append(x)
        cemSet = set(cemeterys)
        miscSet = set(miscs)
        jewishSet = set(jewishs)
        workbook = openpyxl.load_workbook('Veterans.xlsx')
        cemetery = self.singleCem 
        cemPath = os.path.join(networkFolder, fr"Cemetery\{cemetery}")
        letter = self.singleLetter 
        namePath = letter 
        namePath = os.path.join(cemPath, namePath)
        pathA = ""
        rowIndex = 2
        initialID = self.initialID
        global worksheet
        worksheet = workbook[cemetery]
        warFlag = False
        pdfFiles = sorted(os.listdir(namePath))
        breakFlag = False
        
        for y in range(len(pdfFiles)):
            try:
                warFlag = False
                filePath = os.path.join(namePath, pdfFiles[y])
                rowIndex = microsoftOCR.find_next_empty_row(worksheet)
                try:
                    id = worksheet[f'{"A"}{rowIndex-1}'].value + 1
                except TypeError:
                    id = initialID
                while self.paused:
                    self.sleep(1)
                if "output" in pdfFiles[y] or "redacted" in pdfFiles[y]:
                    continue
                else:
                    string = pdfFiles[y][:-4]
                    string = string.split(letter) 
                    string = string[-1].lstrip('0')
                    if "a" not in string and "b" not in string:
                        if id != int(string.replace("a", "").replace("b", "")):
                            if self.stopped:
                                break
                            continue
                        if self.stopped:
                                break
                        self.id_signal.emit(str(id))
                        self.kvs_signal.emit("None")
                        vals, warFlag, nameCoords, serialCoords, printedKVS, kinLast = microsoftOCR.createRecord(filePath, id, cemetery)
                        redactedFile = microsoftOCR.redact(filePath, cemetery, letter, nameCoords, serialCoords)
                        worksheet.cell(row=rowIndex, column=15).value = "PDF Image"
                        worksheet.cell(row=rowIndex, column=15).font = Font(underline="single", color="0563C1")
                        worksheet.cell(row=rowIndex, column=15).hyperlink = redactedFile
                        counter = 1
                        worksheet.cell(row=rowIndex, column=counter, value=id)
                        counter += 1
                        for x in vals:
                            worksheet.cell(row=rowIndex, column=counter, value=x)
                            counter += 1
                        microsoftOCR.highlightSingle(worksheet, cemetery, letter, warFlag, rowIndex, kinLast)
                        id += 1
                        rowIndex += 1
                    else:
                        if id != int(string.replace("a", "").replace("b", "")):
                            continue
                        else:
                            if "a" in string:
                                if (filePath.replace("a.pdf", "") in pdfFiles):
                                    continue
                                pathA = filePath
                                self.id_signal.emit(f"{id} A")
                                self.kvs_signal.emit("None")
                                vals1, warFlag, nameCoords, serialCoords, printedKVS, kinLast = microsoftOCR.tempRecord(filePath, id, cemetery, "A")
                                redactedFile = microsoftOCR.redact(filePath, cemetery, letter, nameCoords, serialCoords)
                            if "b" in string:
                                if (filePath.replace("b.pdf", "") in pdfFiles):
                                    continue
                                self.id_signal.emit(f"{id} B")
                                self.kvs_signal.emit("None")
                                vals2, warFlagB, nameCoords, serialCoords, printedKVS, kinLast = microsoftOCR.tempRecord(filePath, id, cemetery, "B")
                                if not warFlag or not warFlagB:
                                    warFlag = False
                                else:
                                    warFlag = True
                                redactedFile = microsoftOCR.redact(filePath, cemetery, letter, nameCoords, serialCoords)
                                microsoftOCR.mergeRecords(worksheet, vals1, vals2, rowIndex, id, warFlag, cemetery, letter)
                                microsoftOCR.mergeImages(pathA, filePath, cemetery, letter)
                                link_text = "PDF Image"
                                worksheet.cell(row=rowIndex, column=15).value = link_text
                                worksheet.cell(row=rowIndex, column=15).font = Font(underline="single", color="0563C1")
                                worksheet.cell(row=rowIndex, column=15).hyperlink = redactedFile.replace("b redacted.pdf", " redacted.pdf")
                                id += 1
                                rowIndex += 1
                                if self.singleFlag:
                                    breakFlag = True
                self.kvs_signal.emit(printedKVS)
            except Exception as e:
                errorTraceback = traceback.format_exc()
                print(errorTraceback)
                print(f"An error occurred: {e}")
                errorMessage = f"SKIPPED DUE TO ERROR : {errorTraceback}"
                worksheet.cell(row=rowIndex, column=1, value=id)
                worksheet.cell(row=rowIndex, column=2, value=errorMessage)
                highlightColor = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
                for colIndex in range(1, worksheet.max_column + 1):
                    cell = worksheet.cell(row=rowIndex, column=colIndex)
                    cell.fill = highlightColor
                error_file_path = fr'Errors/{cemetery}{letter}{str(id).zfill(5)} Error.txt' 
                with open(error_file_path, 'a') as error_file:
                    error_file.write(f'{printedKVS} \n\n {errorTraceback}')
                id += 1
                rowIndex += 1
            extension1 = str(id-1).zfill(5)
            extension2 = ""
            if "a.pdf" in filePath:
                extension1 = str(id).zfill(5)
                extension2 = "a"
            elif "b.pdf" in filePath:
                extension1 = str(id-1).zfill(5)
                extension2 = "b"
            logFilePath = fr'Logs/{cemetery}{letter}{extension1}{extension2} Extracted.txt' 
            with open(logFilePath, 'w', encoding='utf-8') as logFile:
                logFile.write(printedKVS)
            workbook.save('Veterans.xlsx')
            if breakFlag:
                break
        self.popup_signal.emit(f"{cemetery} letter {letter} has finished processing.\n Please enter the next letter.")
        return
   
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Veteran Extraction")
        self.resize(400, 200)
        self.setGeometry(600, 200, 250, 150)
        self.worker = None  
        self.mainLayout()

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
        programName = QLabel("Union County Clerk's Office\n\n       Veteran Extraction")
        font = QFont()
        font.setBold(True)
        font.setPointSize(26)
        programName.setFont(font)
        topLayout.addWidget(logoLabel)
        topLayout.addWidget(programName)
        topContainer.setLayout(topLayout)

        # Create a group box for the middle left
        middleLeftGroupBox = QGroupBox("Parameters")
        middleLeftGroupBox.setFixedSize(375, 350)
        middleLeftLayout = QFormLayout()
        self.cemeteryBox = QLineEdit()
        self.cemeteryBox.setFixedWidth(140)
        self.letterBox = QLineEdit()
        self.letterBox.setFixedWidth(30)
        self.initialIDBox = QLineEdit()
        self.initialIDBox.setFixedWidth(45)
        middleLeftLayout.addRow(" ", None)
        middleLeftLayout.addRow(" ", None)
        middleLeftLayout.addRow("Cemetery :   ", self.cemeteryBox)
        middleLeftLayout.addRow(" ", None)
        middleLeftLayout.addRow(" ", None)
        middleLeftLayout.addRow("Letter :", self.letterBox)
        middleLeftLayout.addRow(None, QLabel("Note: This is the last name letter to start processing."))
        middleLeftLayout.addRow(" ", None)
        middleLeftLayout.addRow("Initial ID :", self.initialIDBox)
        middleLeftLayout.addRow(None, QLabel("Note: This is the first ID for the cemetery."))
        middleLeftLayout.addRow(" ", None)
        middleLeftLayout.addRow(" ", None)
        middleLeftGroupBox.setLayout(middleLeftLayout)
        
        # Create a group box for the middle right
        middleRightGroupBox = QGroupBox("Results")
        middleRightGroupBox.setFixedSize(375, 350)
        middleRightLayout = QFormLayout()
        global displayLabel
        self.displayLabel = QLabel("None")
        global kvsLabel
        self.kvsLabel = QLabel("None")
        middleRightLayout.addRow(" ", None)
        middleRightLayout.addRow(" ", None)
        middleRightLayout.addRow("Current ID :", self.displayLabel)
        middleRightLayout.addRow(" ", None)
        middleRightLayout.addRow(" ", None)
        middleRightLayout.addRow("Key-Value Pairs :", self.kvsLabel)
        middleRightLayout.addRow(" ", None)
        middleRightLayout.addRow(" ", None)
        middleRightGroupBox.setLayout(middleRightLayout)
        
        # Create the bottom container with a button
        bottomContainer = QWidget()
        bottomLayout = QHBoxLayout()
        self.runButton = QPushButton("Run Code")
        self.runButton.clicked.connect(self.startProcessing)
        self.runButton.setFixedWidth(150)
        self.pauseButton = QPushButton("Pause Code")
        self.pauseButton.clicked.connect(self.togglePauseResume)
        self.pauseButton.setFixedWidth(150)
        self.status = QLabel("Status: Idle")
        self.status.setFixedWidth(100)
        bottomLayout.addWidget(self.runButton)
        bottomLayout.addWidget(self.status)
        bottomLayout.addWidget(self.pauseButton)
        bottomContainer.setLayout(bottomLayout)

        middleContainer = QWidget()
        middleLayout = QHBoxLayout()
        middleLeftGroupBox.setFixedSize(375, 350)
        middleLayout.addWidget(middleLeftGroupBox)
        middleLayout.addWidget(middleRightGroupBox)
        middleContainer.setLayout(middleLayout)
        
        # Create a main layout to arrange the containers
        mainLayout = QVBoxLayout()
        mainLayout.addWidget(topContainer)
        mainLayout.addWidget(middleContainer)
        mainLayout.addWidget(bottomContainer)
        self.stopButton = QPushButton("Stop Code")
        self.stopButton.setFixedWidth(225)
        self.stopButton.clicked.connect(self.stopProcessing)
        mainLayout.addWidget(self.stopButton, 0, alignment=Qt.AlignmentFlag.AlignCenter)
        layout1.addLayout(mainLayout)
        
    def popupWindow(self, text):
        popup = QDialog()
        popup.setWindowTitle(" ")
        parent_path = os.path.dirname(os.getcwd())
        popup.setWindowIcon(QIcon(f"{parent_path}/veteranData/vetIcon.png"))
        popup.setGeometry(850, 500, 250, 150)
        layout = QVBoxLayout()
        message_label = QLabel(text)
        message_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(message_label)
        close_button = QPushButton("Close")
        close_button.clicked.connect(popup.close)
        layout.addWidget(close_button)
        popup.setLayout(layout)
        popup.exec()
        self.runButton.setDisabled(False)
        self.status.setText("Status: Idle")
    
    def updateID(self, id):
        self.displayLabel.setText(id)
    
    def updateKVS(self, printedKVS):
        self.kvsLabel.setText(f"{printedKVS}")  
          
    def startProcessing(self):
        self.worker = Worker(False, self.cemeteryBox.text(), self.letterBox.text(), self.initialIDBox.text())
        self.worker.id_signal.connect(lambda id: self.updateID(id))
        self.worker.kvs_signal.connect(lambda printedKVS: self.updateKVS(printedKVS))
        self.worker.popup_signal.connect(lambda text: self.popupWindow(text))
        self.worker.start()
        self.status.setText("Status: Running...")
        self.runButton.setDisabled(True)
    
    def togglePauseResume(self):
        if self.worker:
            if self.worker.paused:
                self.worker.paused = False
                self.pauseButton.setText("Pause")
                self.status.setText("Status: Running...")
            else:
                self.worker.paused = True
                self.pauseButton.setText("Resume")
                self.status.setText("Status: Paused" )
    
    def stopProcessing(self):
        self.worker.stopped = True
        self.status.setText("Status: Stopping...")
        self.runButton.setDisabled(False)
        
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.setWindowTitle("Veteran Extraction")
    parent_path = os.path.dirname(os.getcwd())
    window.setWindowIcon(QIcon(f"{parent_path}/veteranData/veteranLogo.png"))
    window.show()
    sys.exit(app.exec())