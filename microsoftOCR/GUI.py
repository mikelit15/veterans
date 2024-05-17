from PyQt6.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, \
     QHBoxLayout, QPushButton, QLineEdit, QLabel, QGroupBox, QFormLayout, \
     QComboBox, QMessageBox
from PyQt6.QtGui import QPixmap, QFont, QIcon
from PyQt6.QtCore import Qt, QThread, pyqtSignal
import os
import openpyxl
from openpyxl.styles import PatternFill, Font
import traceback
import sys
import microsoftOCR
import qdarktheme

'''
Dark Mode Styling
'''
dark = qdarktheme.load_stylesheet(
            theme="dark",
            custom_colors=
            {
                "[dark]": 
                {
                    "primary": "#0078D4",
                    "border": "#8A8A8A",
                    "primary>button.hoverBackground": "#2B456E",
                    "background>list": "#3F4042",
                    "background>popup": "#303136",
                }
            },
        )

'''
Light Mode Styling
'''
light = qdarktheme.load_stylesheet(
            theme="light",
            custom_colors=
            {
                "[light]": 
                {
                    "foreground": "#111111",
                    "border": "#111111",
                    "background": "#f0f0f0",
                    "primary>button.hoverBackground": "#adcaf7",
                    "background>list": "#FFFFFF",
                    "background>popup": "#cfcfd1",
                }
            },
        )

class Worker(QThread):
    id_signal = pyqtSignal(str)  # Signal to emit IDs
    kvs_signal = pyqtSignal(str) # Signal to emit KVS
    error_signal = pyqtSignal(str) # Signal to emit Error Message
    image_signal = pyqtSignal(str) # Signal to emit Image Path
    popup_signal = pyqtSignal(str, str) # Signal to update app status 
    
    def __init__(self, singleFlag, singleCem, singleLetter):
        super().__init__()
        self.singleFlag = singleFlag
        self.singleCem = singleCem
        self.singleLetter = singleLetter.upper()
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
        global worksheet
        try:
            worksheet = workbook[cemetery]
        except Exception:
            self.popup_signal.emit('Critical', f"Cemetery '{cemetery}' not found in Veterans.xslx.")
            return
        warFlag = False
        try:
            pdfFiles = sorted(os.listdir(namePath))
        except Exception:
            self.popup_signal.emit('Critical', f"Letter '{letter}' not found in '{cemetery}' folder.")
            return
        breakFlag = False
        for y in range(len(pdfFiles)):
            try:
                warFlag = False
                filePath = os.path.join(namePath, pdfFiles[y])
                rowIndex = microsoftOCR.find_next_empty_row(worksheet)
                try:
                    id = worksheet[f'{"A"}{rowIndex-1}'].value + 1
                except Exception:
                    id = int(pdfFiles[0][:-4].split(letter)[-1].lstrip('0'))
                while self.paused:
                    self.sleep(1)
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
                    self.kvs_signal.emit("")
                    self.error_signal.emit("")
                    self.image_signal.emit("")
                    vals, warFlag, nameCoords, serialCoords, printedKVS, kinLast = microsoftOCR.createRecord(filePath, id, cemetery, "")
                    self.kvs_signal.emit(printedKVS)
                    redactedFile = microsoftOCR.redact(filePath, cemetery, letter, nameCoords, serialCoords)
                    self.image_signal.emit(r"\\ucclerk\pgmdoc\Veterans\temp.png")
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
                        if self.stopped:
                            break
                        continue
                    if self.stopped:
                        break
                    if "a" in string:
                        if (filePath.replace("a.pdf", "") in pdfFiles):
                            continue
                        pathA = filePath
                        self.id_signal.emit(f"{id} A")
                        self.kvs_signal.emit("")
                        self.error_signal.emit("")
                        self.image_signal.emit("")
                        vals1, warFlag, nameCoords, serialCoords, printedKVS, kinLast = microsoftOCR.createRecord(filePath, id, cemetery, "A")
                        self.kvs_signal.emit(printedKVS)
                        redactedFile = microsoftOCR.redact(filePath, cemetery, letter, nameCoords, serialCoords)
                        self.image_signal.emit(r"\\ucclerk\pgmdoc\Veterans\temp.png")
                    if "b" in string:
                        if (filePath.replace("b.pdf", "") in pdfFiles):
                            continue
                        self.id_signal.emit(f"{id} B")
                        self.kvs_signal.emit("")
                        self.error_signal.emit("")
                        self.image_signal.emit("")
                        vals2, warFlagB, nameCoords, serialCoords, printedKVS, kinLast = microsoftOCR.createRecord(filePath, id, cemetery, "B")
                        self.kvs_signal.emit(printedKVS)
                        if not warFlag or not warFlagB:
                            warFlag = False
                        else:
                            warFlag = True
                        redactedFile = microsoftOCR.redact(filePath, cemetery, letter, nameCoords, serialCoords)
                        self.image_signal.emit(r"\\ucclerk\pgmdoc\Veterans\temp.png")
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
                    self.error_signal.emit(errorTraceback)
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
        self.popup_signal.emit('Info', f"{cemetery} letter {letter} has finished processing.\n Please enter the next letter.")
        return
   
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Veteran Extraction")
        self.setGeometry(500, 35, 950, 150)
        self.worker = None  
        if self.loadDisplayMode() == "Light":
            app.setStyleSheet(light)
        else:
            app.setStyleSheet(dark)
        self.mainLayout()

    '''
    Saves display mode to display_mode.txt

    @param mode - the name of the display mode that is being used, and therefore saved

    @author Mike
    '''
    def saveDisplayMode(self, mode):
        parent_path = os.path.dirname(os.getcwd())
        with open(f"{parent_path}/veteranData/display_mode.txt", "w") as file:
            file.write(mode)


    '''
    Loads display mode name from display_mode.txt

    @return string - the name of the display mode that was saved

    @author Mike
    '''
    def loadDisplayMode(self):
        parent_path = os.path.dirname(os.getcwd())
        try:
            with open(f"{parent_path}/veteranData/display_mode.txt", "r") as file:
                return file.read().strip()
        except FileNotFoundError:
            return "Light"
        
        
    '''
    Updates the display mode anytime the selection is changed within the app through
    the display mode selection box. Saves the mode selection to a local .txt file.

    @param app - the main window that is getting the styling adjustment
    @param bottomButton - the QPushButton widget that is getting adjusted 
    @param displayMode - the name of the display mode selected

    @author Mike
    '''
    def changeDisplayStyle(self, app, runButton, pauseButton, stopButton, displayMode):
        if displayMode == "Dark":
            app.setStyleSheet(dark)
        else:
            app.setStyleSheet(light)
        self.saveDisplayMode(displayMode)
        self.updateBottomButtonStyle(runButton, pauseButton, stopButton, displayMode)
    
    
    def updateBottomButtonStyle(self, runButton, pauseButton, stopButton, displayMode):
        if displayMode == "Light":
            runButton.setStyleSheet("""
                QPushButton {
                    background-color: #669df2;  
                    color: #111111;            
                    border: 1px solid #111111; 
                }
                QPushButton:hover {
                    background-color: #0078D4; 
                }
                QPushButton:disabled {
                    background-color: #F0F0F0; 
                    color: #686868; 
                    border: 1px solid #686868;
                }
                """)
            pauseButton.setStyleSheet("""
                QPushButton {
                    background-color: #669df2;  
                    color: #111111;            
                    border: 1px solid #111111; 
                }
                QPushButton:hover {
                    background-color: #0078D4; 
                }
                QPushButton:disabled {
                    background-color: #F0F0F0; 
                    color: #686868; 
                    border: 1px solid #686868;
                }
                """)
            stopButton.setStyleSheet("""
                QPushButton {
                    background-color: #669df2;  
                    color: #111111;            
                    border: 1px solid #111111; 
                }
                QPushButton:hover {
                    background-color: #0078D4; 
                }
                QPushButton:disabled {
                    background-color: #F0F0F0; 
                    color: #686868; 
                    border: 1px solid #686868;
                }
                """)
        else:
            runButton.setStyleSheet("""
                QPushButton {
                    background-color: #0078D4; 
                    color: #FFFFFF;           
                    border: 1px solid #8A8A8A; 
                }
                QPushButton:hover {
                    background-color: #669DF2; 
                }
                QPushButton:disabled {
                    background-color: #1A1A1C; 
                    border: 1px solid #3B3B3B;
                    color: #3B3B3B;   
                }
                """)
            pauseButton.setStyleSheet("""
                QPushButton {
                    background-color: #0078D4; 
                    color: #FFFFFF;           
                    border: 1px solid #8A8A8A; 
                }
                QPushButton:hover {
                    background-color: #669DF2; 
                }
                QPushButton:disabled {
                    background-color: #1A1A1C; 
                    border: 1px solid #3B3B3B;
                    color: #3B3B3B;   
                }
                """)
            stopButton.setStyleSheet("""
                QPushButton {
                    background-color: #0078D4; 
                    color: #FFFFFF;           
                    border: 1px solid #8A8A8A; 
                }
                QPushButton:hover {
                    background-color: #669DF2; 
                }
                QPushButton:disabled {
                    background-color: #1A1A1C; 
                    border: 1px solid #3B3B3B;
                    color: #3B3B3B;   
                }
                """)
        
        
    def mainLayout(self):
        centralWidget = QWidget(self)
        centralWidget.setObjectName("mainCentralWidget")
        self.setCentralWidget(centralWidget)
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

        # Create a group box for the middle top left
        middleTopLeftGroupBox = QGroupBox("Parameters")
        middleTopLeftGroupBox.setFixedSize(375, 375)
        middleTopLeftLayout = QFormLayout()
        self.cemeteryBox = QLineEdit()
        self.cemeteryBox.setFixedWidth(140)
        self.letterBox = QLineEdit()
        self.letterBox.setFixedWidth(30)
        font = QFont("Monterchi Sans Book", 8)
        noteLabel1 = QLabel("Note: This is the cemetery name based on its folder.")
        noteLabel2 = QLabel("Note: This is the last name letter to start processing.")
        noteLabel1.setFont(font)
        noteLabel2.setFont(font)
        self.displayModeBox = QComboBox()
        self.displayModeBox.setFixedWidth(75)
        modes = ["Light", "Dark"]
        self.displayModeBox.addItems(modes)
        middleTopLeftLayout.addRow(" ", None)
        middleTopLeftLayout.addRow("Display Mode: ", self.displayModeBox)
        middleTopLeftLayout.addRow(" ", None)
        middleTopLeftLayout.addRow(" ", None)
        middleTopLeftLayout.addRow(" ", None)
        middleTopLeftLayout.addRow("Cemetery :   ", self.cemeteryBox)
        middleTopLeftLayout.addRow(None, noteLabel1)
        middleTopLeftLayout.addRow(" ", None)
        middleTopLeftLayout.addRow(" ", None)
        middleTopLeftLayout.addRow("Letter :", self.letterBox)
        middleTopLeftLayout.addRow(None, noteLabel2)
        middleTopLeftLayout.addRow(" ", None)
        middleTopLeftLayout.addRow(" ", None)
        middleTopLeftGroupBox.setLayout(middleTopLeftLayout)
        
        # Create a group box for the middle top right
        middleTopRightGroupBox = QGroupBox("Results")
        middleTopRightGroupBox.setFixedSize(535, 375)
        middleTopRightLayout = QFormLayout()
        self.idLabel = QLabel("")
        self.kvsLabel = QLabel("")
        self.errorLabel = QLabel("")
        self.imageLabel = QLabel("")
        middleTopRightLayout.addRow(" ", None)
        middleTopRightLayout.addRow("Current ID :", self.idLabel)
        middleTopRightLayout.addRow(" ", None)
        middleTopRightLayout.addRow("Key-Value Pairs :", self.kvsLabel)
        middleTopRightLayout.addRow(" ", None)
        middleTopRightLayout.addRow("Error :", self.errorLabel)
        middleTopRightLayout.addRow(" ", None)
        middleTopRightGroupBox.setLayout(middleTopRightLayout)

        # Create container for middle top left and top right group boxes
        middleTopContainer = QWidget()
        middleTopLayout = QHBoxLayout()
        middleTopLayout.addWidget(middleTopLeftGroupBox)
        middleTopLayout.addWidget(middleTopRightGroupBox)
        middleTopContainer.setLayout(middleTopLayout)
        
        # Create a group box for the middle bottom
        middleBottomGroupBox = QGroupBox("Image")
        middleBottomGroupBox.setFixedSize(500, 287)
        middleBottomLayout = QFormLayout()
        self.imageLabel = QLabel("")
        middleBottomLayout.addRow(None, self.imageLabel)
        middleBottomGroupBox.setLayout(middleBottomLayout)
        
        # Create a container for the middle bottom group box
        middleBottomContainer = QWidget()
        middleBottomLayout = QHBoxLayout()
        middleBottomLayout.addWidget(middleBottomGroupBox)
        middleBottomContainer.setLayout(middleBottomLayout)
        
        # Create the bottom container with a button
        bottomContainer = QWidget()
        bottomLayout = QHBoxLayout()
        self.runButton = QPushButton("Run Code")
        self.runButton.clicked.connect(self.startProcessing)
        self.runButton.setFixedWidth(150)
        self.pauseButton = QPushButton("Pause Code")
        self.pauseButton.clicked.connect(self.togglePauseResume)
        self.pauseButton.setFixedWidth(150)
        self.stopButton = QPushButton("Stop Code")
        self.stopButton.setFixedWidth(225)
        self.stopButton.clicked.connect(self.stopProcessing)
        self.status = QLabel("              Status :  Idle")
        self.status.setFixedWidth(150)
        self.pauseButton.setDisabled(True)
        self.stopButton.setDisabled(True)
        bottomLayout.addWidget(self.runButton)
        bottomLayout.addWidget(self.status)
        bottomLayout.addWidget(self.pauseButton)
        self.displayModeBox.currentTextChanged.connect(lambda: self.changeDisplayStyle(self, self.runButton, self.pauseButton, self.stopButton, self.displayModeBox.currentText()))
        self.displayModeBox.setCurrentIndex(modes.index(self.loadDisplayMode()))
        self.updateBottomButtonStyle(self.runButton, self.pauseButton, self.stopButton, self.loadDisplayMode())
        bottomContainer.setLayout(bottomLayout)
        # Create a main layout to arrange the containers
        mainLayout = QVBoxLayout()
        mainLayout.addWidget(topContainer)
        mainLayout.addWidget(middleTopContainer)
        mainLayout.addWidget(middleBottomContainer)
        mainLayout.addWidget(bottomContainer)
        mainLayout.addWidget(self.stopButton, 0, alignment=Qt.AlignmentFlag.AlignCenter)
        layout1.addLayout(mainLayout)
        
    def updateStatus(self, type, text):
        if type == "Critical":
            QMessageBox.critical(window, 'Error', text)
        elif type == "Info":
            QMessageBox.information(window, 'Information', text)
        self.runButton.setDisabled(False)
        self.status.setText("              Status :  Idle")
    
    def updateImage(self, filePath):
        pixmap = QPixmap(filePath)
        self.imageLabel.setPixmap(pixmap.scaled(430, 500, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)) 
        
    def updateID(self, id):
        self.idLabel.setText(id)
    
    def updateKVS(self, printedKVS):
        self.kvsLabel.setText(f"{printedKVS}")  
    
    def updateError(self, errorMsg):
        self.errorLabel.setText(f"{errorMsg}")  
          
    def startProcessing(self):
        if self.cemeteryBox.text() == "":
            QMessageBox.warning(window, 'Missing Info', "Cemetery field not filled out.")
            return
        if self.letterBox.text() == "":
            QMessageBox.warning(window, 'Missing Info', "Letter field not filled out.")
            return
        self.worker = Worker(False, self.cemeteryBox.text(), self.letterBox.text())
        self.worker.id_signal.connect(lambda id: self.updateID(id))
        self.worker.kvs_signal.connect(lambda printedKVS: self.updateKVS(printedKVS))
        self.worker.error_signal.connect(lambda errorMsg: self.updateError(errorMsg))
        self.worker.image_signal.connect(lambda imagePath: self.updateImage(imagePath))
        self.worker.popup_signal.connect(lambda type, text: self.updateStatus(type, text))
        self.worker.start()
        self.status.setText("          Status :  Running...")
        self.runButton.setDisabled(True)
        self.pauseButton.setDisabled(False)
        self.stopButton.setDisabled(False)
    
    def togglePauseResume(self):
        if self.worker.paused:
            self.stopButton.setDisabled(False)
            self.worker.paused = False
            self.pauseButton.setText("Pause")
            self.status.setText("          Status :  Running...")
        else:
            self.stopButton.setDisabled(True)
            self.worker.paused = True
            self.pauseButton.setText("Resume")
            self.status.setText("           Status :  Paused" )
    
    def stopProcessing(self):
        self.worker.stopped = True
        self.status.setText("          Status :  Stopping...")
        self.runButton.setDisabled(False)
        
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.setWindowTitle("Veteran Extraction")
    parent_path = os.path.dirname(os.getcwd())
    window.setWindowIcon(QIcon(f"{parent_path}/veteranData/veteranLogo.png"))
    window.show()
    QMessageBox.information(window, 'Instructions', "If code is running, please press \"Stop Code\"\nbefore closing the application.")
    sys.exit(app.exec())