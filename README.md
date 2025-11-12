# Veteran Grave Registration Card Processor

A comprehensive automated system for processing scanned PDF documents of Veteran Grave Registration cards. This application leverages Microsoft Azure OCR technology to extract, clean, validate, and organize veteran data into structured Excel spreadsheets while maintaining privacy compliance through automated redaction of sensitive information.

## 📋 Table of Contents

- [Overview](#overview)
- [Key Features](#key-features)
- [Technology Stack](#technology-stack)
- [Project Structure](#project-structure)
- [Installation](#installation)
- [Applications](#applications)
  - [microsoftOCR GUI](#microsoftocr-gui)
  - [multiCleaner GUI](#multicleaner-gui)
- [Core Modules](#core-modules)
- [Data Processing Pipeline](#data-processing-pipeline)
- [Utilities](#utilities)
- [Testing](#testing)
- [Configuration](#configuration)
- [Network Folder Structure](#network-folder-structure)

## 🎯 Overview

This system automates the digitization of historical Veteran Grave Registration cards, transforming scanned PDF documents into structured, searchable data. The application processes thousands of records across multiple cemeteries, applying sophisticated OCR technology, intelligent data cleaning algorithms, and privacy-compliant redaction techniques.

### Primary Objectives
- **Automate Data Extraction**: Extract text from scanned PDF documents using custom-trained Azure OCR models
- **Ensure Data Quality**: Apply field-specific cleaning rules to standardize names, dates, and military service information
- **Maintain Privacy Compliance**: Automatically redact sensitive information (Social Security numbers)
- **Organize Data Efficiently**: Structure data in Excel spreadsheets with hyperlinks to source documents
- **Enable Historical Research**: Provide accessible, searchable veteran records for administrative and research purposes

## ✨ Key Features

### Intelligent OCR Processing
- **Custom Azure OCR Model**: Trained on 100+ annotated images with bounding boxes for precise field extraction
- **High-Resolution Processing**: Converts PDF pages to high-quality images for optimal OCR accuracy
- **Field-Specific Extraction**: Identifies and extracts 11+ distinct fields from registration cards
- **Error Detection**: Comprehensive error logging and validation for quality assurance

### Advanced Data Cleaning
- **Name Standardization**: Handles multi-part surnames, prefixes (O', Mc, Van, etc.), and suffixes (Jr., Sr., I-IV)
- **Date Normalization**: Converts various date formats into standardized MM/DD/YYYY format
- **Military Service Formatting**: Interprets historical military abbreviations and service branch codes
- **War Record Processing**: Standardizes war/conflict designations using historical knowledge

### Privacy & Security
- **Automated Redaction**: Dynamically calculates and applies black boxes over Social Security numbers
- **Coordinate-Based Masking**: Uses OCR bounding box coordinates for precise redaction placement
- **Dual-Page PDF Creation**: Combines redacted first page with original second page
- **Secure File Management**: Organizes redacted files in separate directory structures

### Data Management
- **Excel Integration**: Populates structured workbooks with validated data
- **Hyperlink Generation**: Creates clickable links from Excel rows to source PDF documents
- **Duplicate Detection**: Identifies potential duplicate records based on name and date matching
- **Batch Processing**: Handles entire cemetery directories organized alphabetically (A-Z)

### User Interface
- **PySide6 GUI**: Modern, responsive interface with dark/light mode support
- **Real-Time Progress Monitoring**: Displays current record ID, extracted values, and errors
- **Process Control**: Pause, resume, and safe-stop functionality to prevent data corruption
- **Visual Feedback**: Shows redacted images and processing status in real-time

## 🛠 Technology Stack

### Core Technologies
- **Python 3.x**: Primary programming language
- **Microsoft Azure Form Recognizer**: Custom OCR model for document analysis
- **PySide6 (Qt)**: Cross-platform GUI framework
- **OpenPyXL**: Excel file manipulation and formatting
- **PyMuPDF (fitz)**: PDF document processing and manipulation
- **OpenCV (cv2)**: Image processing and redaction
- **ReportLab**: PDF generation for redacted documents

### Data Processing Libraries
- **pandas**: Data analysis and duplicate detection
- **numpy**: Numerical operations for data processing
- **dateparser**: Intelligent date parsing and normalization
- **nameparser**: Name parsing and formatting
- **fuzzywuzzy**: Fuzzy string matching for name validation
- **PyPDF2**: PDF reading and metadata extraction

### Additional Dependencies
- **PIL (Pillow)**: Image format conversion
- **qdarktheme**: Modern dark theme styling for Qt applications
- **traceback**: Detailed error logging and debugging

## 📁 Project Structure

```
veterans/
├── microsoftOCR/                    # Main OCR processing application
│   ├── GUI.py                      # PySide6 GUI for OCR application (788 lines)
│   ├── microsoftOCR.py             # Core OCR processing engine (899 lines)
│   ├── microsoftOCR_Old.py         # Previous version for reference (946 lines)
│   ├── nameRule.py                 # Name parsing and formatting rules (313 lines)
│   ├── nameRuleOld.py              # Legacy name processing (163 lines)
│   ├── dateRule.py                 # Date parsing and normalization (1,054 lines)
│   ├── warRule.py                  # War/conflict standardization (97 lines)
│   └── branchRule.py               # Military branch formatting (105 lines)
│
├── utilities/                       # Data management and cleaning utilities
│   ├── GUI.py                      # PySide6 GUI for utilities (1,025 lines)
│   ├── multiCleaner.py             # Batch file organization tool (505 lines)
│   ├── multiCleaner2.0.py          # Enhanced version (471 lines)
│   ├── multiCleanerOld.py          # Legacy version (379 lines)
│   ├── singleCleaner.py            # Individual file processing (337 lines)
│   ├── duplicates.py               # Duplicate record detection (102 lines)
│   ├── cleanerImage.py             # Image file name cleaning (121 lines)
│   ├── cleanerHyperlink.py         # Excel hyperlink management (51 lines)
│   ├── cleanDelete.py              # File deletion with Excel sync (37 lines)
│   ├── cleanRedacted.py            # Redacted file cleanup (46 lines)
│   ├── incrementRedacted.py        # Redacted file numbering (89 lines)
│   ├── hyperlinkFix.py             # Hyperlink repair utility (17 lines)
│   ├── pdfScaling.py               # PDF dimension analysis (55 lines)
│   └── miscFunctions.py            # Helper functions (114 lines)
│
├── testFiles/                       # Testing and debugging scripts
│   ├── microsoftOCRTest.py         # OCR functionality tests
│   ├── multiCleanerTest.py         # Cleaner function tests
│   ├── dateTest.py                 # Date parsing tests
│   ├── nameTest.py                 # Name formatting tests
│   ├── warTest.py                  # War record tests
│   ├── duplicatesTest.py           # Duplicate detection tests
│   └── test.py                     # General testing script
│
├── prevVersions/                    # Historical implementations
│   ├── amazonOCR.py                # AWS Textract implementation
│   ├── amazonOCR orig.py           # Original AWS version
│   ├── hardCoordinates.py          # Fixed-coordinate redaction
│   ├── keyCoordiantes.py           # Key-based coordinate mapping
│   └── keyOffset.py                # Offset calculation methods
│
├── veteranData/                     # Application resources
│   ├── logo.png                    # Application logo (154 KB)
│   ├── veteranLogo.png             # Veteran-specific branding (907 KB)
│   └── display_mode.txt            # UI theme preference storage
│
└── README.md                        # This documentation file
```

### Code Statistics
- **Total Python Files**: 30+
- **Total Lines of Code**: 7,714+ (core modules only)
- **Main Application**: microsoftOCR (899 lines)
- **Largest Module**: dateRule.py (1,054 lines)
- **GUI Components**: 2 major applications (788 + 1,025 lines)

## 🚀 Installation

### Prerequisites
- Python 3.8 or higher
- Microsoft Azure account with Form Recognizer service
- Access to network folders (\\ucclerk\pgmdoc\Veterans\)
- Windows operating system (for network path compatibility)

### Required Python Packages

```bash
pip install azure-ai-formrecognizer azure-core
pip install PySide6 qdarktheme
pip install openpyxl pandas numpy
pip install PyMuPDF PyPDF2 Pillow opencv-python reportlab
pip install dateparser nameparser fuzzywuzzy python-Levenshtein
```

### Azure Configuration
1. Create an Azure Form Recognizer resource
2. Train a custom model with veteran registration card samples
3. Configure credentials in the application (endpoint and API key)

### Network Setup
Ensure access to the following network paths:
- `\\ucclerk\pgmdoc\Veterans\Cemetery\` - Source PDF files
- `\\ucclerk\pgmdoc\Veterans\Cemetery - Redacted\` - Redacted output files
- Excel workbook location for data storage

## 📱 Applications

### microsoftOCR GUI

The primary application for processing veteran registration cards with OCR technology.

#### Features and Usage

**1. Application Launch**
- **Initial Pop-Up**: Warning message advising users to click "Stop Code" before exiting to prevent data corruption
- **Main Interface**: Displays processing status with four states:
  - **Idle**: Ready for input, awaiting "Run Code" button press
  - **Running...**: Actively processing PDF files and extracting data
  - **Paused**: Processing halted, awaiting "Resume" button press
  - **Stopping...**: Safely finishing current record before exit

**2. Required Parameters**
- **Cemetery Name**: Select from dropdown or enter cemetery name
- **Letter Folder**: Specify alphabetical folder (A-Z) to process
- **Validation**: Pop-ups alert if parameters are missing

**3. Processing Workflow**
- **Run Code**: Initiates `microsoftOCR.py` execution
- **Progress Display**:
  - Current record ID being processed
  - Raw key-value pairs extracted from OCR
  - Detailed error messages for validation failures
  - Redacted image preview after processing
- **Completion**: Pop-up notification when all files in folder are processed

**4. Control Buttons**
- **Pause/Resume**: Halt processing to review data, then continue
- **Stop**: Safely complete current record and save Excel before stopping
- **Run Code**: Start batch processing of selected cemetery/letter folder

#### Extracted Fields
- Veteran ID (VID)
- Last Name, First Name, Middle Name, Suffix
- Date of Birth (Month, Day, Year)
- Date of Death (Month, Day, Year)
- War/Conflict
- Branch of Service
- Serial Number (redacted in output)
- Cemetery Location
- Grave/Plot Information

### multiCleaner GUI

Utility application for organizing, cleaning, and managing processed files.

#### Features
- **Batch File Renaming**: Ensures sequential numbering across cemetery folders
- **Redacted File Management**: Handles 'a' and 'b' suffix files for duplicate cards
- **Hyperlink Synchronization**: Updates Excel hyperlinks after file reorganization
- **Duplicate Detection**: Identifies potential duplicate veteran records
- **File Cleanup**: Removes unnecessary or incorrectly processed files

#### Workflow
1. **Select Operation**: Choose from cleaning, renaming, or duplicate detection
2. **Specify Parameters**: Cemetery name, starting ID, folder range
3. **Execute**: Process files with real-time progress feedback
4. **Review**: Check logs and error reports for issues

## 🔧 Core Modules

### microsoftOCR.py (899 lines)
**Primary OCR processing engine**

**Key Functions**:
- `redact(filePath, cemetery, letter, nameCoords, serialCoords, ...)`: Applies black box redaction to Social Security numbers
- `mergeRedacted(pathA, pathB, cemetery, letter)`: Combines multiple versions of registration cards
- `processDocument(filePath, cemetery, letter)`: Main processing pipeline for each PDF
- `extractFields(ocrResult)`: Parses OCR output into structured fields
- `validateData(fields)`: Applies validation rules to extracted data
- `saveToExcel(fields, workbook, sheet)`: Writes cleaned data to Excel

**Processing Steps**:
1. Load PDF and convert first page to high-resolution image
2. Submit image to Azure Form Recognizer custom model
3. Extract bounding box coordinates and text values
4. Apply field-specific cleaning rules
5. Validate data against business rules
6. Redact sensitive information using coordinate-based masking
7. Save data to Excel with hyperlinks
8. Archive redacted PDF to network folder

### nameRule.py (313 lines)
**Name parsing and standardization**

**Capabilities**:
- Handles comma-separated and period-separated name formats
- Recognizes multi-part surname prefixes (O', Mc, Van, De, Di, Del, Le, St.)
- Processes suffixes (Jr., Sr., I, II, III, IV)
- Capitalizes names correctly (including special cases like McDonald, O'Brien)
- Removes parenthetical notes and invalid characters
- Splits full names into: Last, First, Middle, Suffix

**Example Transformations**:
- `"O'BRIEN. JOHN MICHAEL JR"` → Last: O'Brien, First: John, Middle: Michael, Suffix: Jr.
- `"MCDONALD, ROBERT J"` → Last: McDonald, First: Robert, Middle: J
- `"VAN DYKE WILLIAM"` → Last: Van Dyke, First: William

### dateRule.py (1,054 lines)
**Date parsing and normalization**

**Handles Multiple Formats**:
- `MM/DD/YYYY`, `M/D/YY`, `MM-DD-YYYY`
- `Month DD, YYYY` (e.g., "January 15, 1945")
- `Month YYYY` (partial dates)
- Two-digit years with century inference
- Malformed dates with extra characters

**Validation**:
- Checks for valid month (1-12), day (1-31), year ranges
- Flags impossible dates (e.g., February 30)
- Handles partial dates (year only, month/year only)
- Standardizes output to `MM/DD/YYYY` format

### warRule.py (97 lines)
**Military conflict standardization**

**Recognized Conflicts**:
- World War I (WWI, WW1, World War 1)
- World War II (WWII, WW2, World War 2)
- Korean War (Korea)
- Vietnam War (Vietnam, Viet Nam)
- Gulf War, Iraq War, Afghanistan
- Spanish-American War
- Civil War

**Standardization**: Converts abbreviations and variations to consistent format

### branchRule.py (105 lines)
**Military branch formatting**

**Branches**:
- Army (USA, U.S. Army)
- Navy (USN, U.S. Navy)
- Air Force (USAF, U.S. Air Force)
- Marines (USMC, U.S. Marine Corps)
- Coast Guard (USCG)

**Formatting**: Standardizes abbreviations and full names

## 🔄 Data Processing Pipeline

### End-to-End Workflow

```
1. PDF Input
   ↓
2. Image Conversion (High-Resolution PNG)
   ↓
3. Azure OCR Analysis (Custom Model)
   ↓
4. Field Extraction (11+ fields with bounding boxes)
   ↓
5. Data Cleaning
   ├── Name Standardization (nameRule.py)
   ├── Date Normalization (dateRule.py)
   ├── War Formatting (warRule.py)
   └── Branch Formatting (branchRule.py)
   ↓
6. Data Validation
   ├── Required field checks
   ├── Format validation
   └── Business rule enforcement
   ↓
7. Sensitive Data Redaction
   ├── Coordinate calculation
   ├── Black box overlay
   └── Redacted PDF creation
   ↓
8. Excel Integration
   ├── Data insertion
   ├── Hyperlink creation
   ├── Conditional formatting
   └── Workbook save
   ↓
9. File Archiving
   ├── Organize by cemetery
   ├── Organize by last name letter
   └── Network folder storage
```

### Error Handling Strategy

- **Try-Catch Blocks**: Wrap each processing step to capture exceptions
- **Detailed Logging**: Record errors with file path, field name, and stack trace
- **Graceful Degradation**: Continue processing remaining files after errors
- **User Notification**: Display errors in GUI for review
- **Excel Highlighting**: Mark problematic rows with colored cells

## 🛠 Utilities

### duplicates.py
**Identifies potential duplicate veteran records**

**Detection Criteria**:
- Matching last name + first name + death year
- Excludes records in exception list
- Groups duplicates for manual review

**Output**: DataFrame with duplicate groups and row numbers

### multiCleaner.py
**Batch file organization and renaming**

**Functions**:
- `cleanImages()`: Processes all cemetery folders
- `processCemetery()`: Handles A-Z letter subfolders
- `clean()`: Renames files sequentially
- `adjustImageName()`: Handles 'a'/'b' suffix files
- `cleanDelete()`: Removes files and updates Excel

### cleanerHyperlink.py
**Excel hyperlink management**

**Purpose**: Updates hyperlink formulas after file renaming
**Process**: Adjusts numeric IDs in hyperlink paths to match new file names

### singleCleaner.py
**Individual file processing utility**

**Use Cases**:
- Reprocess single problematic file
- Test cleaning rules on specific record
- Manual correction workflow

## 🧪 Testing

### Test Files Overview

**microsoftOCRTest.py**: Comprehensive OCR pipeline testing
**dateTest.py**: 100+ date format test cases
**nameTest.py**: Name parsing edge cases
**warTest.py**: Military conflict standardization tests
**duplicatesTest.py**: Duplicate detection algorithm validation

### Running Tests

```bash
# Run all tests
python testFiles/microsoftOCRTest.py
python testFiles/dateTest.py
python testFiles/nameTest.py

# Test specific functionality
python testFiles/test.py
```

## ⚙️ Configuration

### Azure Credentials
Configure in `microsoftOCR.py`:
```python
endpoint = "https://your-resource.cognitiveservices.azure.com/"
api_key = "your-api-key"
model_id = "your-custom-model-id"
```

### Network Paths
Default paths (configurable):
- Source: `\\ucclerk\pgmdoc\Veterans\Cemetery\`
- Redacted: `\\ucclerk\pgmdoc\Veterans\Cemetery - Redacted\`
- Excel: Network location specified in code

### GUI Theme
Stored in `veteranData/display_mode.txt`:
- `dark`: Dark mode theme
- `light`: Light mode theme

## 📂 Network Folder Structure

```
\\ucclerk\pgmdoc\Veterans\
├── Cemetery\
│   ├── [Cemetery Name]\
│   │   ├── A\
│   │   │   ├── [CemeteryName]A00001.pdf
│   │   │   ├── [CemeteryName]A00002.pdf
│   │   │   └── ...
│   │   ├── B\
│   │   ├── C\
│   │   └── ... (through Z)
│   ├── Jewish\
│   │   ├── [Sub-Cemetery]\
│   │   │   ├── A\
│   │   │   └── ...
│   └── Misc\
│
└── Cemetery - Redacted\
    ├── [Cemetery Name] - Redacted\
    │   ├── A\
    │   │   ├── [CemeteryName]A00001 redacted.pdf
    │   │   └── ...
    │   └── ... (through Z)
    └── ...
```

## 📊 Excel Output Format

### Columns
| Column | Field | Description |
|--------|-------|-------------|
| A | VID | Veteran ID (sequential number) |
| B | VLNAME | Last Name |
| C | VFNAME | First Name |
| D | VMNAME | Middle Name |
| E | VSUFFIX | Suffix (Jr., Sr., etc.) |
| F | VDOBM | Date of Birth - Month |
| G | VDOBD | Date of Birth - Day |
| H | VDOBY | Date of Birth - Year |
| I | VDODM | Date of Death - Month |
| J | VDODD | Date of Death - Day |
| K | VDODY | Date of Death - Year |
| L | VWAR | War/Conflict |
| M | VBRANCH | Branch of Service |
| N | HYPERLINK | Link to redacted PDF |

### Conditional Formatting
- **Yellow Highlight**: Validation warnings
- **Red Highlight**: Critical errors
- **Green Highlight**: Successfully processed

## 🤝 Contributing

This is a production system for government records management. Changes should be:
1. Thoroughly tested in `testFiles/` environment
2. Documented with detailed comments
3. Reviewed for data privacy compliance
4. Validated against historical records

## 📝 Version History

### Current Version (Microsoft OCR)
- Custom Azure Form Recognizer model
- PySide6 GUI with dark/light themes
- Advanced name/date parsing rules
- Coordinate-based redaction

### Previous Versions (prevVersions/)
- **amazonOCR.py**: AWS Textract implementation
- **hardCoordinates.py**: Fixed-position redaction
- **keyOffset.py**: Offset-based coordinate calculation

## 🙏 Acknowledgments

This system digitizes and preserves historical veteran records, honoring those who served. The application balances automation efficiency with data accuracy and privacy protection, ensuring these important records remain accessible for future generations.
