# Veteran Grave Registration Card Processor

This program automates the extraction and cleaning of text from scanned PDF images of Veteran Grave Registration cards, using them to populate a structured Excel sheet. The process includes redacting sensitive information and organizing data efficiently across network folders.

## Project Structure

### Subfolders:
- **microsoftOCR**: Contains the main OCR processing scripts leveraging Microsoft Azure OCR services.
- **prevVersions**: Stores earlier versions of the scripts that used free open-source OCR methods for archival purposes.
- **testFiles**: Includes scripts for testing and debugging specific functions without affecting the production environment.
- **utilities**: Holds utility scripts for deduplication tasks like deduplication.
- **veteranData**: Stores necessary data files and configurations for GUI interactions and script settings.

## microsoftOCR GUI:

### Features and Usage

#### 1. Application Launch
- **Initial Pop-Up**: Upon starting the application, a pop-up window advises the user to click the "Stop Code" button before exiting the application if the code is running. This ensures that all processes are safely concluded to prevent data loss and file corruption.
- **Main Interface**: After dismissing the pop-up, the main application interface is displayed. This interface is the primary area for user interaction and displays the current status of the application at the bottom, which can be in one of several states:
  - **Idle**: The application is ready and waiting for user input and a "Run Code" button press to start processing.
  - **Running...**: Indicates that the code is actively processing data.
  - **Paused**: The application has paused processing and requires the user to press "Resume" to continue.
  - **Stopping...**: The application is finishing up current tasks and preparing to allow the user to safely exit.
    
  ![image1_2_80](https://github.com/mikelit15/veterans/assets/79111823/d6fdaa49-0fa4-456f-9c60-453d2ef22671)

#### 2. Filling in the Required Parameters
  - **Parameter Input**: Users must input parameters such as the cemetery name and the letter of the folder to be processed.

#### 3. Run the Code
  - **Starting the Process**: By pressing the "Run Code" button, the `microsoftOCR.py` code is executed.
  - **Monitoring Progress**:
    - The results box initially displays the ID number being processed.
    - It then shows the RAW key-value pairs extracted from the document.
    - Any errors encountered during the extraction or cleaning of the text are displayed in detail in the error field.
    - Upon completion of data cleaning, the image gets redacted, and the redacted image is then displayed in the app.
    - The results box clears, and the next record begins processing.
  - **Completion**:
    - This process repeats until all records in the specified folder are processed. Upon completion, a pop-up informs the user that all files have been processed, and the app returns to idle status, ready for new inputs.

#### Button Functionality
  - **Pause and Resume**:
    - The pause button halts the processing at any point, allowing users to review the extracted values, error messages, or the redacted image.
    - The button's text switches to "Resume," which can be pressed to continue the process.
  - **Safe Stop**:
    - The stop button completes processing of the current record and ensures the Excel file is saved before terminating the app.
    - This is crucial to prevent data loss or corruption of the Excel file, which can occur if the application is stopped abruptly during processing.

## Detailed Workflow of microsoftOCR.py

### Step 1: Initialization and Configuration
- **Setup**: Initialize the environment by setting up paths to the input files and configuring the OCR service credentials.
- **Configuration Loading**: Load configuration settings that define paths, OCR settings, and rules for data parsing and redaction.

### Step 2: PDF Loading and Image Extraction
- **PDF Loading**: Load the PDF document from a specified directory based on the input parameters.
- **Page Extraction**: Extract the relevant page(s) from the PDF document. This is typically the first page unless specified otherwise in the configurations.
- **Image Conversion**: Convert the extracted page into an image format using a high-resolution setting to ensure that the quality is sufficient for accurate OCR recognition.

### Step 3: Sensitive Information Redaction
- **Detect Sensitive Information**: Scan the image for predefined patterns that indicate sensitive information, such as social security numbers or personal addresses.
- **Calculate Redaction Areas**: Dynamically calculate the coordinates for redaction based on the detected sensitive information.
- **Apply Redaction**: Use image processing techniques to overlay black boxes over these coordinates, ensuring the data cannot be read or recovered.

### Step 4: OCR Processing
- **Image Preprocessing**: Enhance the image quality by adjusting brightness, contrast, and possibly applying filters to reduce noise and improve text clarity.
- **Text Extraction**: Submit the preprocessed, redacted image to the OCR service. This process involves sending the image data to Microsoft Azure, which analyzes the image and returns the text content.
- **OCR Result Parsing**: Parse the raw OCR output to structure the extracted text into a usable format, separating different fields based on the document layout.

### Step 5: Data Cleaning and Validation
- **Initial Data Cleaning**: Apply initial cleaning rules to correct common OCR errors, such as replacing misrecognized characters or splitting joined words.
- **Field-Specific Rules**: Execute sophisticated scripts for each data field, such as:
  - **Name Parsing**: Apply cultural and linguistic rules to handle names correctly.
  - **Date Normalization**: Convert various date formats into a uniform standard format.
  - **Service Details Extraction**: Extract and format military service details using historical and military knowledge to interpret abbreviations and jargon.
- **Validation**: Validate the cleaned data against a set of predefined rules to ensure accuracy and consistency.

### Step 6: Excel Integration
- **Workbook Preparation**: Open an existing Excel workbook or create a new one if it does not exist.
- **Data Mapping**: Map the cleaned and validated data into the appropriate columns in the Excel sheet. Each field from the registration card is carefully placed into designated columns to maintain data integrity and facilitate easy data retrieval.
- **Formatting**: Apply formatting rules to the Excel output to enhance readability and usability, such as setting column widths and applying text formats.

### Step 7: File Management
- **Output Saving**: Save the processed Excel file with a timestamped filename to ensure version control and traceability.
- **Image Archiving**: Archive the redacted image files in a structured directory system that categorizes images by cemetery name and veteran's last initial for efficient organization and retrieval.
- **Logging**: Maintain comprehensive logs of each processing step for auditing and troubleshooting purposes.

## Error Handling
- **Robust Error Management**: Implement robust try-catch blocks around each major processing step to capture and log exceptions.
- **Continuity Assurance**: Ensure that the occurrence of an error in processing one file does not halt the entire batch process, allowing the system to move on to the next file.

## Conclusion
This highly detailed automated system not only significantly reduces manual labor but also increases the accuracy and privacy compliance of managing veteran grave registration data. Through sophisticated OCR technology and meticulous data handling procedures, the program ensures that the data is processed with high precision and security.

## multiCleaner GUI:

### Features and Usage

#### 1. Application Launch


## Detailed Workflow of multiCleaner.py

### Step 1: Initialization
- Import necessary libraries and modules.
- Configure system path to include custom modules like `microsoftOCR`.

### Step 2: File Name Cleaning
- **Function**: `cleanImages()`
- **Purpose**: To iterate through each cemetery directory and initiate cleaning of the image file names.
- **Process**:
  - Define the base directory for cemetery files.
  - List all subdirectories to handle each cemetery individually.
  - For standard cemeteries, directly call `process_cemetery()`.
  - For cemeteries with nested subdirectories like 'Jewish' and 'Misc', process each sub-cemetery individually using `process_cemetery()`.

### Step 3: Processing Individual Cemeteries
- **Function**: `process_cemetery()`
- **Purpose**: To manage file renaming within each cemetery's directory structured alphabetically from 'A' to 'Z'.
- **Process**:
  - Generate an alphabetical list for folder navigation.
  - For each letter subfolder, navigate and call the `clean()` function.
  - Handle file renaming ensuring sequential order and address special cases for redacted files.

### Step 4: Detailed File Renaming and Redacted File Handling
- **Function**: `clean()`
- **Purpose**: To rename files within a specific letter directory ensuring a sequential naming order and to handle the removal of unnecessary redacted files.
- **Process**:
  - Retrieve all PDF files in the target directory.
  - Calculate the last used index from the previous letter directory to ensure continuity in naming.
  - Rename files incrementally; remove 'a' and 'b' suffixes from redacted files if necessary.

### Step 5: Adjusting File Names and Deleting Redacted Files
- **Functions**: `adjustImageName()` and `cleanDelete()`
- **Purpose**: Adjust file names based on good and bad IDs, and handle deletion of improperly redacted files while updating Excel records.
- **Process**:
  - Rename files by appending 'a' or 'b' to correct IDs.
  - Move files to correct directories if needed.
  - Update Excel workbook to reflect changes and delete Excel rows corresponding to bad IDs.

### Step 6: Updating Hyperlinks in Excel
- **Function**: `cleanHyperlinks()`
- **Purpose**: To update hyperlink references in the Excel spreadsheet to reflect the current record IDs following modifications.
- **Process**:
  - Adjust the numeric ID within each hyperlink to ensure they accurately reflect the updated record IDs.
  - Ensure all hyperlinks are active and correctly pointing to the associated PDF files.

## Conclusion
This script provides an automated solution to manage and organize Veteran Grave Registration Card images efficiently, ensuring data accuracy and ease of access for administrative and historical research purposes.




  
