import fitz
import os
import time

inputDirectory = r"\\ucclerk\pgmdoc\Veterans\Cemetery - Redacted - Before Scaling"
outputDirectory = r"\\ucclerk\pgmdoc\Veterans\Cemetery - Redacted - After Scaling"

targetWidth = 842  
targetHeight = 560

retryDelay = 3 

if not os.path.exists(outputDirectory):
    os.makedirs(outputDirectory)

def resizePDF_WithRetry(inputPDF_Path, outputPDF_Path, targetWidth, targetHeight):
    while True: 
        try:
            PDF_Document = fitz.open(inputPDF_Path)
            outputPDF = fitz.open()
            for pageNum in range(len(PDF_Document)):
                page = PDF_Document.load_page(pageNum)
                rect = page.rect
                newPage = outputPDF.newPage(width=targetWidth, height=targetHeight)
                scaleX = targetWidth / rect.width
                scaleY = targetHeight / rect.height
                scaleFactor = min(scaleX, scaleY)
                scaledWidth = rect.width * scaleFactor
                scaledHeight = rect.height * scaleFactor
                xOffset = (targetWidth - scaledWidth) / 2
                yOffset = (targetHeight - scaledHeight) / 2
                matrix = fitz.Matrix(scaleFactor, scaleFactor).pretranslate(xOffset, yOffset)
                newPage.show_pdf_page(newPage.rect, PDF_Document, pageNum, matrix)
            outputPDF.save(outputPDF_Path)
            outputPDF.close()
            PDF_Document.close()
            print(f"Successfully resized: {inputPDF_Path}")
            break 
        except Exception as e:
            print(f"Error processing {inputPDF_Path}: {e}")
            print(f"Retrying in {retryDelay} seconds...")
            time.sleep(retryDelay)  

def processDirectory(inputDir, outputDir):
    for root, dirs, files in os.walk(inputDir):
        for fileName in files:
            if fileName.endswith(".pdf"):
                inputPDF_Path = os.path.join(root, fileName)
                relativePath = os.path.relpath(inputPDF_Path, inputDir)
                outputPDF_Path = os.path.join(outputDir, relativePath)
                outputSubdir = os.path.dirname(outputPDF_Path)
                if not os.path.exists(outputSubdir):
                    os.makedirs(outputSubdir)
                print(f"Resizing {inputPDF_Path} to {targetWidth}x{targetHeight} points (295mm x 205mm)...")
                resizePDF_WithRetry(inputPDF_Path, outputPDF_Path, targetWidth, targetHeight)

processDirectory(inputDirectory, outputDirectory)