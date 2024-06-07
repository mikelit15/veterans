import cv2
import easyocr
from PIL import Image
import numpy as np
import openpyxl
import pytesseract
import os
import aspose.words as aw
from reportlab.pdfgen import canvas
import PyPDF2

def preProcess(imgPath, flag):
    img = cv2.imread(imgPath)
    pt1 = (5340, 1400)
    pt3 = (7055, 1700)
    # if flag:
    #     pt1 = (5400, 680)
    #     pt3 = (7000, 980)
    cv2.rectangle(img, pt1, pt3, (0, 0, 0), thickness=cv2.FILLED)
    cv2.imwrite(imgPath, img)
    image = Image.open(imgPath)
    pdfFile = "output.pdf"
    c = canvas.Canvas(pdfFile, pagesize=(image.width, image.height))
    c.drawImage(imgPath, 0, 0, width=image.width, height=image.height)
    c.save()

    img2 = cv2.imread(imgPath)
    gray = cv2.cvtColor(img2, cv2.COLOR_BGR2GRAY)
    thresh = cv2.adaptiveThreshold(~gray, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY, 9, -2)
    kernal = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 5))
    thresh = cv2.dilate(thresh, kernal, iterations=3)
    thresh = cv2.erode(thresh, kernal, iterations=1)
    horz = np.copy(thresh)
    (_, horzcol) = np.shape(horz)
    horzSize = int(horzcol/26)
    horzStructure = cv2.getStructuringElement(cv2.MORPH_RECT, (horzSize, 1))
    horz = cv2.erode(horz, horzStructure, iterations=1)
    horz = cv2.dilate(horz, horzStructure, iterations=1)
    outimage = cv2.bitwise_and(~gray, ~horz)
    outimage = ~outimage
    edges = cv2.Canny(outimage, 50, 150, apertureSize=3)
    minLineLength = 5
    maxLineGap = 10
    lines = cv2.HoughLinesP(edges, 1, np.pi/180, 50, minLineLength, maxLineGap)
    for x1, y1, x2, y2 in lines[0]:
        cv2.line(outimage, (x1, y1), (x2, y2), (0, 255, 0), 2)
    cv2.imwrite(f'{imgPath.replace(".png", "")} edit.png', outimage)

def easyOCR(imagePath):
    reader = easyocr.Reader(['en'])
    text = reader.readtext(imagePath)
    coords = []
    keys = [[1, "name"], [6, "born"], [20, "record"], [21, "service"]]
    for x in text:
        print(x)
    prev_bbox = None  
    currentWord = "" 
    current_bbox = None
    # Loop through the OCR results
    for key in keys:
        found = False 
        for bbox, word, score in text:
            if (prev_bbox and 
                    0 <= (bbox[0][0] - prev_bbox[1][0]) <= 70 and 
                    (bbox[0][1] - prev_bbox[1][1]) <= 10):                    
                    currentWord += " " + word
                    current_bbox[1] = bbox[1]
                    current_bbox[2] = bbox[2]
            else:
                if key[1].lower() in currentWord.lower() and ("REGISTRATION" not in currentWord):
                    if bbox[0][1] > 600:
                        coords.append([currentWord.upper(), current_bbox])
                        found = True
                        break
                currentWord = word
                current_bbox = bbox.copy()
            prev_bbox = bbox
        if not found:
            dummy_coords = [[0, 0], [0, 0], [0, 0], [0, 0]]
            coords.append([key[1].upper(), dummy_coords])

    return coords

def tesseractOCR(imagePath, row_index, coords):
    image = Image.open(imagePath)
    count = 0
    counter = 1
    print(coords)
    # nt = image.crop((685, 668, 4500, 915))
    # print(pytesseract.image_to_string(nt, lang='eng'))
    offset = [3700, 1481, 5875, 2696]
    for x in coords: 
        cropped_image = image.crop((x[1][2][0] - 15, x[1][0][1] - 120, x[1][1][0] + offset[count], x[1][2][1] - 9))
        # cropped_image = image.crop((219, 452, 413, 520))
        text = pytesseract.image_to_string(cropped_image, lang='eng')
        filtered_text = text.replace("♀", "").replace("\n", "").replace("«", "") \
                        .replace("_", " ").replace("—. ", " ").replace("(", " ") \
                        .replace("CEMETERY", "").strip()
        worksheet.cell(row=row_index, column=counter, value=filtered_text)
        print(filtered_text)
        counter += 1
        count += 1
    os.remove(imagePath)

# Loop through all images
workbook = openpyxl.load_workbook('Veterans.xlsx')
worksheet = workbook['Evergreen']
current_directory = os.getcwd()
row_index = 2
pytesseract.pytesseract.tesseract_cmd = r'C:\Users\michael.litvin\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'
os.environ['TESSDATA_PREFIX'] = r'C:\Users\michael.litvin\Downloads'
name_coordinates = []
png_files = [file for file in os.listdir(current_directory) if file.lower().endswith('.png')]
for x in range(len(png_files)):
    if x % 2 != 0:
        continue
    elif "output" not in png_files[x] and "edit" not in png_files[x]:
        preProcess(png_files[x], png_files[x+1])
    
edit_files = [file for file in os.listdir(current_directory) if file.lower().endswith('edit.png')]
for x in edit_files:
    name_coordinates = easyOCR(x)
    tesseractOCR(x, row_index, name_coordinates)
    row_index += 1
workbook.save('Veterans.xlsx')

# Process one image
# imagePath = "00009 edit.png"
# image = Image.open(imagePath)
# current_directory = os.getcwd()
# preProcess(imagePath)
# nt = image.crop((1792, 2739, 2307, 2903))
# print(pytesseract.image_to_string(nt, lang='eng'))

# name_coordinates = easyOCR(imagePath)
# tesseractOCR(imagePath, row_index, name_coordinates)
