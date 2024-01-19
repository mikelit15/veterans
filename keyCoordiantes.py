import cv2
import easyocr
from PIL import Image
import numpy as np
import openpyxl
import pytesseract
import os

def preProcess(image):
    img = cv2.imread(image)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    thresh = cv2.adaptiveThreshold(~gray, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY, 7, -2)
    kernal = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 5))
    thresh = cv2.dilate(thresh, kernal, iterations=2)
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
    cv2.imwrite('output.png', outimage)
    # Load the image
    # image = cv2.imread(image_path)

    # # Convert the image to grayscale
    # gray_image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

    # # Perform text detection using pytesseract
    # custom_config = r'--oem 3 --psm 6'  # OCR Engine Mode: 3 (Default) and Page Segmentation Mode: 6 (Assume a single uniform block of text)
    # results = pytesseract.image_to_data(gray_image, config=custom_config, output_type=pytesseract.Output.DICT)

    # # Iterate over each detected word
    # for i in range(len(results['text'])):
    #     # Extract word and bounding box information
    #     word = results['text'][i]
    #     x, y, w, h = results['left'][i], results['top'][i], results['width'][i], results['height'][i]

    #     # Check if the word is non-empty and small based on the threshold
    #     small_text_threshold = 370  # Adjust as needed
    #     if word and w < small_text_threshold and h < small_text_threshold:
    #         # Invert the color of the small text (assuming a white background)
    #         inverted_text = cv2.bitwise_not(image[y:y+h, x:x+w])

    #         # Replace the small text region with the inverted color
    #         image[y:y+h, x:x+w] = inverted_text

    # # Display the resulting image
    # resize_factor = 0.3
    # smaller_image = cv2.resize(image, None, fx=resize_factor, fy=resize_factor)
    # cv2.imshow('Image with small text removed', smaller_image)
    # cv2.waitKey(0)
    # cv2.destroyAllWindows()

def easyOCR():
    imagePath = "output.png"
    reader = easyocr.Reader(['en'])
    text = reader.readtext(imagePath)
    name_coordinates = []
    keys = [[1, "name"], [2, "serial"], [3, "home"], [4, "kin"], [5, "address"], [6, "born"], 
            [7, "at"], [8, "death"], [9, "place"], [10, "buried"], [11, "19"], [12, "in"], 
            [13,"city"], [14, "county"], [15, "division"], [16, "section"], [17, "lot"], 
            [18, "block"], [19, "grave no"], [20, "war"], [21, "service"], [22, "rank"],
            [23, "enlisted"], [24, "discharged"], [25, "information"], [26, "relationship"]]
    garb = []
    key_found = False
    for x in text:
        print(x)
    # Loop through the OCR results
    prev_bbox = None  
    current_word = "" 
    current_bbox = None
    for key in keys:
        key_found = False
        flag1 = False
        for bbox, word, score in text:
            if key[0] == 7:
                counter = 0
                for x in name_coordinates:
                    if x[0] == "BORN":
                        name_coordinates.append(["AT", [[name_coordinates[counter][1][1][0] + 734, name_coordinates[counter][1][0][1]], 
                            [name_coordinates[counter][1][1][0] + 827, name_coordinates[counter][1][0][1]],
                            [name_coordinates[counter][1][1][0] + 827, name_coordinates[counter][1][2][1]], 
                            [name_coordinates[counter][1][1][0] + 734, name_coordinates[counter][1][2][1]]]])                
                        break
                    counter += 1
                break
            if word.isupper() or word.lower() == "lot" or word.lower() == "19":
                # if (prev_bbox and bbox[0][0] - prev_bbox[1][0] <= 150 and prev_bbox and bbox[0][0] - prev_bbox[1][0] > 0):
                if (prev_bbox and 
                    0 <= (bbox[0][0] - prev_bbox[1][0]) <= 70 and 
                    (bbox[0][1] - prev_bbox[1][1]) <= 10):                    
                    current_word += " " + word
                    current_bbox[1] = bbox[1]
                    current_bbox[2] = bbox[2]
                else:
                    if current_word and "UNION" not in current_word:
                        if key[0] == 5 or key[0] == 12:
                            flag1 = True
                        if key[1].lower() == current_word.lower():
                            # print([current_word.upper(), current_bbox])
                            name_coordinates.append([current_word.upper(), current_bbox])
                            key_found = True
                            break
                        elif key[1].lower() in current_word.lower() and not flag1:
                            # print([current_word.upper(), current_bbox])
                            name_coordinates.append([current_word.upper(), current_bbox])
                            key_found = True
                            break 
                    current_word = word
                    current_bbox = bbox.copy()
                prev_bbox = bbox
        if not key_found:
            garb.append(key[1])
            name_coordinates.append([key[1].upper() + " not found", \
                                     [[2100, 3600], [2100, 3600], [2100, 3600], [2100, 3600]]])

    return garb, name_coordinates

def tesseractOCR(row_index, name_coordinates):
    imagePath = "output.png"
    image = Image.open(imagePath)
    count = 0
    counter = 1
    # nt = image.crop((419, 748, 1139, 871))
    # print(pytesseract.image_to_string(nt, lang='eng'))
    for x in name_coordinates: 
        if(count < len(name_coordinates)-1):
            # if (name_coordinates[count+1][1][0] > x[1][0]) and x != [[395, 1402], [662, 1402], [662, 1458], [395, 1458]] and count != 18:
            # print(x)
            if (name_coordinates[count+1][1][1][0] > x[1][1][0]) and "WAR" not in x[0]:
                cropped_image = image.crop((x[1][2][0] + 20, x[1][0][1]-70, 
                                            name_coordinates[count+1][1][0][0] - 20, x[1][2][1]))
            else:
                cropped_image = image.crop((x[1][2][0] + 20, x[1][0][1]-70, 3600, x[1][2][1]))
            # cropped_image = image.crop((219, 452, 413, 520))
            if x[0] == "NAME" or x[0] == "BORN" or x[0] == "BRANCH OF SERVICE" or x[0] == "WAR RECORD":
                text = pytesseract.image_to_string(cropped_image, lang='eng')
                filtered_text = text.replace("♀", "").replace("\n", "").replace("«", "") \
                                .replace("_", " ").replace("—. ", " ").replace("(", " ") \
                                .replace("CEMETERY", "").strip()
                # print(filtered_text)
                worksheet.cell(row=row_index, column=counter, value=filtered_text)
                counter += 1
            count += 1
        else:
            # print(x)
            cropped_image = image.crop((x[1][2][0] + 20, x[1][0][1]-70, 3600, x[1][2][1]))
            text = pytesseract.image_to_string(cropped_image, lang='eng')
            if x[0] == "NAME" or x[0] == "BORN" or x[0] == "BRANCH OF SERVICE" or x[0] == "WAR RECORD":
                text = pytesseract.image_to_string(cropped_image, lang='eng')
                filtered_text = text.replace("♀", "").replace("\n", "").replace("«", "") \
                                .replace("_", " ").replace("—. ", " ").replace("(", " ") \
                                .replace("CEMETERY", "").strip()
                worksheet.cell(row=row_index, column=counter, value=filtered_text)
                # print(filtered_text)
workbook = openpyxl.load_workbook('Veterans.xlsx')
worksheet = workbook['Sheet1']
current_directory = os.getcwd()
row_index = 2
png_files = [file for file in os.listdir(current_directory) if file.lower().endswith('.png')]
for x in png_files:
    if x != "output.png":
        preProcess(x)
        garb, name_coordinates = easyOCR()
        tesseractOCR(row_index, name_coordinates)
        row_index += 1
workbook.save('Veterans.xlsx')
# preProcess("4.png")
# easyOCR()
# imagePath = "output.png"
# image = Image.open(imagePath)
# nt = image.crop((192, 499, 378, 549))
# print(pytesseract.image_to_string(nt, lang='eng'))
# garb, name_coordinates = easyOCR()
# tesseractOCR(garb, name_coordinates)
