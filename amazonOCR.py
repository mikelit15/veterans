import boto3
import re
import openpyxl 
from openpyxl.styles import Font
from collections import defaultdict
import os
import PyPDF2
import fitz
import cv2
from reportlab.pdfgen import canvas
from PIL import Image
from nameparser import HumanName
import dateparser
from nameparser.config import CONSTANTS
from openpyxl.styles import PatternFill
from fuzzywuzzy import fuzz

def redact(file, flag, cemetery, letter):
    pdf_document  = fitz.open(file)
    first_page  = pdf_document.load_page(0)
    first_page_image  = first_page.get_pixmap(matrix=fitz.Matrix(600/72, 600/72))
    image_file  = "temp.png"
    first_page_image.save(image_file)
    img = cv2.imread(image_file)
    pt1 = (2670, 500)
    pt3 = (3526, 950)
    if flag:
        pt1 = (2670, 260)
        pt3 = (3500, 570)
    cv2.rectangle(img, pt1, pt3, (0, 0, 0), thickness=cv2.FILLED)
    cv2.imwrite(image_file, img)
    image = Image.open(image_file)
    pdfFile = "temp.pdf"
    c = canvas.Canvas(pdfFile, pagesize=(image.width, image.height))
    c.drawImage(image_file, 0, 0, width=image.width, height=image.height)
    c.save()

    fullLocation = r"\\ucclerk\pgmdoc\Veterans"
    redactedLocation = f'{cemetery} - Redacted'
    fullLocation = os.path.join(fullLocation, redactedLocation)
    fullLocation = os.path.join(fullLocation, letter)
    fileName = file.split(letter)
    redacted_pdf_document  = fitz.open(pdfFile)
    new_pdf_document  = fitz.open()
    new_pdf_document.insert_pdf(redacted_pdf_document, from_page=0, to_page=0)
    new_pdf_document.insert_pdf(pdf_document, from_page=1, to_page=1)
    new_pdf_file  = f'{fullLocation}\\{cemetery}{letter}{fileName[2].replace(".pdf", "")} redacted.pdf'
    new_pdf_document.save(new_pdf_file)
    new_pdf_document.close()
    redacted_pdf_document.close()
    pdf_document.close()
    return new_pdf_file    

def mergeImages(pathA, pathB, cemetery, letter):
    fullLocation = r"\\ucclerk\pgmdoc\Veterans"
    redactedLocation = f'{cemetery} - Redacted'
    fullLocation = os.path.join(fullLocation, redactedLocation)
    fullLocation = os.path.join(fullLocation, letter)
    fileName = pathA.split(letter)
    merger = PyPDF2.PdfMerger()
    merger.append(pathA)
    merger.append(pathB)
    string = pathA[:-5]
    string = string.split(letter) 
    string = string[-1].lstrip('0')
    # with open(f'\\ucclerk\pgmdoc\Veterans\\{cemetery}{letter}{fileName[2].replace("a.pdf", ".pdf")}', 'wb') as out_pdf:
    #     merger.write(out_pdf)
    merger2 = PyPDF2.PdfMerger()
    merger2.append(f'{fullLocation}\\{cemetery}{letter}{fileName[2].replace(".pdf", "")} redacted.pdf')
    merger2.append(f'{fullLocation}\\{cemetery}{letter}{fileName[2].replace("a.pdf", "b")} redacted.pdf')
    string = pathA[:-5]
    string = string.split(letter) 
    string = string[-1].lstrip('0')
    with open(f'{fullLocation}\\{cemetery}{letter}{fileName[2].replace("a.pdf", "")} redacted.pdf', 'wb') as out_pdf:
        merger2.write(out_pdf)

def get_kv_map(file_name, id, cemetery):
    with open(file_name, 'rb') as file:
        img_test = file.read()
        bytes_test = bytearray(img_test)
        print('Image loaded', id)
    client = boto3.client('textract', region_name='us-east-1')
    response = client.analyze_document(Document={'Bytes': bytes_test}, FeatureTypes=['FORMS'])
    blocks = response['Blocks']
    key_map = {}
    civil = ""
    world = ""
    cem = ""
    bYear = ""
    value_map = {}
    block_map = {}
    flag = False
    flag2 = False
    stopFlag = False
    for block in blocks:
        try:
            if block['Text'] == "Civil War":
                civil = "Civil War"
            if block['Text'] == 'World War':
                world = "World War"
            if block['Text'] == cemetery:
                cem = cemetery
                stopFlag = True
            if fuzz.partial_ratio(block['Text'].lower(), cemetery.lower()) > 80:
                cem = cemetery
                stopFlag = True
            if flag:
                if len(block['Text']) == 4:
                    bYear = block['Text'][:2]
                    if block['Text'][-2:].lower() == "in":
                        flag2 = True
                elif block['Text'].lower() == "in":
                    flag2 = True
                else:
                    bYear = block['Text']
                flag = False
            elif block['BlockType'] == "WORD" and block['Text'] == "19":
                flag = True
            elif block['BlockType'] == "WORD" and block['Text'][:2] == "19" \
                and len(block['Text']) > 2:
                if len(block['Text']) == 4:
                    if block['Text'][-2:].isnumeric():
                        bYear = block['Text'][-2:]
                elif len(block['Text']) == 3:
                    pass
                elif len(block['Text']) == 5:
                    temp = re.sub(r'[^0-9]', '', block['Text'])
                    if temp[-2:].isnumeric():
                        bYear = temp[-2:]
            elif flag2 and not stopFlag:
                cem = block['Text']
                flag2 = False
            elif block['BlockType'] == "WORD" and block['Text'].lower() == "in":
                flag2 = True
        except KeyError:
            pass
        block_id = block['Id']
        block_map[block_id] = block
        if block['BlockType'] == "KEY_VALUE_SET":
            if 'KEY' in block['EntityTypes']:
                key_map[block_id] = block
            else:
                value_map[block_id] = block
    return key_map, value_map, block_map, civil, cem, world, bYear

def get_kv_relationship(key_map, value_map, block_map):
    kvs = defaultdict(list)
    for block_id, key_block in key_map.items():
        value_block = find_value_block(key_block, value_map)
        key = get_text(key_block, block_map)
        key = re.sub(r'[^a-zA-Z0-9\s]', '', key)
        val = get_text(value_block, block_map)
        kvs[key].append(val)
    return kvs

def find_value_block(key_block, value_map):
    for relationship in key_block['Relationships']:
        if relationship['Type'] == 'VALUE':
            for value_id in relationship['Ids']:
                value_block = value_map[value_id]
    return value_block

def get_text(result, blocks_map):
    text = ''
    if 'Relationships' in result:
        for relationship in result['Relationships']:
            if relationship['Type'] == 'CHILD':
                for child_id in relationship['Ids']:
                    word = blocks_map[child_id]
                    if word['BlockType'] == 'WORD':
                        text += word['Text'] + ' '
                    if word['BlockType'] == 'SELECTION_ELEMENT':
                        if word['SelectionStatus'] == 'SELECTED':
                            text += 'X '
    return text

def print_kvs(kvs):
    for key, value in kvs.items():
        print(key, ":", value)
    
    print("\n")
    
def search_value(kvs, search_key):
    for key, value in kvs.items():
        key = str.rstrip(key)
        if key.upper() == search_key.upper():
            return value

def search_value_x(kvs, search_key):
    search_pattern = r'\b' + re.escape(search_key) + r'\b' 
    for key, value in kvs.items():
        if re.search(search_pattern, key, re.IGNORECASE):
            return value
        
def nameRule(finalVals, value):
    CONSTANTS.force_mixed_case_capitalization = True
    name = HumanName(value)
    name.capitalize()
    firstName = name.first
    middleName = name.middle
    lastName = name.last
    suffix = name.suffix
    suffi = ["Jr.", "Sr.", "I", "II", "III", "IV", "V"]
    temp = value.replace("Jr.", "").replace("Sr.", "").replace("I", "").replace("II", "")\
        .replace("III", "").replace("IV", "").replace("V", "")
    if ("," in value and "." in firstName):
        if len(middleName) > 0:
            middleName = name.first
            firstName = name.middle
        else:
            lastName = name.last
            firstName = name.first
    elif ("." in firstName):
        if len(middleName) > 0:
            lastName = name.first
            firstName = name.middle
            middleName = name.last
        else:
            lastName = name.first
            firstName = name.last
    elif ("." in lastName):
        if len(middleName) == 0:
            lastName = name.last
            firstName = name.first
        else:
            lastName = name.first
            middleName = name.last
            firstName = name.middle
    elif len(suffix) > 0 and not middleName:
        suffix = suffix.replace(", ", "")
        middleName = suffix.replace("Sr.", "").replace("Jr.", "").replace("I", "").replace("II", "")\
        .replace("III", "").replace("IV", "").replace("V", "")
        suffix = suffix.replace(middleName, "")
    elif len(suffix) > 0 and middleName:
        suffix = suffix.replace(", ", "")
        suffix = suffix.replace(middleName, "")
    if value in suffi and "." not in temp and "," not in temp:
        if len(middleName) > 0:
            firstName = name.middle
            middleName = name.last
            lastName = name.first
        else:
            firstName = name.last
            lastName = name.first
    elif "." not in temp and "," not in temp:
        if len(middleName) > 0:
            firstName = name.middle
            middleName = name.last
            lastName = name.first
        else:
            lastName = name.first
            firstName = name.last
    if len(middleName) > 2:
        middleName = middleName.replace(".", "")
    dots = 0
    for x in firstName:
        if (x == "."):
            dots+=1
    if (dots == 2):
        firstName = firstName.upper()
    else:
        firstName = firstName.replace(".", "")
    if suffix:
        if not "." in suffix:
            suffix += "."
    finalVals.append(lastName.replace(",", "").replace(".", ""))
    finalVals.append(firstName.replace(",", ""))
    finalVals.append(middleName.replace(",", "."))
    finalVals.append(suffix.replace(",", "."))

def parseBirth(birth, bYear, birthYYFlag):
    pattern = r'([A-Za-z]+),\s*(\d{4})'
    match = re.match(pattern, birth.strip())
    if birth[-1:] == " ":
        birth = birth[:-1]
    if match:
        bYear = match.group(2)
        birth = ""
    elif "/" in birth:
        year = birth.split('/')[-1]
        year = year.replace(" ", "")
        if len(year) == 4:
            birth = dateparser.parse(birth)
            bYear = birth.strftime("%Y")
            birth = birth.strftime("%m/%d/%Y")
        elif len(year) == 2:
            birthYYFlag = True
            birth = dateparser.parse(birth)
            bYear = birth.strftime("%Y")[-2:]
            birth = birth.strftime("%m/%d/%Y")[:-2]
        elif len(year) == 6 or len(year) == 5:
            bYear = year[-4:]
            birth = birth.replace(",", " ")
            birth = dateparser.parse(birth)
            birth = birth.strftime("%m/%d/%Y")
        elif len(year) > 6:
            temp = year[:4]
            if temp.isnumeric():
                death_no_spaces = birth.replace(' ', '')
                start_index = death_no_spaces.find(year[4:])
                num_spaces_before = birth[:start_index].count(' ')
                adjusted_start_index = start_index + num_spaces_before
                birth = birth[:adjusted_start_index].strip()
                birth = dateparser.parse(birth)
                birth = birth.strftime("%m/%d/%Y")
                bYear = temp
            else:
                birth = ""
        else:
            birth = ""
            bYear = ""
    elif "," in birth:
        year = birth.split(',')[-1]
        year = year.replace(" ", "")
        if len(year) == 4:
            birth = dateparser.parse(birth)
            bYear = birth.strftime("%Y")
            birth = birth.strftime("%m/%d/%Y")
        elif len(year) == 2:
            birthYYFlag = True
            birth = dateparser.parse(birth)
            bYear = birth.strftime("%Y")[-2:]
            birth = birth.strftime("%m/%d/%Y")[:-2]
        elif len(year) == 6 or len(year) == 5:
            bYear = year[-4:]
            birth = birth.replace(",", " ")
            birth = dateparser.parse(birth)
            birth = birth.strftime("%m/%d/%Y")
        elif len(year) > 6:
            temp = year[:4]
            if temp.isnumeric():
                death_no_spaces = birth.replace(' ', '')
                start_index = death_no_spaces.find(year[4:])
                num_spaces_before = birth[:start_index].count(' ')
                adjusted_start_index = start_index + num_spaces_before
                birth = birth[:adjusted_start_index].strip()
                birth = dateparser.parse(birth)
                birth = birth.strftime("%m/%d/%Y")
                bYear = temp
            else:
                birth = ""
        else:
            birth = ""
    elif " " in birth:
        year = birth.split(' ')[-1]
        year = year.replace(" ", "")
        if len(year) == 4:
            birth = dateparser.parse(birth)
            bYear = birth.strftime("%Y")
            birth = birth.strftime("%m/%d/%Y")
        elif len(year) == 2:
            birthYYFlag = True
            birth = dateparser.parse(birth)
            bYear = birth.strftime("%Y")[-2:]
            birth = birth.strftime("%m/%d/%Y")[:-2]
        elif len(year) == 6 or len(year) == 5:
            bYear = year[-4:]
            birth = birth.replace(",", " ")
            birth = dateparser.parse(birth)
            birth = birth.strftime("%m/%d/%Y")
        elif len(year) > 6:
            temp = year[:4]
            if temp.isnumeric():
                death_no_spaces = birth.replace(' ', '')
                start_index = death_no_spaces.find(year[4:])
                num_spaces_before = birth[:start_index].count(' ')
                adjusted_start_index = start_index + num_spaces_before
                birth = birth[:adjusted_start_index].strip()
                birth = dateparser.parse(birth)
                birth = birth.strftime("%m/%d/%Y")
                bYear = temp
            else:
                birth = ""
        else:
            birth = ""
    return birth, bYear, birthYYFlag

def parseDeath(death, dYear, deathYYFlag):
    pattern = r'([A-Za-z]+),\s*(\d{4})'
    match = re.match(pattern, death.strip())
    if death[-1:] == " ":
        death = death[:-1]
    if match:
        dYear = match.group(2)
        death = ""
    elif "/" in death:
        year3 = death.split('/')[-1]
        year3 = year3.replace(" ", "")
        if len(year3) == 4:
            death = dateparser.parse(death)
            dYear = death.strftime("%Y")
            death = death.strftime("%m/%d/%Y")
        elif len(year3) == 2:
            deathYYFlag = True
            death = dateparser.parse(death)
            dYear = death.strftime("%Y")[-2:]
            death = death.strftime("%m/%d/%Y")[:-2]
        elif len(year3) == 6 or len(year3) == 5:
            dYear = year3[-4:]
            death = death.replace(",", " ")
            death = dateparser.parse(death)
            death = death.strftime("%m/%d/%Y")
        elif len(year3) > 6:
            temp = year3[:4]
            if temp.isnumeric():
                death_no_spaces = death.replace(' ', '')
                start_index = death_no_spaces.find(year3[4:])
                num_spaces_before = death[:start_index].count(' ')
                adjusted_start_index = start_index + num_spaces_before
                death = death[:adjusted_start_index].strip()
                death = dateparser.parse(death)
                death = death.strftime("%m/%d/%Y")
                dYear = temp
            else:
                death = ""
        else:
            death = ""
    elif "," in death:
        year3 = death.split(',')[-1]
        year3 = year3.replace(" ", "")
        if len(year3) == 4:
            death = dateparser.parse(death)
            if death:
                dYear = death.strftime("%Y")
                death = death.strftime("%m/%d/%Y")
            else:
                dYear = year3
                death = ""
        elif len(year3) == 2:
            deathYYFlag = True     
            death = dateparser.parse(death)
            dYear = death.strftime("%Y")[-2:]
            death = death.strftime("%m/%d/%Y")[:-2]
        elif len(year3) == 6 or len(year3) == 5:
            dYear = year3[-4:]
            death = death.replace(",", " ")
            death = dateparser.parse(death)
            death = death.strftime("%m/%d/%Y")
        elif len(year3) > 6:
            temp = year3[:4]
            if temp.isnumeric():
                death_no_spaces = death.replace(' ', '')
                start_index = death_no_spaces.find(year3[4:])
                num_spaces_before = death[:start_index].count(' ')
                adjusted_start_index = start_index + num_spaces_before
                death = death[:adjusted_start_index].strip()
                death = dateparser.parse(death)
                death = death.strftime("%m/%d/%Y")
                dYear = temp
            else:
                death = ""
        else:
            death = ""
    elif " " in death:
        year3 = death.split(' ')[-1]
        year3 = year3.replace(" ", "")
        if len(year3) == 4:
            death = dateparser.parse(death)
            dYear = death.strftime("%Y")
            death = death.strftime("%m/%d/%Y")
        elif len(year3) == 2:
            deathYYFlag = True     
            death = dateparser.parse(death)
            dYear = death.strftime("%Y")[-2:]
            death = death.strftime("%m/%d/%Y")[:-2]
        elif len(year3) == 6 or len(year3) == 5:
            dYear = year3[-4:]
            death = death.replace(",", " ")
            death = dateparser.parse(death)
            death = death.strftime("%m/%d/%Y")
        elif len(year3) > 6:
            temp = year3[:4]
            if temp.isnumeric():
                death_no_spaces = death.replace(' ', '')
                start_index = death_no_spaces.find(year3[4:])
                num_spaces_before = death[:start_index].count(' ')
                adjusted_start_index = start_index + num_spaces_before
                death = death[:adjusted_start_index].strip()
                death = dateparser.parse(death)
                death = death.strftime("%m/%d/%Y")
                dYear = temp
            else:
                death = ""
        else:
            death = ""
    return death, dYear, deathYYFlag

def dateRule(finalVals, value, dob, buried, buriedYear, cent, war):
    warFlag = False
    warsFlag = False
    wars = ["World War 1", "World War 2", "Korean War", "Vietnam War", "Mexican Border War"]
    if war in wars:
        warsFlag = True
    if not buriedYear:
        buriedYear = buriedRule(buried, buriedYear, cent, warsFlag)
    birthYYFlag = False
    deathYYFlag = False
    bYear = ""
    dYear = ""
    birth = dob.replace(":", ".").replace("I", "1").replace(".", " ").replace("&", "").replace("x", "")
    if "At" in birth:
        birth = birth.split("At")
        birth = birth[0] 
    elif "AT" in birth:
        birth = birth.split("AT")
        birth = birth[0] 
    if birth[-1:] == " ":
        birth = birth[:-1]
    if birth[-1:] == ".":
        birth = birth[:-1]
    if birth[-1:] == "/":
        birth = birth[:-1]
    if "Age" in birth:
        match = re.search(r'\b\d{4}\b', birth)
        if match:
            bYear = match.group()
            birth = ""
        else:
            match = re.search(r'\b\d{2}\b', birth)
            if match:
                birth = ""
    death = value.replace(":", ".").replace("I", "1").replace(".", " ").replace("&", "").replace("x", "")
    if death[-1:] == " ":
        death = death[:-1]
    if death[-1:] == ".":
        death = death[:-1]
    if death[-1:] == "/":
        death = death[:-1]
    if "-" in birth:
        birth = birth.replace("-", "/")
    if "-" in death:
        death = death.replace("-", "/")
    if len(birth.replace(" ", "")) == 4:
        if len(death.replace(" ", "")) == 4:
            bYear = birth
            birth = ""
            dYear = death
            death = ""
        else:
            bYear = birth
            birth = ""
    elif len(death.replace(" ", "")) == 4:
        dYear = death
        death = ""
    if birth != "" and death != "":
        if dateparser.parse(birth, settings={'STRICT_PARSING': True}) != None:
            if dateparser.parse(death, settings={'STRICT_PARSING': True}) != None:
                birth, bYear, birthYYFlag = parseBirth(birth, bYear, birthYYFlag)
                death, dYear, deathYYFlag = parseDeath(death, dYear, deathYYFlag)
                if birthYYFlag and deathYYFlag:
                    if cent.isnumeric():
                        if bYear > dYear:
                            bYear = "18" + bYear
                            dYear = "19" + dYear
                        elif bYear < dYear:
                            bYear = "19" + bYear
                            dYear = "19" + dYear
                        birth = birth[:-2] + bYear
                        death = death[:-2] + dYear
                        finalVals.append(birth)
                        finalVals.append(int(bYear))
                        finalVals.append(death)
                        finalVals.append(int(dYear))
                    elif war:
                        if war in wars and dYear[0] == "0":
                            bYear = "19" + bYear
                            birth = birth[:-2] + bYear
                            dYear = "20" + dYear
                            death = death[:-2] + dYear
                            finalVals.append(birth)
                            finalVals.append(int(bYear))
                            finalVals.append(death)
                            finalVals.append(int(dYear))
                        elif war in wars and bYear > dYear:
                            bYear = "18" + bYear
                            birth = birth[:-2] + bYear
                            dYear = "19" + dYear
                            death = death[:-2] + dYear
                            finalVals.append(birth)
                            finalVals.append(int(bYear))
                            finalVals.append(death)
                            finalVals.append(int(dYear))
                        elif war in wars and bYear < dYear:
                            bYear = "19" + bYear
                            birth = birth[:-2] + bYear
                            dYear = "19" + dYear
                            death = death[:-2] + dYear
                            finalVals.append(birth)
                            finalVals.append(int(bYear))
                            finalVals.append(death)
                            finalVals.append(int(dYear))
                    elif buriedYear:
                        if buriedYear[:2] == "18" and dYear > bYear:
                            bYear = "18" + bYear
                            dYear = "18" + dYear
                        elif buriedYear[:2] == "19" and dYear < bYear:
                            bYear = "18" + bYear
                            dYear = "19" + dYear
                        elif buriedYear[:2] == "19" and dYear > bYear:
                            bYear = "19" + bYear
                            dYear = "19" + dYear
                        elif bYear < dYear:
                            bYear = "19" + bYear
                            dYear = "19" + dYear
                            warFlag = True
                        elif bYear > dYear:
                            bYear = "18" + bYear
                            dYear = "19" + dYear
                            warFlag = True
                        birth = birth[:-2] + bYear
                        death = death[:-2] + dYear
                        finalVals.append(birth)
                        if bYear:
                            finalVals.append(int(bYear))
                        else:
                            finalVals.append("")
                        finalVals.append(death)
                        if dYear:
                            finalVals.append(int(dYear))
                        else:
                            finalVals.append("")
                    elif dYear < bYear:
                        bYear = "18" + bYear
                        birth = birth[:-2] + bYear
                        dYear = "19" + dYear
                        death = death[:-2] + dYear
                        finalVals.append(birth)
                        finalVals.append(int(bYear))
                        finalVals.append(death)
                        finalVals.append(int(dYear))
                        warFlag = True
                    elif dYear > bYear:
                        bYear = "19" + bYear
                        birth = birth[:-2] + bYear
                        dYear = "19" + dYear
                        death = death[:-2] + dYear
                        finalVals.append(birth)
                        finalVals.append(int(bYear))
                        finalVals.append(death)
                        finalVals.append(int(dYear))
                        warFlag = True
                    else:
                        finalVals.append("")
                        finalVals.append("") 
                        finalVals.append("")
                        finalVals.append("")
                    return warFlag
                elif not birthYYFlag and deathYYFlag:
                    if bYear[:2] == "19" and bYear[-2:] < dYear:
                        dYear = "19" + dYear
                    elif bYear[:2] == "19" and bYear[-2:] > dYear:
                        dYear = "20" + dYear
                    elif bYear[:2] == "18" and bYear[-2:] < dYear:
                        dYear = "18" + dYear
                    elif bYear[:2] == "18" and bYear[-2:] > dYear: 
                        dYear = "19" + dYear
                    death = death[:-2] + dYear
                    if birth:
                        finalVals.append("")
                    else:
                        finalVals.append("")
                    if bYear:
                        finalVals.append("")
                    else:
                        finalVals.append("")
                    if len(dYear) == 2:
                        dYear = "19" + dYear
                        finalVals.append(death[:-2] + dYear)
                        finalVals.append(int(dYear))
                    elif len(dYear) == 4:
                        finalVals.append(death)
                        finalVals.append(int(dYear))
                    else:
                        finalVals.append("")
                    return warFlag
                elif birthYYFlag and not deathYYFlag:
                    if dYear[:2] == "19" and dYear[-2:] < bYear:
                        bYear = "18" + bYear
                    elif dYear[:2] == "19" and dYear[-2:] > bYear:
                        bYear = "19" + bYear
                    elif dYear[:2] == "18" and dYear[-2:] < bYear:
                        bYear = "19" + bYear
                    elif dYear[:2] == "18" and dYear[-2:] > bYear: 
                        bYear = "18" + bYear
                    elif dYear[:2] == "20":
                        bYear = "19" + bYear
                    birth = birth[:-2] + bYear
                    finalVals.append(birth)
                    finalVals.append(int(bYear))
                    finalVals.append(death)
                    finalVals.append(int(dYear))
                    return warFlag
                else:
                    finalVals.append(birth)
                    finalVals.append(int(bYear))
                    finalVals.append(death)
                    finalVals.append(int(dYear))
                    return warFlag
            else:
                birth = dateparser.parse(birth)
                bYear = birth.strftime("%Y")
                birth = birth.strftime("%m/%d/%Y")
                death, dYear, deathYYFlag = parseDeath(death, dYear, deathYYFlag)
                if not deathYYFlag:
                    finalVals.append(birth)
                    if bYear:
                        finalVals.append(int(bYear))
                    else:
                        finalVals.append("")
                    finalVals.append(death)
                    if dYear:
                        finalVals.append(int(dYear))
                    else:
                        finalVals.append("")
                    return warFlag
                else:
                    if bYear[:2] == "19" and bYear[-2:] < dYear:
                        dYear = "20" + dYear
                    elif bYear[:2] == "19" and bYear[-2:] > dYear:
                        dYear = "19" + dYear
                    elif bYear[:2] == "18" and bYear[-2:] < dYear:
                        dYear = "19" + dYear
                    elif bYear[:2] == "18" and bYear[-2:] > dYear: 
                        dYear = "18" + dYear
                    death = death[:-2] + dYear
                    finalVals.append(birth)
                    if bYear:
                        finalVals.append(int(bYear))
                    else:
                        finalVals.append("")
                    finalVals.append(death)
                    if dYear:
                        finalVals.append(int(dYear))
                    else:
                        finalVals.append("")
                    return warFlag
        elif dateparser.parse(death, settings={'STRICT_PARSING': True}) != None:
            finalVals.append("")
            finalVals.append("")
            death = dateparser.parse(death)
            finalVals.append(death.strftime("%m/%d/%Y"))
            finalVals.append(int(death.strftime("%Y")))
            return warFlag
        else:
            birth, bYear, birthYYFlag = parseBirth(birth, bYear, birthYYFlag)
            death, dYear, deathYYFlag = parseDeath(death, dYear, deathYYFlag)
            if birthYYFlag and deathYYFlag:
                if cent.isnumeric() or buriedYear.isnumeric():
                    if bYear > dYear:
                        bYear = "18" + bYear
                        dYear = "19" + dYear
                    elif bYear < dYear:
                        bYear = "19" + bYear
                        dYear = "19" + dYear
                    birth = birth[:-2] + bYear
                    death = death[:-2] + dYear
                    finalVals.append(birth)
                    finalVals.append(int(bYear))
                    finalVals.append(death)
                    finalVals.append(int(dYear))
                    return warFlag
                elif dYear < bYear:
                    finalVals.append("")
                    finalVals.append(int("18" + bYear))
                    finalVals.append("")
                    finalVals.append(int("19" + dYear))
                    return warFlag
            elif birthYYFlag and not deathYYFlag:
                if dYear[:2] == "19" and dYear[-2:] < bYear:
                    bYear = "18" + bYear
                elif dYear[:2] == "19" and dYear[-2:] > bYear:
                    bYear = "19" + bYear
                elif dYear[:2] == "18" and dYear[-2:] < bYear:
                    bYear = "19" + bYear
                elif dYear[:2] == "18" and dYear[-2:] > bYear: 
                    bYear = "18" + bYear
                elif dYear[:2] == "20":
                    bYear = "19" + bYear
                birth = birth[:-2] + bYear
                finalVals.append(birth)
                if bYear: 
                    finalVals.append(int(bYear))
                else:
                    finalVals.append("")
                finalVals.append(death)
                finalVals.append(int(dYear))
                return warFlag
            elif not birthYYFlag and deathYYFlag:
                if bYear[:2] == "19" and bYear[-2:] < dYear:
                    dYear = "20" + dYear
                elif bYear[:2] == "19" and bYear[-2:] > dYear:
                    dYear = "19" + dYear
                elif bYear[:2] == "18" and bYear[-2:] < dYear:
                    dYear = "19" + dYear
                elif bYear[:2] == "18" and bYear[-2:] > dYear: 
                    dYear = "18" + dYear
                death = death[:-2] + dYear
                finalVals.append(birth)
                finalVals.append(int(bYear))
                finalVals.append(death)
                finalVals.append(int(dYear))
                return warFlag
    elif birth == "" and death != "":
        death, dYear, deathYYFlag = parseDeath(death, dYear, deathYYFlag)
        if not deathYYFlag:
            if bYear:
                finalVals.append("")
                finalVals.append(int(bYear))
            else:
                finalVals.append("")
                finalVals.append("")
            finalVals.append(death)
            if dYear:
                finalVals.append(int(dYear))
            else:
                finalVals.append("")
        else:
            if bYear:
                if bYear[:2] == "18" and bYear[-2:] < dYear:
                    dYear = "18" + dYear
                elif bYear[:2] == "18" and bYear[-2:] > dYear:
                    dYear = "19" + dYear
                elif bYear[:2] == "19" and bYear[-2:] < dYear:
                    dYear = "19" + dYear
                elif bYear[:2] == "19" and bYear[-2:] > dYear:
                    dYear = "20" + dYear
                finalVals.append("")
                finalVals.append(int(bYear))
                death = death[:-2] + dYear
                finalVals.append(death)
                finalVals.append(int(dYear))
            elif cent.isnumeric() or buriedYear.isnumeric():
                finalVals.append("")
                finalVals.append("") 
                dYear = "19" + dYear
                finalVals.append(death[:-2] + dYear)
                finalVals.append(int(dYear))
            elif war in wars:
                finalVals.append("")
                finalVals.append("") 
                dYear = "19" + dYear
                finalVals.append(death[:-2] + dYear)
                finalVals.append(int(dYear))
            else:
                finalVals.append("")
                finalVals.append("") 
                finalVals.append("")
                finalVals.append("") 
    elif birth != "" and death == "":
        birth, bYear, birthYYFlag = parseBirth(birth, bYear, birthYYFlag)
        if not birthYYFlag:
            finalVals.append(birth)
            finalVals.append(int(bYear))
            if dYear != "":
                finalVals.append("")
                finalVals.append(int(dYear))
                return warFlag
            elif buriedYear != "":
                if buriedYear < bYear[:-2]:
                    dYear = "19" + buriedYear
                else:
                    dYear = "18" + buriedYear
                finalVals.append("")
                finalVals.append(int(dYear))
                return warFlag
            else:
                finalVals.append("")
                finalVals.append("") 
        else:
            if dYear:
                if dYear[:2] == "18" and dYear[-2:] > bYear:
                    bYear = "18" + bYear
                elif dYear[:2] == "19" and dYear[-2:] < bYear:
                    bYear = "18" + bYear
                elif dYear[:2] == "19" and dYear[-2:] > bYear:
                    bYear = "19" + bYear
                elif dYear[:2] == "20" and dYear[-2:] < bYear:
                    bYear = "19" + bYear
                birth = birth[:-2] + bYear
                finalVals.append(birth)
                finalVals.append(int(bYear))
                finalVals.append("")
                finalVals.append(int(dYear))
                return warFlag
            elif cent.isnumeric():
                if bYear > cent:
                    bYear = "18" + bYear
                else:
                    bYear = "19" + bYear
                birth = birth[:-2] + bYear
                finalVals.append(birth)
                finalVals.append(int(bYear))
            else:
                finalVals.append("")
                finalVals.append("") 
                finalVals.append("")
                finalVals.append("") 
    elif birth == "" and death == "":
        if buriedYear.isnumeric():
            if len(buriedYear) == 2:
                dYear = "19" + buriedYear
            elif len(buriedYear) == 4:
                dYear = buriedYear
        elif cent.isnumeric():
            dYear = "19" + cent
        if bYear != "" and dYear != "":
            finalVals.append("")
            finalVals.append(int(bYear))
            finalVals.append("")
            finalVals.append(int(dYear))
        elif bYear != "" and dYear == "":
            finalVals.append("")
            finalVals.append(int(bYear))
            finalVals.append("")
            finalVals.append("")
        elif bYear == "" and dYear != "":
            finalVals.append("")
            finalVals.append("") 
            finalVals.append("")
            finalVals.append(int(dYear))
        else:
            finalVals.append("")
            finalVals.append("") 
            finalVals.append("")
            finalVals.append("") 
    return warFlag

def buriedRule(value, bYear, cent, warsFlag):
    buried = value
    buried = buried.replace("I", "1").replace("-", "/")
    buriedYear = bYear
    if cent:
        buriedYear = "19" + cent
    elif buried or buriedYear:
        if dateparser.parse(buried, settings={'STRICT_PARSING': True}) == None:
            if buriedYear:
                if len(buriedYear) == 2:
                    buriedYear = buriedYear
                elif len(buriedYear) == 4:
                    buriedYear = buriedYear[2:]
                else:
                    buriedYear = ""
                if "," in buried:
                    temp = buried + " 19" + buriedYear
                    if dateparser.parse(temp, settings={'STRICT_PARSING': True}) != None:
                        buried = dateparser.parse(temp)
                        buried = buried.strftime("%m/%d/%Y")
                        buriedYear = "19" + buriedYear
                else:
                    temp = buried + ", 19" + buriedYear
                    if dateparser.parse(temp, settings={'STRICT_PARSING': True}) != None:
                        buried = dateparser.parse(temp)
                        buried = buried.strftime("%m/%d/%Y")
                        if len(bYear) == 2:
                            buriedYear = "19" + bYear
                        else:
                            buriedYear = bYear
                    else:
                        buried = ""
                        if len(bYear) == 2:
                            buriedYear = "19" + bYear
                        else:
                            buriedYear = bYear
            else:
                try:
                    before_comma = buried.split(',', 1)[0].strip()
                    after_comma = buried.split(',', 1)[1].strip()
                    if len(after_comma) > 4:
                        buriedYear = buriedRule(before_comma, after_comma[:4], cent, warsFlag) 
                except IndexError:
                    print("buried is weird")
                    buried = ""
        else:
            year5 = buried.split('/')[-1]
            if year5[-1:] == " ":
                year5 = year5[:-1]
            if len(year5) == 4:
                buriedYear = year5
            elif len(year5) == 2:
                if warsFlag:
                    buriedYear = "19" + year5
                else:
                    buriedYear = ""
    return buriedYear

def warRule(value, civil, world):
    war = value
    if war:
        if war[0] == " ":
            war = war[1:]
    ww1years = ["1914", "1915", "1916", "1917", "1918"]
    if "N." in war or "N" in war:
        war = war.replace("N", "W")
    if "T" in war:
        if "1" in war:
            war = war.replace("T", "1")
        elif "I" in war:
            war = war.replace("T", "I")
        elif "L" in war:
            war = war.replace("T", "L")
        elif "l" in war:
            war = war.replace("T", "l")
    war = war.replace(" ", "")
    tempText = war.replace("-", "").replace(".", "").replace("#", "").replace(" ", "").replace(",", "")
    ww1_and_2_pattern = re.compile(r'(WW1|WWI|W\.?W\.?\s?(1|I|One|ONE)|World\s*War\s*(1|I|One|ONE))\s*(and|&)?\s*'\
    r'(WW2|WWII|W\.?W\.?\s?(2|II|Two|TWO)|World\s*War\s*(2|II|Two|TWO))', re.IGNORECASE)    
    # ww1_pattern = re.compile(r'WWI|WW1|WW\s?1|World\s*War\s*(1|I|ONE)|WorldWar1', re.IGNORECASE)
    ww1_pattern = re.compile(r'WW1|WWI|WWl|\b1\b|W\.?W\.?\s?(1|I|l|One|ONE)?|World\s*War\s*(1|I|l|One|ONE)?|WorldWar\s*(1|I|l|One|ONE)?', re.IGNORECASE)    
    # ww2_pattern = re.compile(r'WWII|WW2|WW11|WWLL|WWll|WW\s?2|World\s*War\s*(II|2|Two|LL|ll)|\b2\b', re.IGNORECASE)
    ww2_pattern = re.compile(r'WW2|WWII|WWll|WW11|\b2\b|W\.?W\.?\s?(2|II|ll|Two|TWO)|World\s*War\s*(2|II|ll|Two|TWO)|WorldWar\s*(2|II|ll|Two|TWO)', re.IGNORECASE)    
    if ww1_and_2_pattern.findall(tempText.upper()):
        war = "World War 1 and World War 2"
    elif (ww2_pattern.findall(tempText.upper())):
        war = "World War 2"
    elif (ww1_pattern.findall(tempText.upper())):
        war = "World War 1"
    # if (ww1_and_2_pattern.findall(tempText)):
    #     war = "World War 1 and 2"
    else:
        war = war.replace(".", "").replace("Calv", "").replace("Vols", "")
        if "Korea" in war:
            war = "Korean War"
        elif "Vietnam" in war:
            war = "Vietnam War"
        elif "Civil" in war or "Citil" in war or "Gettysburg" in war:
            war = "Civil War"
        elif "Spanish" in war or "Amer" in war or "American" in war:
            war = "Spanish American War"
        elif "Mexican" in war:
            war = "Mexican Border War"
        elif "Rebellion" in war:
            war = "War of the Rebellion"
        else:
            war = ""
        words = war.split()
        for x in words:
            if x in ww1years:
                war = "World War 1"
                break
    if civil != "" and war != "Civil War":
        war = "Civil War"
    if world != "" and war != "World War 1":
        war = "World War 1"
    if "Army" in war or "US" in war:
        war = ""
    return war

def branchRule(finalVals, value, war):
    armys = ["co", "army", "inf", "infantry", "infan", "usa", "med", \
            "cav", "div", "sig", "art", "corps", "corp"]
    navys = ["hospital", "navy", "naval"]
    guards = ["113", "102d", "114", "44", "181", "250", "112"]
    branch = value
    branch = branch.replace("/", " ").replace(".", " ").replace("th", "").replace("-", "")\
        .replace(", ", " ")
    words = branch.split()
    if war in value and war != "":
        branch = ""
        value = ""
    for word in words:
        word = word.lower()
        if "air force" in branch.lower():
            branch = "Air Force"
            break
        elif "marine" in branch.lower():
            branch = "Marine Corps"
            break
        elif "air" in word:
            branch = "Air Force"
            break
        elif word in armys:
            branch = "Army"
            break
        elif word in navys:
            branch = "Navy"
            break
        elif "coast" in word:
            branch = "Coast Guard"
            break
        elif word in guards:
            branch = "National Guard"
            break
        else:
            branch = ""
    finalVals.append(value)   
    finalVals.append(branch)   

def createRecord(file_name, id, cemetery):
    civil = ""
    world = ""
    war = ""
    bYear = ""
    cent = ""
    buried = ""
    warFlag = False
    cem = []
    finalVals = []
    pageReader = PyPDF2.PdfReader(open(file_name, 'rb'))
    page = pageReader.pages[0]
    pdf_writer = PyPDF2.PdfWriter()
    pdf_writer.add_page(page)
    with open("temp.pdf", 'wb') as output_pdf:
        pdf_writer.write(output_pdf)
    key_map, value_map, block_map, civil, cem, world, bYear = get_kv_map("temp.pdf", id, cemetery)
    kvs = get_kv_relationship(key_map, value_map, block_map)
    print_kvs(kvs)
    keys = ["", "NAME", "WAR RECORD", "BORN", "19", "BURIED", "DATE OF DEATH", "WAR RECORD", "BRANCH OF SERVICE" , "IN"]
    flag = False
    flag3 = False
    if search_value_x(kvs, "Care Assigned") == None:
        flag = False
    else:
        flag = True
    dob = ""
    for x in keys:
        try:
            value = search_value_x(kvs, x)[0]
        except TypeError:
            value = "" 
        if x == "NAME":
            nameRule(finalVals, value)
        elif x == "BORN":
            dob = value
        elif x == "DATE OF DEATH":
            warFlag = dateRule(finalVals, value, dob, buried, bYear, cent, war)
        elif x == "WAR RECORD" and flag3:
            finalVals.append(value)
            finalVals.append(war)
            flag3 = False
        elif x == "WAR RECORD":
            if "&" in value:
                wars = value.split("&")
                war1 = warRule(wars[0], civil, world)
                war2 = warRule(wars[1], civil, world)
                war = war1 + " and " + war2
            elif "and" in value:
                wars = value.split("and")
                war1 = warRule(wars[0], civil, world)
                war2 = warRule(wars[1], civil, world)
                war = war1 + " and " + war2
            else:
                war = warRule(value, civil, world)
            flag3 = True
        elif x == "BRANCH OF SERVICE":
            branchRule(finalVals, value, war)
        elif x == "IN":
            try:
                cem = cem.replace(",", "").replace(".", "").replace(":", "")
                cem = cem.lower()
                first = cem[0].upper()
                cem = first + cem[1:]
                finalVals.append(cem)
            except IndexError:
                finalVals.append("")
        elif x == "19":
            if value:
                try:
                    tempCent = value.replace(",", "").replace(".", "").replace(":", "").replace("/", "").replace(" ", "")
                    for y in tempCent:
                        if y.isnumeric():
                            cent += y
                except IndexError:
                    pass
            else:
                pass
        elif x == "BURIED":
            buried = value
            if "in" in buried:
                buried = buried.replace("?", "")
                buried = buried.split("in")
                try:
                    cem = buried[1].replace(" ", "").replace(",", "").replace(".", "").replace(":", "")
                except IndexError:
                    pass
                buried = buried[0]
    return finalVals, flag, warFlag

def tempRecord(file_name, val, id, cemetery):
    print("Performing Temp", id, val.upper())
    buried = ""
    civil = ""
    world = ""
    war = ""
    bYear = ""
    cent = ""
    warFlag = False
    cem = []
    finalVals = []
    pageReader = PyPDF2.PdfReader(open(file_name, 'rb'))
    page = pageReader.pages[0]
    pdf_writer = PyPDF2.PdfWriter()
    pdf_writer.add_page(page)
    with open("temp.pdf", 'wb') as output_pdf:
        pdf_writer.write(output_pdf)
    key_map, value_map, block_map, civil, cem, world, bYear = get_kv_map("temp.pdf", id, cemetery)
    kvs = get_kv_relationship(key_map, value_map, block_map)
    print_kvs(kvs)
    keys = ["", "NAME", "WAR RECORD", "BORN", "19", "DATE OF DEATH", "WAR RECORD", "BRANCH OF SERVICE" , "IN"]
    flag = False
    flag3 = False
    if search_value_x(kvs, "Care Assigned") == None:
        flag = False
    else:
        flag = True
    dob = ""
    for x in keys:
        try:
            value = search_value(kvs, x)[0]
        except TypeError:
            value = "" 
        if x == "NAME":
            nameRule(finalVals, value)
        elif x == "":
            value = value.replace(" ", "")
            buried = value[-2:]
        elif x == "BORN":
            dob = value
        elif x == "DATE OF DEATH":
            warFlag = dateRule(finalVals, value, dob, buried, bYear, cent, war)
        elif x == "WAR RECORD" and flag3:
            finalVals.append(value)
            finalVals.append(war)
            flag3 = False
        elif x == "WAR RECORD":
            if "&" in value:
                wars = value.split("&")
                war1 = warRule(wars[0], civil, world)
                war2 = warRule(wars[1], civil, world)
                war = war1 + " and " + war2
            elif "and" in value:
                wars = value.split("and")
                war1 = warRule(wars[0], civil, world)
                war2 = warRule(wars[1], civil, world)
                war = war1 + " and " + war2
            else:
                war = warRule(value, civil, world)
            flag3 = True
        elif x == "BRANCH OF SERVICE":
            branchRule(finalVals, value, war)
        elif x == "IN":
            try:
                cem = cem.replace(",", "").replace(".", "").replace(":", "")
                cem = cem.lower()
                first = cem[0].upper()
                cem = first + cem[1:]
                finalVals.append(cem)
            except IndexError:
                finalVals.append("")
        elif x == "19":
            if value:
                try:
                    cent = value.replace(",", "").replace(".", "").replace(":", "").replace("/", "").replace(" ", "")
                except IndexError:
                    pass
            else:
                pass
        elif x == "BURIED":
            buried = value
            if "in" in buried:
                buried = buried.replace("?", "")
                buried = buried.split("in")
                try:
                    cem = buried[1].replace(" ", "").replace(",", "").replace(".", "").replace(":", "")
                except IndexError:
                    pass
                buried = buried[0]
    return finalVals, flag, warFlag

def mergeRecords(vals1, vals2, rowIndex, id, warFlag):
    counter = 1
    def length(item):
        return len(str(item))
    merged_array = [max(a, b, key=length) for a, b in zip(vals1, vals2)] 
    worksheet.cell(row=rowIndex, column=counter, value=id)
    counter += 1
    for x in merged_array:
        worksheet.cell(row=rowIndex, column=counter, value=x)
        counter += 1
    # Orange if merged record and no issues
    highlight_color = PatternFill(start_color="E3963E", end_color="E3963E", fill_type="solid")
    for colIndex in range(2, 15):
        cell = worksheet.cell(row=rowIndex, column=colIndex)
        cell.fill = highlight_color
        cell = worksheet.cell(row=rowIndex, column=16)
        cell.fill = highlight_color
    # Teal if merged, and record adjusted using comparison
    if warFlag:
        highlight_color = PatternFill(start_color="00FFB9", end_color="00FFB9", fill_type="solid")
        for colIndex in range(2, 15):
            cell = worksheet.cell(row=rowIndex, column=colIndex)
            cell.fill = highlight_color
            cell = worksheet.cell(row=rowIndex, column=16)
            cell.fill = highlight_color
    # Blue if merged, and record has no DOD
    if (worksheet[f'{"I"}{rowIndex}'].value == ""):
        highlight_color = PatternFill(start_color="00C6FF", end_color="00C6FF", fill_type="solid")
        for colIndex in range(2, 15):
            cell = worksheet.cell(row=rowIndex, column=colIndex)
            cell.fill = highlight_color
            cell = worksheet.cell(row=rowIndex, column=16)
            cell.fill = highlight_color
    # Pink if merged, and cemetery does not match
    if (worksheet[f'{"N"}{rowIndex}'].value) != cemetery:
        highlight_color = PatternFill(start_color="FC80AC", end_color="FC80AC", fill_type="solid")
        for colIndex in range(2, 15):
            cell = worksheet.cell(row=rowIndex, column=colIndex)
            cell.fill = highlight_color
            cell = worksheet.cell(row=rowIndex, column=16)
            cell.fill = highlight_color
    # Red is merged record and last name letter does not match
    if worksheet[f'{"B"}{rowIndex}'].value[0] != letter:
        highlight_color = PatternFill(start_color="D2042D", end_color="D2042D", fill_type="solid")
        for colIndex in range(2, 15):
            cell = worksheet.cell(row=rowIndex, column=colIndex)
            cell.fill = highlight_color
            cell = worksheet.cell(row=rowIndex, column=16)
            cell.fill = highlight_color
    # Dark Green if merged, and first and last name are the same, bug
    if (worksheet[f'{"B"}{rowIndex}'].value) == (worksheet[f'{"C"}{rowIndex}'].value):
        highlight_color = PatternFill(start_color="0B8C36", end_color="0B8C36", fill_type="solid")
        for colIndex in range(2, 15):
            cell = worksheet.cell(row=rowIndex, column=colIndex)
            cell.fill = highlight_color
            cell = worksheet.cell(row=rowIndex, column=16)
            cell.fill = highlight_color
            
def clean(cemetery, letter, name_path, cem_path, counterA, is_first_file):
    pdf_files = sorted(os.listdir(name_path))
    letters = sorted([folder for folder in os.listdir(cem_path) if os.path.isdir(os.path.join(cem_path, folder))])
    try:
        current_letter_index = letters.index(letter)
    except ValueError:
        return
    folder_before_index = max(0, current_letter_index - 1)
    folder_before = letters[folder_before_index]
    pdf_files_before = sorted([file for file in os.listdir(os.path.join(cem_path, folder_before)) if file.lower().endswith('.pdf')])
    max_counter = 0
    for file in pdf_files_before:
        match = re.search(r'\d+', file)
        if match:
            number = int(match.group())
            max_counter = max(max_counter, number)
    counter = max_counter + 1
    for x in pdf_files:
        if is_first_file:
            counter = counterA 
        if "a.pdf" in x:
            newName = f"{cemetery}{letter}{counter:05d}a.pdf"
            os.rename(os.path.join(name_path, x), os.path.join(name_path, newName))
            counter -= 1
        elif "b.pdf"  in x:
            newName = f"{cemetery}{letter}{counter:05d}b.pdf"
            os.rename(os.path.join(name_path, x), os.path.join(name_path, newName))
        elif "a redacted" in x:
            newName = f"{cemetery}{letter}{counter:05d}a redacted.pdf"
            os.rename(os.path.join(name_path, x), os.path.join(name_path, newName))
            counter -= 1
        elif "b redacted" in x:
            newName = f"{cemetery}{letter}{counter:05d}b redacted.pdf"
            os.rename(os.path.join(name_path, x), os.path.join(name_path, newName))
            counter -= 1
        elif "redacted" in x:
            newName = f"{cemetery}{letter}{counter:05d} redacted.pdf"
            os.rename(os.path.join(name_path, x), os.path.join(name_path, newName))
            counter -= 1
        else:    
            newName = f"{cemetery}{letter}{counter:05d}.pdf"
            os.rename(os.path.join(name_path, x), os.path.join(name_path, newName))
        counter += 1
        is_first_file = False

def find_next_empty_row(worksheet):
    for row in worksheet.iter_rows(min_row=2, min_col=1, max_col=1):
        if row[0].value is None:
            return row[0].row
    return worksheet.max_row + 1 

def main():
    cemeterys = ['Beth David', 'Beth Isreal', 'B\'Nai Abraham', 'B\'Nai Isreal', \
        'B\'Nai Jeshurum', 'Elizabeth Jewish', 'Evergreen', 'EvergreenP', 'Extra', \
        'Fairview', 'Gomel', 'Graceland', 'Hazelwood', 'Hebrew CemeteryN', \
        'Hebrew CemeterySP', 'Hillside', 'Historian', 'Hollywood', 'Hollywood Memorial', \
        'Isreal Verein', 'Jewish Educational', 'Medal of Honor', 'Misc', 'Misc OOS', \
        'Mt. Calvary', 'Mt. Lebanon', 'Mt. Moriah', 'Mt. Olivet', 'Oheb Shalom', \
        'Oheb ShalomL', 'Rahway', 'Rheimahvvim', 'Rosedale', 'Rosehill', \
        'Rosehill Crematory', 'St. Gertrude', 'St. Mary', 'St. MaryP', 'St. Teresa', \
        'Torah']
    miscs = ['Alpine', 'Arlington', 'Atlantic County', 'Atlantic View', 'BaptistPL', \
        'BaptistSP', 'Bay ViewJC', 'Bay ViewL', 'Belvidere', 'Bloomfield', 'Bloomsbury', \
        'Brainard', 'Brick Dock', 'Bronze Memorials', 'Calvary', 'Cedar HillEM', \
        'CedarHillM', 'Cedar Lawn', 'Christ Church', 'Clinton', 'Clover']
    global cemSet
    cemSet = set(cemeterys)
    network_folder = r"\\ucclerk\pgmdoc\Veterans\Cemetery"
    os.chdir(network_folder)
    # workbook = openpyxl.load_workbook('Veterans.xlsx')
    global cemetery
    cemetery = "Evergreen"
    cem_path = os.path.join(network_folder, cemetery)
    if cemetery == "Misc":
        miscCem = "Arlington"
        cem_path = os.path.join(cem_path, miscCem)
    global letter
    letter = "A"
    name_path = letter
    name_path = os.path.join(cem_path, name_path)
    initialCount = 42772
    pathA = ""
    rowIndex = 2
    global worksheet
    # worksheet = workbook[cemetery]
    warFlag = False
    
    uppercase_alphabet = [chr(i) for i in range(ord('A'), ord('Z') + 1)]
    firstFileFlag = True
    for letter in uppercase_alphabet:
        try:
            name_path = os.path.join(cem_path, letter)
            print(letter)
            clean(cemetery, letter, name_path, cem_path, initialCount, firstFileFlag)
            firstFileFlag = False
        except FileNotFoundError:
            continue
    
    # pdf_files = sorted(os.listdir(name_path))
    # initialID = 1
    # for y in range(len(pdf_files)):
    #     warFlag = False
    #     file_path = os.path.join(name_path, pdf_files[y])
    #     rowIndex = find_next_empty_row(worksheet)
    #     try:
    #         id = worksheet[f'{"A"}{rowIndex-1}'].value + 1
    #     except TypeError:
    #         id = initialID
    #     if "output" in pdf_files[y] or "redacted" in pdf_files[y]:
    #         continue
    #     else:
    #         string = pdf_files[y][:-4]
    #         string = string.split(letter) 
    #         string = string[-1].lstrip('0')
    #         if "a" not in string and "b" not in string:
    #             if id != int(string.replace("a", "").replace("b", "")):
    #                 continue
    #             vals, flag, warFlag = createRecord(file_path, id, cemetery)
    #             redactedFile = redact(file_path, flag, cemetery, letter)
    #             link_text = "PDF Image"
    #             worksheet.cell(row=rowIndex, column=15).value = link_text
    #             worksheet.cell(row=rowIndex, column=15).font = Font(underline="single", color="0563C1")
    #             worksheet.cell(row=rowIndex, column=15).hyperlink = redactedFile
    #             counter = 1
    #             worksheet.cell(row=rowIndex, column=counter, value=id)
    #             counter += 1
    #             for x in vals:
    #                 worksheet.cell(row=rowIndex, column=counter, value=x)
    #                 counter += 1
    #             # Grey if record adjusted using comparison
    #             if warFlag:
    #                 highlight_color = PatternFill(start_color="899499", end_color="899499", fill_type="solid")
    #                 for colIndex in range(2, 15):
    #                     cell = worksheet.cell(row=rowIndex, column=colIndex)
    #                     cell.fill = highlight_color
    #                     cell = worksheet.cell(row=rowIndex, column=16)
    #                     cell.fill = highlight_color
    #             # Purple if cemetery does not match
    #             if (worksheet[f'{"N"}{rowIndex}'].value) != cemetery:
    #                 highlight_color = PatternFill(start_color="CF9FFF", end_color="CF9FFF", fill_type="solid")
    #                 for colIndex in range(2, 15):
    #                     cell = worksheet.cell(row=rowIndex, column=colIndex)
    #                     cell.fill = highlight_color
    #                     cell = worksheet.cell(row=rowIndex, column=16)
    #                     cell.fill = highlight_color
    #             # Light Blue if record has no DOD
    #             if (worksheet[f'{"I"}{rowIndex}'].value) == "":
    #                 highlight_color = PatternFill(start_color="A7C7E7", end_color="A7C7E7", fill_type="solid")
    #                 for colIndex in range(2, 15):
    #                     cell = worksheet.cell(row=rowIndex, column=colIndex)
    #                     cell.fill = highlight_color
    #                     cell = worksheet.cell(row=rowIndex, column=16)
    #                     cell.fill = highlight_color
    #             # Yellow if record last name does not match
    #             try:
    #                 if (worksheet[f'B{rowIndex}'].value)[0] != letter:
    #                     highlight_color = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    #                     for colIndex in range(2, 15):
    #                         cell = worksheet.cell(row=rowIndex, column=colIndex)
    #                         cell.fill = highlight_color
    #                         cell = worksheet.cell(row=rowIndex, column=16)
    #                     cell.fill = highlight_color
    #             except IndexError:
    #                     highlight_color = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    #                     for colIndex in range(2, 15):
    #                         cell = worksheet.cell(row=rowIndex, column=colIndex)
    #                         cell.fill = highlight_color
    #                         cell = worksheet.cell(row=rowIndex, column=16)
    #                         cell.fill = highlight_color
    #             id += 1
    #             rowIndex += 1
    #         else:
    #             if id != int(string.replace("a", "").replace("b", "")):
    #                 continue
    #             else:
    #                 if "a" in string:
    #                     if (file_path.replace("a.pdf", "") in pdf_files):
    #                         continue
    #                     pathA = file_path
    #                     vals1, flag, warFlag = tempRecord(file_path, "a", id, cemetery)
    #                     redactedFile = redact(file_path, flag, cemetery, letter)
    #                 if "b" in string:
    #                     if (file_path.replace("b.pdf", "") in pdf_files):
    #                         continue
    #                     vals2, flag, warFlagB = tempRecord(file_path, "b", id, cemetery)
    #                     if not warFlag or not warFlagB:
    #                         warFlag = False
    #                     else:
    #                         warFlag = True
    #                     redactedFile = redact(file_path, flag, cemetery, letter)
    #                     mergeRecords(vals1, vals2, rowIndex, id, warFlag)
    #                     mergeImages(pathA, file_path, cemetery, letter)
    #                     link_text = "PDF Image"
    #                     worksheet.cell(row=rowIndex, column=15).value = link_text
    #                     worksheet.cell(row=rowIndex, column=15).font = Font(underline="single", color="0563C1")
    #                     worksheet.cell(row=rowIndex, column=15).hyperlink = redactedFile.replace("b redacted.pdf", " redacted.pdf")
    #                     id += 1
    #                     rowIndex += 1
    #     workbook.save('Veterans.xlsx')

if __name__ == "__main__":
    main()