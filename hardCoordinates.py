import pytesseract
from PIL import Image, ImageEnhance
import openpyxl
import os

def enhance(worksheet):
    # current_directory = r"//ucclerk/pgmdoc/Veterans"
    current_directory = os.getcwd()
    png_files = [file for file in os.listdir(current_directory) if file.lower().endswith('.png')]
    row_index = 2
    for x in png_files:
        # image = Image.open(r"//ucclerk/pgmdoc/Veterans/"+x)
        image = Image.open(x)
        # enhancer = ImageEnhance.Contrast(image)
        # image = enhancer.enhance(3.0)
        count = 0
        for x in coords:
            cropped_image = image.crop(x)
            text = pytesseract.image_to_string(cropped_image, lang='eng')
            filtered_text = text.replace("♀", "").replace("\n", "").strip()
            data[count] = filtered_text
            count+=1 
            worksheet.cell(row=row_index, column=count, value=filtered_text)
        row_index += 1
coords = [(205, 194, 1153, 258), (1357, 196, 1710, 254), (372, 263, 1777, 309),
          (321, 321, 901, 367), (1072, 319, 1783, 370), (211, 380, 553, 426), 
          (624, 378, 1762, 426), (371, 434, 1004, 488), (1505, 438, 1730, 516),
          (239, 497, 651, 545), (697, 499, 812, 545), (872, 496, 1490, 545), 
          (193, 535, 1109, 603), (1261, 555, 1781, 600), (263, 614, 433, 644), 
          (584, 616, 750, 665), (899, 614, 1056, 663), (1178, 618, 1354, 662), 
          (1541, 618, 1721, 662), (330, 676, 1774, 716), (450, 730, 1115, 776),
          (1217, 730, 1743, 776), (273, 786, 916, 838), (1138, 785, 1736, 838),
          (513, 844, 1068, 895), (1315, 844, 1738, 895)]
name = (205, 194, 1153, 258)
social = (1357, 196, 1710, 254)
homeAddress = (372, 263, 1777, 309)
nextKin = (321, 321, 901, 367)
address = (1072, 319, 1783, 370)
born = (211, 380, 553, 426)
at = (624, 378, 1762, 426)
dod = (371, 434, 1004, 488)
marker = (1505, 438, 1730, 516)
buried = (239, 497, 651, 545)
year = (697, 499, 812, 545)
cemetary = (872, 496, 1490, 545)
city = (193, 535, 1109, 603)
county = (1261, 555, 1781, 600)
division = (263, 614, 433, 644)
section = (584, 616, 750, 665) #not detecting
lotNum = (899, 614, 1056, 663) #not detecting
block = (1178, 618, 1354, 662)
graveNum = (1541, 618, 1721, 662)
warRecord = (330, 676, 1774, 716) #enhance
branch = (450, 730, 1115, 776)
rank = (1217, 730, 1743, 776)
enlisted = (273, 788, 916, 838)
discharged = (1138, 786, 1736, 836)
infoGivenBy = (513, 844, 1068, 895)
relationship = (1315, 844, 1738, 895)
remarks = ()



data = [name, social, homeAddress, nextKin, address, born, at,
        dod, marker, buried, year, cemetary, city, county, division, 
        section, lotNum, block, graveNum, warRecord, branch, rank,
        enlisted, discharged, infoGivenBy, relationship]
# print(enhance('1.jpg'))

workbook = openpyxl.load_workbook('Veterans.xlsx')
worksheet = workbook['Sheet1']
enhance(worksheet)
for x in data:
    print(x)
workbook.save('Veterans.xlsx')
