from nameparser.config import CONSTANTS
from nameparser import HumanName
import re

def parseName(value):
    finalVals = []
    suffiFlag = False
    value = value.replace("NAME", "").replace("Name", "").replace("name", "")\
        .replace("\n", " ").replace("SERIAL", "").replace("Serial", "").replace("serial", "")
    CONSTANTS.force_mixed_case_capitalization = True
    name = HumanName(value)
    flag = True
    if name.last.isupper():
        name.last = name.last[0] + name.last[1:].lower()
    tempLast = name.last[0]
    for letter in name.last[1:]:
        if letter.isupper() and flag:
            tempLast += " "
            flag = False
        elif letter == " ":
            flag = False
        tempLast += letter
    name.last = tempLast
    name.capitalize(force=True)
    firstName = name.first
    middleName = name.middle
    lastName = name.last
    suffix = name.suffix
    title = name.title
    if title and not suffix:
        suffix = title
    suffi = ["Jr", "Sr", "I", "II", "III", "IV", "V"]
    temp = value.replace("Jr", "").replace("Sr", "").replace("I", "").replace("II", "")\
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
        middleName = suffix.replace("Sr.", "").replace("Jr.", "").replace("I.", "").replace("II.", "")\
        .replace("III.", "").replace("IV.", "").replace("V.", "")
        suffix = suffix.replace(middleName, "")
    elif len(suffix) > 0 and middleName:
        suffix = suffix.replace(", ", "")
        # suffix = suffix.replace(middleName, "")
    for x in suffi:
        if x in value:
            suffiFlag = True
    if suffiFlag and "." not in temp and "," not in temp:
        if len(middleName) > 0:
            firstName = name.middle
            middleName = name.last
            lastName = name.first
        else:
            firstName = name.last
            lastName = name.first
    # elif "." not in temp and "," not in temp:
    #     if len(middleName) > 0:
    #         firstName = name.middle
    #         middleName = name.last
    #         lastName = name.first
    #     else:
    #         lastName = name.first
    #         firstName = name.last
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
    if middleName in suffi:
        suffi = middleName
        middleName = ""
    suffix = re.sub(r"[^a-zA-Z']", '', suffix)
    if suffix:
        if not "." in suffix:
            suffix += "."
    try:
        if middleName[-1] == "/":
            middleName = middleName[:-1]
    except IndexError:
        pass
    middleName = re.sub(r"[^a-zA-Z']", '', middleName)
    if len(middleName) == 1:
        middleName += "."
    if firstName.replace(".", "").lower() == "wm":
        firstName = "William"
    elif firstName.replace(".", "").lower() == "chas":
        firstName = "Charles"
    elif firstName.replace(".", "").lower() == "geo":
        firstName = "George"
    elif firstName.replace(".", "").lower() == "thos":
        firstName = "Thomas"
    elif firstName.replace(".", "").lower() == "jos":
        firstName = "Joseph"
    elif firstName.replace(".", "").lower() == "edw":
        firstName = "Edward"
    elif firstName.replace(".", "").lower() == "benj":
        firstName = "Benjamin" 
    lastName = lastName.replace("' ", "'")
    lastName = lastName[0].upper() + lastName[1:]
    if len(lastName) == 1 and middleName:
        temp = lastName
        lastName = middleName
        middleName = temp + "."
    if firstName:
        if not firstName[0].isalpha():
            firstName = firstName[1:]
        if not firstName[-1].isalpha():
            firstName = firstName[:-1]
    if middleName:
        if not middleName[0].isalpha():
            middleName = middleName[1:]
        if not middleName[-1].isalpha() and middleName[-1] != ".":
            middleName = middleName[:-1]
    if lastName:
        if not lastName[0].isalpha():
            lastName = lastName[1:]
        if not lastName[-1].isalpha():
            lastName = lastName[:-1]
    if "Mc" in lastName:
        lastName = lastName.replace(" ", "")
    finalVals.append(re.sub(r"[^a-zA-Z' ]", '', lastName))
    finalVals.append(re.sub(r"[^a-zA-Z']", '', firstName))
    finalVals.append(middleName.replace("0", "O"))
    finalVals.append(suffix)
    return finalVals

print(parseName("Bamberger, Judge Jr."))
# Example usage
# print(parseName("Smith, John"))
# print(parseName("Smith, John D."))
# print(parseName("Smith, John, D."))
# print(parseName("Smith, John Doe"))
# print(parseName("Smith. John Doe"))
# print(parseName("Smith, John, D., Jr."))
# print(parseName("Smith. John D."))
# print(parseName("John D. Smith"))
# print(parseName("Smith. John"))
# print(parseName("Smith John"))
# print(parseName("Smith John D. Jr."))
# print(parseName("Smith John Jr."))
# print(parseName("Smith, John Doe Jr."))
# print(parseName("Smith, John, Sr."))
# print(parseName("Smith John D. Sr."))

