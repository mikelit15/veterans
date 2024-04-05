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
    suffi = ["Jr.", "Sr.", "I.", "II.", "III.", "IV.", "V."]
    if suffix not in suffi:
        if firstName == "":
            firstName = suffix
        if middleName == "":
            middleName = suffix
        if lastName == "":
            lastName = suffix
        suffix = ""
    for x in suffi:
        if x in middleName:
         pass   
    for x in suffi:
        if x in value:
            suffiFlag = True
            value = value.replace(x, "")
    if value.count(".") == 1 and value.count(",") == 2:
        temp = middleName
        middleName = firstName
        firstName = lastName
        lastName = temp
    if value.count(".") == 1 and not value.count(",") == 1:
        temp = middleName
        middleName = lastName
        lastName = firstName
        firstName = temp
    finalVals.append(re.sub(r"[^a-zA-Z' ]", '', lastName))
    finalVals.append(re.sub(r"[^a-zA-Z']", '', firstName))
    finalVals.append(middleName.replace("0", "O"))
    finalVals.append(suffix)
    return finalVals

# print(parseName("Bamberger, Judge Jr."))
# Example usage
# print(parseName("Smith, John"))
# print(parseName("Smith, John D."))
# print(parseName("Smith, John, D."))
# print(parseName("Smith, John Doe"))
# print(parseName("Smith. John Doe"))
print(parseName("Smith, John, D., Jr."))
print(parseName("Smith. John D."))
print(parseName("John D. Smith"))
print(parseName("Smith. John"))
print(parseName("Smith John"))
print(parseName("Smith John D. Jr."))
print(parseName("Smith John Jr."))
print(parseName("Smith, John Doe Jr."))
print(parseName("Smith, John, Sr."))
print(parseName("Smith John D. Sr."))

