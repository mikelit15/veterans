from nameparser.config import CONSTANTS
from nameparser import HumanName

def parseName(value):
    finalVals = []
    CONSTANTS.force_mixed_case_capitalization = True
    name = HumanName(value)
    name.capitalize()
    firstName = name.first
    middleName = name.middle
    lastName = name.last
    suffix = name.suffix
    temp = value.replace("Jr.", "").replace("Sr.", "")
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
        middleName = suffix.replace("Sr.", "").replace("Jr.", "")
        suffix = suffix.replace(middleName, "")
    elif len(suffix) > 0 and middleName:
        suffix = suffix.replace(", ", "")
        suffix = suffix.replace(middleName, "")
    if ("Jr." in value or "Sr." in value) and "." not in temp and "," not in temp:
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
    return finalVals
# Example usage
print(parseName("Smith, John"))
print(parseName("Smith, John D."))
print(parseName("Smith, John, D."))
print(parseName("Smith, John Doe"))
print(parseName("Smith. John Doe"))
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

