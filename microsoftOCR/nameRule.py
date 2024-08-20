import re
from collections import OrderedDict

'''
Parses and formats a full name into its constituent parts: first name, middle name, 
last name, and suffix. The function handles various naming conventions and appends 
the formatted name components to a list.

@param finalVals (list) - List where the processed name components will be appended
@param value (str) - The extracted raw name to be processed

@author Mike
''' 
def nameRule(finalVals, name):
    suffi = ["jr", "sr", "i", "ii", "iii", "iv", "jr.", "sr.", "i.", "ii.", "iii.", "iv."]
    multiPartSurnamePrefixes = ['St.', 'De', 'O\'', 'Van', "Di", "Del", "Mc", "Le"]

    def preprocessName(name):
        name = re.sub(r"\(.*?\)", "", name)
        name = re.sub(r"[^a-zA-Z',. ]", '', name)
        name = name.replace("\n", ", ")          
        if name.count('.') == 1 and name.count(',') == 0 and name.split()[-1].lower() not in suffi:
            parts = name.split('.')
            if len(parts) == 2:
                if " " in parts[0]:
                    split = name.split(" ")
                    surname = split.pop(0)
                    remainingNames = split[1] + " " + split[0]
                else:
                    surname = parts[0].strip()
                    remainingNames = parts[1].strip()
                return f"{surname}, {remainingNames}"
        return name

    def formatName(name):
        nameParts = name.split()
        formattedName = []
        for part in nameParts:
            isMultiPartSurnamePrefix = False
            for prefix in multiPartSurnamePrefixes:
                if prefix.lower() in part[:3].lower():
                    isMultiPartSurnamePrefix = True
                    break
            if isMultiPartSurnamePrefix:
                part = part.capitalize()
                part = part.split(prefix)
                part = part[1].capitalize()
                part = prefix + part
                nameParts[0] = part
                if "\'" in part:
                    split = part.split('\'')
                    split[1] = split[1].capitalize()
                    split = '\''.join(split)
                    formattedName.append(split)
                elif "." in part:
                    split = part.split('.')
                    split[1] = split[1].capitalize()
                    split = '. '.join(split)
                    formattedName.append(split)
                elif len(nameParts) >= 3:
                    newName = nameParts.pop(0)
                    newName += nameParts.pop(0)
                    formattedName.append(newName)
                    formattedName.append(nameParts[0])
                else:
                    count = 0
                    newString = []
                    multi = False
                    for char in part:
                        if char.isupper():
                            count += 1
                            if count == 2:
                                newString.append(' ')
                                multi = True
                        newString.append(char)
                    if multi:
                        temp = "".join(newString)
                        temp = temp.split(" ")
                        temp[0] = temp[0].capitalize()
                        temp[1] = temp[1].capitalize()
                        newLName = " ".join(temp)
                        formattedName.append(newLName)
                    else:
                        temp = "".join(newString)
                        formattedName.append(temp)
            else:
                formattedName.append(part.capitalize())
        return ' '.join(formattedName)
    
    def handleFormatError(name):
        suffix = ""
        tokens = name.split(',')
        if len(tokens) == 2:
            surname = tokens[0].strip()
            firstNames = tokens[1].strip().split()
            try:
                for x in range(len(firstNames)):
                    firstNames[x] = firstNames[x].replace(" ", "")
                    if firstNames[x].lower() in suffi:
                        suffix = firstNames.pop(x)
            except Exception:
                pass
            if len(firstNames) == 1:
                return OrderedDict([
                    ('Surname', surname),
                    ('firstName', firstNames[0]),
                    ('MiddleName', ""),
                    ('SuffixGenerational', suffix)
                ])
            elif len(firstNames) == 2:
                return OrderedDict([
                    ('Surname', surname),
                    ('firstName', firstNames[0]),
                    ('MiddleName', firstNames[1]),
                    ('SuffixGenerational', suffix)
                ])
            elif len(firstNames) == 3 and firstNames[2].lower() in suffi:
                return OrderedDict([
                    ('Surname', surname),
                    ('firstName', firstNames[0]),
                    ('MiddleName', firstNames[1]),
                    ('SuffixGenerational', firstNames[2] + '.')
                ])
            elif len(firstNames) == 3:
                flag = False
                for x in range(len(firstNames)):
                    if firstNames[x].lower() in suffi:
                        suffix = firstNames.pop(x)
                        flag = True
                if flag:
                    return OrderedDict([
                        ('Surname', surname),
                        ('firstName', firstNames[0]),
                        ('MiddleName', firstNames[1]),
                        ('SuffixGenerational', suffix)
                    ])
                else:
                    return OrderedDict([
                        ('Surname', surname),
                        ('firstName', firstNames[0]),
                        ('MiddleName', firstNames[2]),
                        ('SuffixGenerational', "")
                    ])
        elif len(tokens) == 3:
            surname = tokens[0].strip()
            firstNames = tokens[1].strip().split()
            try:
                for x in range(len(tokens)):
                    tokens[x] = tokens[x].replace(" ", "")
                    if tokens[x].lower() in suffi:
                        suffix = tokens.pop(x)
            except Exception:
                pass
            if len(firstNames) == 1 and not suffix:
                return(handleNoComma(name.replace(",", "")))
            if len(firstNames) == 1:
                return OrderedDict([
                    ('Surname', surname),
                    ('firstName', firstNames[0]),
                    ('MiddleName', ""),
                    ('SuffixGenerational', suffix)
                ])
            elif len(firstNames) == 2:
                return OrderedDict([
                    ('Surname', surname),
                    ('firstName', firstNames[0]),
                    ('MiddleName', firstNames[1]),
                    ('SuffixGenerational', suffix)
                ])
            elif len(firstNames) == 3:
                for x in range(len(firstNames)):
                    if firstNames[x].lower() in suffi:
                        suffix = firstNames.pop(x)
                return OrderedDict([
                    ('Surname', surname),
                    ('firstName', firstNames[0]),
                    ('MiddleName', firstNames[1]),
                    ('SuffixGenerational', firstNames[2] + '.')
                ])
        elif len(tokens) == 4:
            for x in range(len(tokens)):
                tokens[x] = tokens[x].replace(" ", "")
                if tokens[x].lower() in suffi:
                    suffix = tokens.pop(x)
            surname = tokens[0].strip()
            firstName = tokens[1].strip()
            middleName = tokens[2].strip()
            return OrderedDict([
                    ('Surname', surname),
                    ('firstName', firstName),
                    ('MiddleName', middleName),
                    ('SuffixGenerational', suffix)
                ])
        return None

    def handleNoComma(name):
        tokens = name.split()
        if len(tokens) >= 2:
            for prefix in multiPartSurnamePrefixes:
                suffix = ""
                if name.startswith(prefix) or f" {prefix} " in name:
                    surnameTokens = []
                    while tokens and (tokens[0].startswith(prefix) or any(tokens[0].startswith(p) for p in multiPartSurnamePrefixes)):
                        surnameTokens.append(tokens.pop(0))
                    surname = ' '.join(surnameTokens)
                    surname += f" {tokens.pop(0)}"
                    for x in range(len(tokens)):
                        if tokens[x].lower() in suffi:
                            suffix = tokens.pop(x)
                    if tokens:
                        firstName = tokens.pop(0)
                    else:
                        firstName = ''
                    middleName = tokens.pop(0) if tokens else ''
                    return OrderedDict([
                        ('Surname', surname),
                        ('firstName', firstName),
                        ('MiddleName', middleName),
                        ('SuffixGenerational', suffix)
                    ])

            if len(tokens) == 4 and tokens[-1].lower() in suffi:
                return OrderedDict([
                    ('Surname', tokens[0]),
                    ('firstName', tokens[1]),
                    ('MiddleInitial', tokens[2]),
                    ('SuffixGenerational', tokens[3] if tokens[3].endswith('.') else tokens[3] + '.')
                ])
            elif len(tokens) == 3 and tokens[-1].lower() in suffi:
                return OrderedDict([
                    ('Surname', tokens[0]),
                    ('firstName', tokens[1]),
                    ('MiddleName', ''),
                    ('SuffixGenerational', tokens[2] if tokens[2].endswith('.') else tokens[2] + '.')
                ])
            elif len(tokens) == 3:
                if len(tokens[1]) < 3:
                    return OrderedDict([
                        ('Surname', tokens[0]),
                        ('firstName', tokens[2]),
                        ('MiddleName', tokens[1]),
                        ('SuffixGenerational', '')
                    ])
                else:
                    return OrderedDict([
                        ('Surname', tokens[0]),
                        ('firstName', tokens[2]),
                        ('MiddleName', tokens[1]),
                        ('SuffixGenerational', '')
                    ])
            elif len(tokens) == 2:
                return OrderedDict([
                    ('Surname', tokens[0]),
                    ('firstName', tokens[1]),
                    ('MiddleName', ''),
                    ('SuffixGenerational', '')
                ])
            elif len(tokens) == 3 and '.' in tokens[0]:
                return OrderedDict([
                    ('Surname', tokens[0] + " " + tokens[1]),
                    ('firstName', tokens[2]),
                    ('MiddleName', ''),
                    ('SuffixGenerational', '')
                ])
            elif len(tokens) == 2 and '.' in tokens[0]:
                return OrderedDict([
                    ('Surname', tokens[0] + " " + tokens[1]),
                    ('firstName', ''),
                    ('MiddleName', ''),
                    ('SuffixGenerational', '')
                ])
        return None

    name = preprocessName(name)
    name = formatName(name)
    
    if "," not in name:
        result = handleNoComma(name)
    else:
        result = handleFormatError(name)
    lastName = result.get('Surname', result.get('PrefixOther', '').replace(".", ""))
    firstName = result.get('firstName', '')
    if lastName[-1] == ".":
        lastName = lastName[:-1]
    middleName = result.get('MiddleName', result.get('MiddleInitial', ''))
    middleName = middleName.replace("0", "O")
    suffix = result.get('SuffixGenerational', '')
    if len(middleName) == 1:
        middleName += "."
    if len(middleName) == 2:
        if middleName[-1] != ".":
            middleName = middleName[:-1] + "."
    if suffix and suffix[-1] != ".":
        suffix += "."
    if firstName.replace(".", "").lower() == "wm":
        firstName = "William"
    elif firstName.replace(".", "").lower() ==  "jas":
        firstName = "James"
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
    finalVals.append(re.sub(r"[^a-zA-Z'. ]", '', lastName))
    finalVals.append(re.sub(r"[^a-zA-Z']", '', firstName.capitalize()))
    finalVals.append(re.sub(r"[^a-zA-Z'. ]", '', middleName.capitalize()))
    finalVals.append(re.sub(r"[^a-zA-Z'.]", '', suffix))