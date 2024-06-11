from nameparser import HumanName
from nameparser.config import CONSTANTS
import re
from collections import OrderedDict
import probablepeople

'''
Parses and formats a full name into its constituent parts: first name, middle name, 
last name, and suffix. The function handles various naming conventions and appends 
the formatted name components to a list.

@param finalVals (list) - List where the processed name components will be appended
@param value (str) - The extracted raw name to be processed

@author Mike
''' 
# def nameRule(finalVals, value):
#     suffiFlag = False
#     value = value.replace("NAME", "").replace("Name", "").replace("name", "")\
#         .replace("\n", " ").replace("SERIAL", "").replace("Serial", "").replace("serial", "")
#     CONSTANTS.force_mixed_case_capitalization = True
#     name = HumanName(value)
#     flag = True
#     if name.last.isupper():
#         name.last = name.last[0] + name.last[1:].lower()
#     tempLast = name.last[0]
#     for letter in name.last[1:]:
#         if letter.isupper() and flag:
#             tempLast += " "
#             flag = False
#         elif letter == " ":
#             flag = False
#         tempLast += letter
#     name.last = tempLast
#     name.capitalize(force=True)
#     firstName = name.first
#     middleName = name.middle
#     lastName = name.last
#     suffix = name.suffix
#     title = name.title
#     suffi = ["Jr", "Sr", "I", "II", "III", "IV"]
#     temp = value.replace("Jr.", "").replace("Sr.", "").replace("I.", "").replace("II.", "")\
#         .replace("III.", "").replace("IV.", "")
#     temp = value.replace("Jr", "").replace("Sr", "").replace("I", "").replace("II", "")\
#         .replace("III", "").replace("IV", "")
#     if ("," in value and "." in firstName):
#         if len(middleName) > 0:
#             middleName = name.first
#             firstName = name.middle
#         else:
#             lastName = name.last
#             firstName = name.first
#     elif ("." in firstName):
#         if len(middleName) > 0:
#             lastName = name.first
#             firstName = name.middle
#             middleName = name.last
#         else:
#             lastName = name.first
#             firstName = name.last
#     elif ("." in lastName):
#         if len(middleName) == 0:
#             lastName = name.last
#             firstName = name.first
#         else:
#             lastName = name.first
#             middleName = name.last
#             firstName = name.middle
#     elif len(suffix) > 0 and not middleName:
#         suffix = suffix.replace(", ", "")
#         middleName = suffix.replace("Sr.", "").replace("Jr.", "").replace("I.", "").replace("II.", "")\
#         .replace("III.", "").replace("IV.", "")
#         middleName = suffix.replace("Sr", "").replace("Jr", "").replace("I", "").replace("II", "")\
#         .replace("III", "").replace("IV", "")
#         suffix = suffix.replace(middleName, "")
#     elif len(suffix) > 0 and middleName:
#         suffix = suffix.replace(", ", "")
#         # suffix = suffix.replace(middleName, "")
#     for x in suffi:
#         if x in value:
#             suffiFlag = True
#     if suffiFlag and "." not in temp and "," not in temp:
#         if len(middleName) > 0:
#             firstName = name.middle
#             middleName = name.last
#             lastName = name.first
#         else:
#             firstName = name.last
#             lastName = name.first
#     # elif "." not in temp and "," not in temp:
#     #     if len(middleName) > 0:
#     #         firstName = name.middle
#     #         middleName = name.last
#     #         lastName = name.first
#     #     else:
#     #         lastName = name.first
#     #         firstName = name.last
#     if len(middleName) > 2:
#         middleName = middleName.replace(".", "")
#     dots = 0
#     for x in firstName:
#         if (x == "."):
#             dots+=1
#     if (dots == 2):
#         firstName = firstName.upper()
#     else:
#         firstName = firstName.replace(".", "")
#     if title and not suffix:
#         suffix = title
#     if middleName in suffi:
#         suffi = middleName
#         middleName = ""
#     suffix = re.sub(r"[^a-zA-Z']", '', suffix)
#     if suffix:
#         if not "." in suffix:
#             suffix += "."
#     try:
#         if middleName[-1] == "/":
#             middleName = middleName[:-1]
#     except IndexError:
#         pass
#     middleName = re.sub(r"[^a-zA-Z']", '', middleName)
#     if len(middleName) == 1:
#         middleName += "."
#     if firstName.replace(".", "").lower() == "wm":
#         firstName = "William"
#     elif firstName.replace(".", "").lower() == "chas":
#         firstName = "Charles"
#     elif firstName.replace(".", "").lower() == "geo":
#         firstName = "George"
#     elif firstName.replace(".", "").lower() == "thos":
#         firstName = "Thomas"
#     elif firstName.replace(".", "").lower() == "jos":
#         firstName = "Joseph"
#     elif firstName.replace(".", "").lower() == "edw":
#         firstName = "Edward"
#     elif firstName.replace(".", "").lower() == "benj":
#         firstName = "Benjamin" 
#     lastName = lastName.replace("' ", "'")
#     lastName = lastName[0].upper() + lastName[1:]
#     if len(lastName) == 1 and middleName:
#         temp = lastName
#         lastName = middleName
#         middleName = temp + "."
#     if firstName:
#         if not firstName[0].isalpha() or firstName[0] == " ":
#             firstName = firstName[1:]
#         if not firstName[-1].isalpha() or firstName[0] == " ":
#             firstName = firstName[:-1]
#     if middleName:
#         if not middleName[0].isalpha() or middleName[0] == " ":
#             middleName = middleName[1:]
#         if (not middleName[-1].isalpha() and middleName[-1] != ".") or middleName[-1] == " ":
#             middleName = middleName[:-1]
#     if lastName:
#         if not lastName[0].isalpha() or lastName[0] == " ":
#             lastName = lastName[1:]
#         if not lastName[-1].isalpha() or lastName[-1] == " ":
#             lastName = lastName[:-1]
#     if "Mc" in lastName:
#         lastName = lastName.replace(" ", "")
#     finalVals.append(re.sub(r"[^a-zA-Z' ]", '', lastName))
#     finalVals.append(re.sub(r"[^a-zA-Z']", '', firstName))
#     finalVals.append(middleName.replace("0", "O"))
#     finalVals.append(suffix)
def nameRule(finalVals, name):
    suffi = ["jr", "sr", "i", "ii", "iii", "iv", "jr.", "sr.", "i.", "ii.", "iii.", "iv."]
    multi_part_surname_prefixes = ['St.', 'De', 'O\'', 'Van']

    def preprocess_name(name):
        name = re.sub(r"[^a-zA-Z',. ]", '', name)
        name = name.replace("\n", ", ")
        if name.count('.') == 1 and name.count(',') == 0 and name.split()[-1].lower() not in suffi:
            parts = name.split('.')
            if len(parts) == 2:
                if " " in parts[0]:
                    split = name.split(" ")
                    surname = split.pop(0)
                    remaining_names = split[1] + " " + split[0]
                else:
                    surname = parts[0].strip()
                    remaining_names = parts[1].strip()
                return f"{surname}, {remaining_names}"
        return name

    def format_name(name):
        name_parts = name.split()
        formatted_name = []
        for part in name_parts:
            is_multi_part_surname_prefix = False
            for prefix in multi_part_surname_prefixes:
                if prefix.lower() in part.lower():
                    is_multi_part_surname_prefix = True
                    break
            if is_multi_part_surname_prefix:
                if "\'" in part:
                    split = part.split('\'')
                    split[1] = split[1].capitalize()
                    split = '\''.join(split)
                    formatted_name.append(split)
                elif "." in part:
                    split = part.split('.')
                    split[1] = split[1].capitalize()
                    split = '. '.join(split)
                    formatted_name.append(split)
                else:
                    formatted_name.append(part)
            else:
                formatted_name.append(part.capitalize())
        return ' '.join(formatted_name)
    
    def handle_repeated_label_error(name):
        suffix = ""
        tokens = name.split(',')
        if len(tokens) == 2:
            surname = tokens[0].strip()
            given_names = tokens[1].strip().split()
            for x in range(len(given_names)):
                given_names[x] = given_names[x].replace(" ", "")
                if given_names[x].lower() in suffi:
                    suffix = given_names.pop(x)
            if len(given_names) == 2:
                return OrderedDict([
                    ('Surname', surname),
                    ('GivenName', given_names[0]),
                    ('MiddleName', given_names[1]),
                    ('SuffixGenerational', suffix)
                ])
            elif len(given_names) == 3 and given_names[2].lower() in suffi:
                return OrderedDict([
                    ('Surname', surname),
                    ('GivenName', given_names[0]),
                    ('MiddleName', given_names[1]),
                    ('SuffixGenerational', given_names[2] + '.')
                ])
            elif len(given_names) == 3:
                return OrderedDict([
                    ('Surname', surname),
                    ('GivenName', given_names[0]),
                    ('MiddleInitial', given_names[1]),
                    ('SuffixGenerational', given_names[2])
                ])
        elif len(tokens) == 3:
            surname = tokens[0].strip()
            firstName = tokens[1].strip()
            middleName = tokens[2].strip()
            return OrderedDict([
                    ('Surname', surname),
                    ('GivenName', firstName),
                    ('MiddleName', middleName),
                    ('SuffixGenerational', suffix)
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
                    ('GivenName', firstName),
                    ('MiddleName', middleName),
                    ('SuffixGenerational', suffix)
                ])
        return None

    def handle_no_comma_name(name):
        tokens = name.split()
        if len(tokens) >= 2:
            for prefix in multi_part_surname_prefixes:
                suffix = ""
                if name.startswith(prefix) or f" {prefix} " in name:
                    surname_tokens = []
                    while tokens and (tokens[0].startswith(prefix) or any(tokens[0].startswith(p) for p in multi_part_surname_prefixes)):
                        surname_tokens.append(tokens.pop(0))
                    surname = ' '.join(surname_tokens)
                    surname += f" {tokens.pop(0)}"
                    for x in range(len(tokens)):
                        if tokens[x].lower() in suffi:
                            suffix = tokens.pop(x)
                    if tokens:
                        given_name = tokens.pop(0)
                    else:
                        given_name = ''
                    middle_name = tokens.pop(0) if tokens else ''
                    return OrderedDict([
                        ('Surname', surname),
                        ('GivenName', given_name),
                        ('MiddleName', middle_name),
                        ('SuffixGenerational', suffix)
                    ])

            if len(tokens) == 4 and tokens[-1].lower() in suffi:
                return OrderedDict([
                    ('Surname', tokens[0]),
                    ('GivenName', tokens[1]),
                    ('MiddleInitial', tokens[2]),
                    ('SuffixGenerational', tokens[3] if tokens[3].endswith('.') else tokens[3] + '.')
                ])
            elif len(tokens) == 3 and tokens[-1].lower() in suffi:
                return OrderedDict([
                    ('Surname', tokens[0]),
                    ('GivenName', tokens[1]),
                    ('MiddleName', ''),
                    ('SuffixGenerational', tokens[2] if tokens[2].endswith('.') else tokens[2] + '.')
                ])
            elif len(tokens) == 3:
                if len(tokens[1]) < 3:
                    return OrderedDict([
                        ('Surname', tokens[0]),
                        ('GivenName', tokens[2]),
                        ('MiddleName', tokens[1]),
                        ('SuffixGenerational', '')
                    ])
                else:
                    return OrderedDict([
                        ('Surname', tokens[0]),
                        ('GivenName', tokens[2]),
                        ('MiddleName', tokens[1]),
                        ('SuffixGenerational', '')
                    ])
            elif len(tokens) == 2:
                return OrderedDict([
                    ('Surname', tokens[0]),
                    ('GivenName', tokens[1]),
                    ('MiddleName', ''),
                    ('SuffixGenerational', '')
                ])
            elif len(tokens) == 3 and '.' in tokens[0]:
                return OrderedDict([
                    ('Surname', tokens[0] + " " + tokens[1]),
                    ('GivenName', tokens[2]),
                    ('MiddleName', ''),
                    ('SuffixGenerational', '')
                ])
            elif len(tokens) == 2 and '.' in tokens[0]:
                return OrderedDict([
                    ('Surname', tokens[0] + " " + tokens[1]),
                    ('GivenName', ''),
                    ('MiddleName', ''),
                    ('SuffixGenerational', '')
                ])
        return None

    name = preprocess_name(name)
    name = format_name(name)
    
    if "," not in name and "." not in name:
        result = handle_no_comma_name(name)
    else:
        result = handle_repeated_label_error(name)
    #     try:
    #         result, category = probablepeople.tag(name, type="person")
    #     except probablepeople.RepeatedLabelError:
    #         result = handle_repeated_label_error(name)
    #     except Exception as e:
    #         result = None
    #     if not result:
    #         result = handle_no_comma_name(name)
    lastName = result.get('Surname', result.get('PrefixOther', '').replace(".", ""))
    firstName = result.get('GivenName', '')
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
    finalVals.append(re.sub(r"[^a-zA-Z'. ]", '', lastName))
    finalVals.append(re.sub(r"[^a-zA-Z']", '', firstName))
    finalVals.append(re.sub(r"[^a-zA-Z'. ]", '', middleName))
    finalVals.append(re.sub(r"[^a-zA-Z'.]", '', suffix))