import probablepeople
from collections import OrderedDict
import re

def parseName(name):
    finalVals = []
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
    return finalVals

# Example usage
print(parseName('Wahl, Erwin, J.'))
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
# print(parseName("Smith John Doe"))
# print(parseName("Smith John D. Jr."))
# print(parseName("Smith John Jr."))
# print(parseName("Smith, John Doe Jr."))
# print(parseName("Smith, John, Sr."))
# print(parseName("Smith John D. Sr."))
# print(parseName("Bamberger, Judge Jr."))
# print(parseName("St. John, Harold S."))
# print(parseName("St. Clair James"))