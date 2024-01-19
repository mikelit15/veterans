import re

# war = "WW I 11/12/1414-12/414/1515"
# if war[0] == " ":
#     war = war[1:]
# ww1years = ["1914", "1915", "1916", "1917", "1918"]
# if "N." in war:
#     war = war.replace("N", "W")
# tempText = war.replace("-", "").replace(".", "").replace("#", "").replace(" ", "")
# ww1_pattern = re.compile(r'WWI|WW1|WW\s*I|World\s*War\s*(1|I)|World\s*War|WW\s*1|1')
# ww2_pattern = re.compile(r'WWII|WW2|WW\s*2|WW\s*#?II|World\s*War\s*(II|2)|2')
# if (ww2_pattern.findall(tempText)):
#     war = "World War 2"
# elif (ww1_pattern.findall(tempText)):
#     war = "World War 1"
# else:
#     war = war.replace(".", "").replace("Calv", "").replace("Vols", "")
#     if "Korea" in war:
#         war = "Korean War"
#     elif "Vietnam" in war:
#         war = "Vietnam War"
#     elif "Civil" in war:
#         war = "Civil War"
#     elif "Spanish" in war or "Amer" in war or "American" in war:
#         war = "Spanish American War"
#     elif "Peacetime" in war or "Not Listed" in war:
#         war = ""
#     words = war.split()
#     for x in words:
#         if x in ww1years:
#             war = "World War 1"
#             break
# print(war)

# war = "U S ARMY"
# value = "Co, G 1 Field Signal Battnk "
# armys = ["co", "army", "inf", "infantry", "Infan", "usa", "med", \
#          "cav", "div", "sig", "art", "corps"]
# navys = ["hospital", "navy", "naval"]
# guards = ["113", "102d", "114", "44", "181", "250", "112"]
# branch = value
# branch = branch.replace("/", " ").replace(".", " ").replace("th", "").replace("-", "").replace(",", "")
# words = branch.split()
# if war == "Civil War":
#     branch = "Army"
# for word in words:
#     word = word.lower()
#     if "Air Force" in branch:
#         branch = "Air Force"
#         break
#     if "Marine" in branch:
#         branch = "Marine Corps"
#         # break
#     if word in armys:
#         branch = "Army"
#     if word in navys:
#         branch = "Navy"
#     if word in guards:
#         branch = "National Guard"
#         break
#     if "Not Listed" in branch:
#         branch = ""
# print(branch)


# state_info = {
#     ('Alabama', 'AL'), ('Alaska', 'AK'), ('Arizona', 'AZ'), ('Arkansas', 'AR'), ('California', 'CA'),
#     ('Colorado', 'CO'), ('Connecticut', 'CT'), ('Delaware', 'DE'), ('Florida', 'FL'), ('Georgia', 'GA'),
#     ('Hawaii', 'HI'), ('Idaho', 'ID'), ('Illinois', 'IL'), ('Indiana', 'IN'), ('Iowa', 'IA'), ('Kansas', 'KS'),
#     ('Kentucky', 'KY'), ('Louisiana', 'LA'), ('Maine', 'ME'), ('Maryland', 'MD'), ('Massachusetts', 'MA'),
#     ('Michigan', 'MI'), ('Minnesota', 'MN'), ('Mississippi', 'MS'), ('Missouri', 'MO'), ('Montana', 'MT'),
#     ('Nebraska', 'NE'), ('Nevada', 'NV'), ('New Hampshire', 'NH'), ('New Jersey', 'NJ'), ('New Mexico', 'NM'),
#     ('New York', 'NY'), ('North Carolina', 'NC'), ('North Dakota', 'ND'), ('Ohio', 'OH'), ('Oklahoma', 'OK'),
#     ('Oregon', 'OR'), ('Pennsylvania', 'PA'), ('Rhode Island', 'RI'), ('South Carolina', 'SC'),
#     ('South Dakota', 'SD'), ('Tennessee', 'TN'), ('Texas', 'TX'), ('Utah', 'UT'), ('Vermont', 'VT'),
#     ('Virginia', 'VA'), ('Washington', 'WA'), ('West Virginia', 'WV'), ('Wisconsin', 'WI'), ('Wyoming', 'WY'),
#     ('United States', 'US')}
# branch = "Co.C.3N.J."
# if "Army" in branch:
#     branch = "Army"
# elif "Navy" in branch:
#     branch = "Navy"
# elif "Air Force" in branch:
#     branch = "Air Force"
# elif "USA" in branch:
#     branch = branch.replace("USA ", "")
# else:
#     company = ""
#     state = ""
#     branch = branch.replace('"', "").replace("/", " ")
#     company_letter_match = re.search(r'(Co\.|Co|C)\s*([A-Z])', branch)
#     try:
#         if len(company_letter_match.group(1)) != 1:
#             company = company_letter_match.group(2)
#         else:
#             company = company_letter_match.group(1)
#         number_match = re.search(r'(\d+)', branch)
#         reg = number_match.group(1)
#         word = branch.split()
#         for y in word:
#             y = y.replace(".", "").replace("Inf", "")
#             for z in state_info:
#                 if y == z[1]:
#                     state = z[1]
#                 elif y == z[0]:
#                     state = z[1]
#     except AttributeError:
#         number_match = re.search(r'(\d+)', branch)
#         reg = number_match.group(1)
#         word = branch.split()
#         for y in word:
#             y = y.replace(".", "").replace("Inf", "")
#             for z in state_info:
#                 if y == z[1]:
#                     state = z[1]
#                 elif y == z[0]:
#                     state = z[1]
#     if state == "":
#         branch = branch.replace(" ", "")
#         branch = branch.replace("N.Y.", " NY ").replace("N.J.", " NJ ")
#         word = branch.split()
#         for y in word:
#             for z in state_info:
#                 if y == z[1]:
#                     state = z[1]
#                 elif y == z[0]:
#                     state = z[1]
#     if state == "":
#         print("US Reg " + reg)

#     elif company == "":
#         print(state + " Reg " + reg)
#     else:
#         print(state + " Reg " + reg + " - Company " + company)

import dateparser
import re

def buriedRule(value, bYear):
    buried = value
    buriedYear = bYear
    if buried or buriedYear:
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
                        buriedYear = "19" + buriedYear
                    else:
                        buried = ""
                        buriedYear = "19" + buriedYear
            else:
                try:
                    before_comma = buried.split(',', 1)[0].strip()
                    after_comma = buried.split(',', 1)[1].strip()
                    if len(after_comma) > 4:
                        buried, buriedYear = buriedRule(before_comma, after_comma[:4]) 
                except IndexError:
                    print("buried is weird")
                    buried = ""
                    
        else:
            buried = dateparser.parse(buried)
            try:
                year = buried.strftime("%y")
                buried = buried.strftime("%m/%d/19" + year)
                buriedYear = "19" + year
            except ValueError:
                year = buried.strftime("%Y")
                buried = buried.strftime("%m/%d/%Y")
                buriedYear = "19" + year
    return buriedYear

def dodRule(finalVals, value, dob, buried, bYear):
    century = False
    birthFlag = False
    deathFlag = False
    if bYear:
        century = True
    birth = dob
    birth = birth.replace(":", ".")
    if birth[-1:] == " ":
        birth = birth.replace(" ", "")
    if len(birth) < 4:
        birth = ""
    try:
        if dateparser.parse(birth, settings={'STRICT_PARSING': True}) == None:
            birth = birth[-4:]
        else:
            birthFlag = True
            birth = dateparser.parse(birth)
    except dateparser.ParserError:
        birth = "" 
    death = value
    death = death.replace(":", ".")
    try:
        if dateparser.parse(death, settings={'STRICT_PARSING': True}) == None:
            death = death
        else:
            deathFlag = True
            death = dateparser.parse(death)
    except dateparser.ParserError:
        death = "" 
    if birth == "" and death == "":
        bYear = buriedRule(buried, bYear)
        finalVals.append("")
        finalVals.append("")
        if bYear:
            century = True
            finalVals.append("")
            finalVals.append(int(bYear))
        else:
            finalVals.append("")
            finalVals.append("")
    elif birth != "" and death == "":
        bYear = buriedRule(buried, bYear)
        if isinstance(birth, str):
            finalVals.append("")
            birth = re.sub(r'[^0-9]', '', birth)
            if(birth != ""):
                finalVals.append(int(birth))
            else:
                finalVals.append("")
            if bYear:
                century = True
                finalVals.append("")
                finalVals.append(int(bYear))
            else:
                finalVals.append("")
                finalVals.append("")
        else:
            if(birth.strftime("%Y")[:2] == "20"):
                dobYear = "19" + birth.strftime("%x")[-2:]
                finalVals.append(birth.strftime("%m/%d/" + dobYear))
                if(dobYear != ""):
                    finalVals.append(int(dobYear))
                else:
                    finalVals.append("")
                if buried:
                    if(len(buried) == 4):
                        finalVals.append("")
                        finalVals.append(int(buried))
                    else:
                        finalVals.append(buried)
                        finalVals.append(int(buried[-4:]))
                else:
                    finalVals.append("")
                    finalVals.append("")
            else:
                finalVals.append(birth.strftime("%m/%d/%Y"))
                if(birth.year != ""):
                    finalVals.append(int(birth.year))
                else:
                    finalVals.append("")
                if buried:
                    if(len(buried) == 4):
                        finalVals.append("")
                        finalVals.append(int(buried))
                    else:
                        finalVals.append(buried)
                        finalVals.append(int(buried[-4:]))
                else:
                    finalVals.append("")
                    finalVals.append("")
    elif birth == "" and death != "":
        if isinstance(death, str):
            finalVals.append("")
            finalVals.append("")
            finalVals.append("")
            death = re.sub(r'[^0-9]', '', death)
            if(death != ""):
                finalVals.append(int(death[-4:]))
            else:
                finalVals.append("")
        else:
            if(death.strftime("%Y")[:2] == "20"):
                dodYear = "19" + death.strftime("%x")[-2:]
                finalVals.append("")
                finalVals.append("")
                finalVals.append(death.strftime("%m/%d/" + dodYear))
                if(dodYear != ""):
                    finalVals.append(int(dodYear))
                else:
                    finalVals.append("")
            else:    
                finalVals.append("")
                finalVals.append("")
                finalVals.append(death.strftime("%m/%d/%Y"))
                if(death.year != ""):
                    finalVals.append(int(death.year))
                else:
                    finalVals.append("")
    else:
        try:
            pattern = re.compile(r'\b\d{1,2}/\d{1,2}/(\d{4})\b')
            match = pattern.search(dob)
            yearb = match.group(1)
            match2 = pattern.search(value)
            yeard = match2.group(1)
            # birthLen = len(birth.strftime("%x"))
            # deathLen = len(death.strftime("%x"))
            # if birthLen >= 4 and deathLen >= 4:
            #     dobYear = birth.strftime("%x")[-2:]
            #     dodYear = death.strftime("%x")[-2:]
            # if(dobYear > dodYear):
            #     yearb = "18" + dobYear
            #     yeard = "19" + dodYear
            # else:
            #     yearb = "19" + dobYear
            #     yeard = "19" + dodYear
            finalVals.append(birth.strftime("%m/%d/" + yearb))
            if(yearb != ""):
                finalVals.append(int(yearb))
            else:
                finalVals.append("")
            finalVals.append(death.strftime("%m/%d/" + yeard))
            if(yeard != ""):
                finalVals.append(int(yeard))
            else:
                finalVals.append("")
        except AttributeError:
            if not birthFlag and not deathFlag:
                finalVals.append("")
                birth = re.sub(r'[^0-9]', '', birth)
                # if len(birth) > 4:
                #     worksheet.cell(row=rowIndex, column=counter, value=int(birth[4:]))
                # else:
                if(birth != ""):
                    finalVals.append(int(birth))
                else:
                    finalVals.append("")
                death = re.sub(r'[^0-9]', '', death)
                # if len(death) > 4:
                #     worksheet.cell(row=rowIndex, column=counter, value=int(death[-4:]))
                # else:
                finalVals.append("")
                if(death != ""):
                    finalVals.append(int(death))
                else:
                    finalVals.append("")
            elif birthFlag and not deathFlag:
                dodYear = death[-2:]
                dobYear = birth.strftime("%x")[-2:]
                dobCent = birth.strftime("%Y")[:2]
                if dobYear > dodYear:
                    yearb = "18" + dobYear
                    yeard = "19" + dodYear
                else:
                    if dobCent == "18":
                        yearb = "18" + dobYear
                        yeard = "18" + dodYear
                    else:    
                        yearb = "19" + dobYear
                        yeard = "19" + dodYear
                finalVals.append(birth.strftime("%m/%d/" + yearb))
                if(yearb != ""):
                    finalVals.append(int(yearb))
                else:
                    finalVals.append("")
                finalVals.append("")
                if(yeard != ""):
                    finalVals.append(int(yeard))
                else:
                    finalVals.append("")
            elif not birthFlag and deathFlag:
                if len(birth) > 4:
                    index = birth.find("18")
                    if index:
                        dobYear = birth[index+2:index+4]
                        dobYear = re.sub(r'[^0-9]', '', dobYear)
                    elif not index:
                        index = birth.find("19")
                        dobYear = birth[index+2:index+4]
                        dobYear = re.sub(r'[^0-9]', '', dobYear)
                    else:
                        dobYear = birth[:4]
                else:
                    dobYear = birth[-2:]
                dodYear = death.strftime("%x")[-2:]
                dodCent = death.strftime("%Y")[:2]
                if dobYear > dodYear:
                    yearb = "18" + dobYear
                    yeard = "19" + dodYear
                else:
                    if dodCent == "18":
                        yearb = "18" + dobYear
                        yeard = "18" + dodYear
                    else:    
                        if dobYear == "":
                            yearb = ""
                            yeard = "19" + dodYear
                        else:
                            yearb = "19" + dobYear
                            yeard = "19" + dodYear
                birth = re.sub(r'[^0-9]', '', birth)
                finalVals.append("")
                if(yearb != ""):
                    finalVals.append(int(yearb))
                else:
                    finalVals.append("")
                finalVals.append(death.strftime("%m/%d/" + yeard))
                if(yeard != ""):
                    finalVals.append(int(yeard))
                else:
                    finalVals.append("")
    if century:
        finalVals.append("Yes")
    else:
        finalVals.append("No Field")



finalVals = []
birth = "11/21/1843"
death = "05/21/1896"
buried = ""
bYear = ""
dodRule(finalVals, death, birth, buried, bYear)
for x in finalVals:
    print(x)

# from nameparser import HumanName
# from nameparser.config import CONSTANTS
# import re
# import dateparser
# century = False
# birthFlag = False
# deathFlag = False
# buried = ""
# war = "" 
# wars19 = ["Vietnam War", "Korean War", "World War 2"]
# birth = ""
# death = '"arch, 19,1873'
# birth = birth.replace(":", ".")
# death = death.replace(":", ".")
# if birth[-1:] == " ":
#     birth = birth.replace(" ", "")
# try:
#     if dateparser.parse(birth, settings={'STRICT_PARSING': True}) == None:
#         birth = birth
#     else:
#         birthFlag = True
#         birth = dateparser.parse(birth)
# except dateparser.ParserError:
#     birth = "" 
# try:
#     if dateparser.parse(death, settings={'STRICT_PARSING': True}) == None:
#         death = death
#     else:
#         deathFlag = True
#         death = dateparser.parse(death)
# except dateparser.ParserError:
#     death = "" 
# if birth == "" and death == "":
#     print(buried)
# elif birth != "" and death == "":
#     if isinstance(birth, str):
#         birth = re.sub(r'[^0-9]', '', birth)
#         print(birth)
#     else:
#         if(birth.strftime("%Y")[:2] == "20"):
#             dobYear = "19" + birth.strftime("%x")[-2:]
#             print(birth.strftime("%m/%d/") + dobYear)
#             print(dobYear)
#         else:
#             print(birth.strftime("%m/%d/%Y"))
#             print(birth.year)
# elif birth == "" and death != "":
#     if isinstance(death, str):
#         death = re.sub(r'[^0-9]', '', death)
#         print(death[-4:])
#     else:
#         if(death.strftime("%Y")[:2] == "20"):
#             dodYear = "19" + death.strftime("%x")[-2:]
#             print(death.strftime("%m/%d/") + dodYear)
#             print(dodYear)
#         else:
#             print(death.strftime("%m/%d/%Y"))
#             print(death.year)
# else:
#     try:
#         birthLen = len(birth.strftime("%x"))
#         deathLen = len(death.strftime("%x"))
#         if birthLen >= 4 and deathLen >= 4:
#             dobYear = birth.strftime("%x")[-2:]
#             dodYear = death.strftime("%x")[-2:]
#         if(dobYear > dodYear):
#             year = "18" + dobYear
#             year3 = "19" + dodYear
#         else:
#             year = "19" + dobYear
#             year3 = "19" + dodYear
#         print(birth.strftime("%m/%d/" + year))
#         print(year)
#         print(death.strftime("%m/%d/" + year3))
#         print(year3)
#     except AttributeError:
#         if not birthFlag and not deathFlag:
#             birth = re.sub(r'[^0-9]', '', birth)
#             death = re.sub(r'[^0-9]', '', death)
#             print(birth)
#             print(death)
#         elif birthFlag and not deathFlag:
#             dodYear = death[-2:]
#             dobYear = birth.strftime("%x")[-2:]
#             dobCent = birth.strftime("%Y")[:2]
#             if dobYear > dodYear:
#                 year = "18" + dobYear
#                 year3 = "19" + dodYear
#             else:
#                 if dobCent == "18":
#                     year = "18" + dobYear
#                     year3 = "18" + dodYear
#                 else:    
#                     year = "19" + dobYear
#                     year3 = "19" + dodYear
#             print(birth.strftime("%m/%d/" + year))
#             print(year)
#             death = re.sub(r'[^0-9]', '', death)
#             print(year3)
#         elif not birthFlag and deathFlag:
#             if len(birth) > 4:
#                 index = birth.find("18")
#                 if index:
#                     dobYear = birth[index+2:index+4]
#                     dobYear = re.sub(r'[^0-9]', '', dobYear)
#                 elif not index:
#                     index = birth.find("19")
#                     dobYear = birth[index+2:index+4]
#                     dobYear = re.sub(r'[^0-9]', '', dobYear)
#                 else:
#                     dobYear = birth[:4]
#             else:
#                 dobYear = birth[-2:]
#             dodYear = death.strftime("%x")[-2:]
#             dodCent = death.strftime("%Y")[:2]
#             if dobYear > dodYear:
#                 year = "18" + dobYear
#                 year3 = "19" + dodYear
#             else:
#                 if dodCent == "18":
#                     year = "18" + dobYear
#                     year3 = "18" + dodYear
#                 else:
#                     if dobYear == "":
#                         year = ""
#                         year3 = "19" + dodYear
#                     else:
#                         year = "19" + dobYear
#                         year3 = "19" + dodYear
#             birth = re.sub(r'[^0-9]', '', birth)
#             print(year)
#             print(death.strftime("%m/%d/" + year3))
#             print(year3)

# from nameparser import HumanName
# from nameparser.config import CONSTANTS
# value = "CARLSTROM.WALTERWILLIAM  "
# CONSTANTS.force_mixed_case_capitalization = True
# name = HumanName(value)
# name.capitalize()
# firstName = name.first
# middleName = name.middle
# lastName = name.last
# suffix = name.suffix
# temp = value.replace("Jr.", "").replace("Sr.", "")
# if ("," in value and "." in firstName):
#     if len(middleName) > 0:
#         middleName = name.first
#         firstName = name.middle
#     else:
#         lastName = name.last
#         firstName = name.first
# elif ("." in firstName):
#     if len(middleName) > 0:
#         lastName = name.first
#         firstName = name.middle
#         middleName = name.last
#     else:
#         lastName = name.first
#         firstName = name.last
# elif ("." in lastName):
#     if len(middleName) == 0:
#         lastName = name.last
#         firstName = name.first
#     else:
#         lastName = name.first
#         middleName = name.last
#         firstName = name.middle
# elif len(suffix) > 0:
#     suffix = suffix.replace(", ", "")
#     middleName = suffix.replace("Sr.", "").replace("Jr.", "")
#     suffix = suffix.replace(middleName, "")
# if ("Jr." in value or "Sr." in value) and "." not in temp and "," not in temp:
#     if len(middleName) > 0:
#         firstName = name.middle
#         middleName = name.last
#         lastName = name.first
#     else:
#         firstName = name.last
#         lastName = name.first
# elif "." not in temp and "," not in temp:
#     if len(middleName) > 0:
#         firstName = name.middle
#         middleName = name.last
#         lastName = name.first
#     else:
#         lastName = name.first
#         firstName = name.last
# if len(middleName) > 2:
#     middleName = middleName.replace(".", "")
# dots = 0
# for x in firstName:
#     if (x == "."):
#         dots+=1
# if (dots == 2):
#     firstName = firstName.upper()
# else:
#     firstName = firstName.replace(".", "")
# print(lastName.replace(",", "").replace(".", ""))
# print(firstName.replace(",", ""))
# print(middleName.replace(",", "."))
# print(suffix.replace(",", "."))