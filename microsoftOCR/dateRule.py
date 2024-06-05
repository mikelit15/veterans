import dateparser
import re

'''
Processes and formats birth date information. Handles partial and full dates, converting 
them into a consistent format.

@param birth (str) - Extracted raw full birth date
@param bYear (str) - Birth year, particularly if only the year is provided
@param birthYYFlag (bool) - Flag indicating if the birth year is provided as a two-digit 
                            number

@return birth (str) - Cleaned and formatted full birth date
@return bYear (str) - Birth year, seperated from full date if applicable
@return birthYYFlag (bool) - Updated flag indicating if the birth year is provided as 
                             a two-digit number

@author Mike
''' 
def parseBirth(birth, bYear, birthYYFlag):
    if birth.count('/') == 1:
        birth = birth.replace("/", " ")
        parseBirth(birth, bYear, birthYYFlag)
    pattern = r'([A-Za-z]+),\s*(\d{4})'
    match = re.match(pattern, birth.strip())
    if birth[-1:] == " ":
        birth = birth[:-1]
    if match:
        bYear = match.group(2)
        birth = ""
    elif birth.count('/') == 2:
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
            bYear = year[:4]
            birth = birth.replace(",", " ")
            birth = dateparser.parse(birth)
            birth = birth.strftime("%m/%d/%Y")
        elif len(year) > 6:
            temp = year[:4]
            if temp.isnumeric():
                death_no_spaces = birth.replace(' ', '')
                startIndex = death_no_spaces.find(year[4:])
                num_spaces_before = birth[:startIndex].count(' ')
                adjusted_start_index = startIndex + num_spaces_before
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
                startIndex = death_no_spaces.find(year[4:])
                num_spaces_before = birth[:startIndex].count(' ')
                adjusted_start_index = startIndex + num_spaces_before
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
        if len(year) == 4 and year[0] == "1":
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
                startIndex = death_no_spaces.find(year[4:])
                num_spaces_before = birth[:startIndex].count(' ')
                adjusted_start_index = startIndex + num_spaces_before
                birth = birth[:adjusted_start_index].strip()
                birth = dateparser.parse(birth)
                birth = birth.strftime("%m/%d/%Y")
                bYear = temp
            else:
                birth = ""
        else:
            birth = ""
    return birth, bYear, birthYYFlag


'''
Processes and formats death date information. Handles partial and full dates, converting 
them into a consistent format.

@param death (str) - Extracted raw full death date 
@param dYear (str) - Death year, particularly if only the year is provided
@param deathYYFlag (bool) - Flag indicating if the death year is provided as a two-digit 
                            number
 
@return death (str) - Cleaned and formatted full death date
@return dYear (str) - Death year, seperated from full date if applicable
@return deathYYFlag (bool) - Updated flag indicating if the death year is provided as a 
                             two-digit number

@author Mike
''' 
def parseDeath(death, dYear, deathYYFlag):
    if death.count('/') == 1:
        death = death.replace("/", " ")
        parseDeath(death, dYear, deathYYFlag)
    pattern = r'([A-Za-z]+),\s*(\d{4})'
    match = re.match(pattern, death.strip())
    if death[-1:] == " ":
        death = death[:-1]
    if match:
        dYear = match.group(2)
        death = ""
    elif death.count('/') == 2:
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
            try:
                death = death.strftime("%m/%d/%Y")
            except AttributeError:
                dYear = ""
                death = ""
        elif len(year3) > 6:
            temp = year3[:4]
            if temp.isnumeric():
                death_no_spaces = death.replace(' ', '')
                startIndex = death_no_spaces.find(year3[4:])
                num_spaces_before = death[:startIndex].count(' ')
                adjusted_start_index = startIndex + num_spaces_before
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
                startIndex = death_no_spaces.find(year3[4:])
                num_spaces_before = death[:startIndex].count(' ')
                adjusted_start_index = startIndex + num_spaces_before
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
        if len(year3) == 4 and (year3[0] == "1" or year3[0] == "2"):
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
                startIndex = death_no_spaces.find(year3[4:])
                num_spaces_before = death[:startIndex].count(' ')
                adjusted_start_index = startIndex + num_spaces_before
                death = death[:adjusted_start_index].strip()
                death = dateparser.parse(death)
                death = death.strftime("%m/%d/%Y")
                dYear = temp
            else:
                death = ""
        else:
            dYear = temp[-1]
            death = ""
    else:
        death = ""
        dYear = ""
    return death, dYear, deathYYFlag


'''
Processes and formats the burial year information. It adjusts the year based on the 
century and war involvement, and reconciles it with other date-related information.

@param value (str) - Extracted raw burial date
@param cent (str) - Value in "19" field, confirms death year is in 1900s and provides year
@param warsFlag (bool) - Flag indicating involvement in a 19th century war

@return buried4Year (str) - The full year of burial, 4 digits
@return buried2Year (str) - The partial year of burial, last 2 digits if full is not 
                            available

@author Mike
'''
def buriedRule(value, cent, warsFlag):
    buried = value.replace("I", "1").replace("-", "/")
    buried4Year = ""
    buried2Year = ""
    if len(cent) == 2:
        buried4Year = "19" + cent
    elif len(cent) == 4:
        buried4Year = cent
    elif len(cent) == 3:
        buried4Year = "19" + cent[-2:]
    elif buried:
        if dateparser.parse(buried, settings={'STRICT_PARSING': True}) != None:
            if buried.count('/') == 2:
                year5 = buried.split('/')[-1]
                if year5 == buried:
                    year5 = buried.split(',')[-1]
                    if year5[0] == " ":
                        year5 = year5[1:]
                    if len(year5) == 4:
                        buried4Year = year5
                    elif len(year5) == 2:
                        buried2Year = year5
                else:
                    if year5[-1:] == " ":
                        year5 = year5[:-1]
                    if len(year5) == 4:
                        buried4Year = year5
                    elif len(year5) == 2:
                        if warsFlag:
                            buried4Year = "19" + year5
                        else:
                            buried2Year = year5
            elif "," in buried:
                year5 = buried.split(',')[-1]
                if len(year5) == 4:
                    buried4Year = year5
                elif len(year5) == 2:
                    if warsFlag:
                        buried4Year = "19" + year5
                    else:
                        buried2Year = year5
            elif " " in buried:
                if buried.count('.') == 1:
                    buried = buried.replace(".", " ")
                year5 = buried.split(' ')
                if year5[-1] == " ":
                    year5 = year5[-2]
                else:
                    year5 = year5[-1]
                if len(year5) == 4:
                    buried4Year = year5
                elif len(year5) == 2:
                    if warsFlag:
                        buried4Year = "19" + year5
                    else:
                        buried2Year = year5
    return buried4Year, buried2Year

'''
Analyzes and formats date-related information based on various rules and conditions. 
It reconciles birth and death dates with other date-related information. Handles conditions
where missing birth and/or death dates and/or years. Attempts to correct 2-digit years
using cent date, else uses buried date, else uses war info, else uses "<" comparison and 
this comparison calls for the activation of warFlag signaling to highlight the row due to 
possible wrong century.

@param finalVals (list) - List where the processed date information will be appended
@param value (str) - Current value being processed (typically related to death date)
@param dob (str) - Raw date of birth
@param buried (str) - Raw date of burial
@param cent (str) -  Value in "19" field, confirms death year is in 1900s and provides year
@param war (str) - War during which the person served, if applicable
@param app (str) - The 4 digit year when the app was sent, used for correcting 2 digit years

@return warFlag (bool) - Flag indicating if a 2-digit year comparison was made to 
                         determine the century

@author Mike
''' 
def dateRule(finalVals, value, dob, buried, cent, war, app):
    warFlag = False
    warsFlag = False
    wars = ["World War 1", "World War 2", "Korean War", "Vietnam War", "Mexican Border War"]
    temp = war.split(" and ")
    for x in temp:
        if x in wars:
            warsFlag = True
    birthYYFlag = False
    deathYYFlag = False
    bYear = ""
    dYear = ""
    buried4Year = ""
    buried2Year = ""
    cent = cent.replace(" ", "")
    appYear = ""
    if "." in app or "," in app:
        app = app.replace(".", ",")
        tempYear = app.split(",")[-1].replace(" ", "")
        while tempYear and not tempYear[-1].isnumeric():
            tempYear = tempYear[:-1]
        if len(tempYear) == 4:
            appYear = tempYear
    elif app:
        tempYear = app
        while tempYear and not tempYear[-1].isnumeric():
            tempYear = tempYear[:-1]
        tempYear = tempYear.split(" ")[-1]
        if len(tempYear) == 4:
            appYear = tempYear
    birth = dob
    if birth:
        if birth.count("/") == 2:
            birth = birth.replace(".", "")
        birth = birth.replace(":", " ").replace("I", "1").replace(".", " ")\
            .replace("&", "").replace("x", "").replace("\n", " ").replace(";", " ")\
            .replace("_", "").replace("X", "")
        if "at" in birth.lower():
            birth = birth.lower().split("at")[0]
        while birth and not birth[-1].isnumeric():
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
        if "born" in birth.lower():
            birth = birth.lower().split("born")[1]
        if len(birth.replace(" ", "")) == 4:
            bYear = birth.replace(" ", "")
            birth = ""
        if "/" in birth:
            birth = birth.replace(" ", "")
    # temp = birth[:3].replace('7', '/')
    # birth = temp + birth[3:]
    death = value
    if death:
        if death.count("/") == 2:
            death = death.replace(".", "")
        death = death.replace(":", " ").replace("I", "1").replace(".", " ")\
            .replace("&", "").replace("x", "").replace("\n", " ").replace(";", " ")\
            .replace("_", "").replace("X", "")
        if death[-1] == " ":
            death = death[:-1]
        if death[-1] == "/":
            death = death[:-1]
        if "-" in birth:
            birth = birth.replace("-", "/")
        if "-" in death:
            death = death.replace("-", "/")
        if death[0] == "/":
            death = death[1:]
        death = death.replace("Found", "").replace("on", "")
        if len(death.replace(" ", "")) == 4:
            dYear = death.replace(" ", "")
            death = ""
        if "death" in death.lower():
            death = death.lower().split("death")[1]
        while death and not death[-1].isnumeric():
            death = death[:-1]
        if "/" in death:
            death = death.replace(" ", "")
    while buried and not buried[-1].isnumeric():
        buried = buried[:-1]
    # temp2 = death[:3].replace('7', '/')
    # death = temp2 + death[3:]
    try:
        buried4Year, buried2Year = buriedRule(buried.replace("_", ""), cent, warsFlag)
        if buried4Year.count("19") == 2:
            buried4Year = buried4Year.split("19")
            if all(item == "" for item in buried4Year):
                buried4Year = "1919"
            else:
                for x in buried4Year:
                    if x != "":
                        buried4Year = "19" + x
        else:
            if buried4Year[2:] == "19":
                buried4Year = buried4Year[2:] + buried4Year[:2]
            else:
                buried4Year = buried4Year
    except Exception:
        print("Buried didn't parse")
    if birth != "" and death != "":
        if dateparser.parse(birth, settings={'STRICT_PARSING': True}) != None:
            if dateparser.parse(death, settings={'STRICT_PARSING': True}) != None:
                birth, bYear, birthYYFlag = parseBirth(birth, bYear, birthYYFlag)
                death, dYear, deathYYFlag = parseDeath(death, dYear, deathYYFlag)
                if birthYYFlag and deathYYFlag:
                    if cent:
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
                    elif buried4Year:
                        if buried4Year[:2] == "17" and buried4Year[2:] > bYear:
                            bYear = "17" + bYear
                            dYear = buried4Year
                        elif buried4Year[:2] == "18" and buried4Year[2:] < bYear:
                            bYear = "17" + bYear
                            dYear = buried4Year
                        elif buried4Year[:2] == "18" and buried4Year[2:] > bYear:
                            bYear = "18" + bYear
                            dYear = buried4Year
                        elif buried4Year[:2] == "19" and buried4Year[2:] < bYear:
                            bYear = "18" + bYear
                            dYear = buried4Year
                        elif buried4Year[:2] == "19" and buried4Year[2:] > bYear:
                            bYear = "19" + bYear
                            dYear = buried4Year
                        elif buried4Year[:2] == "20" and buried4Year[2:] < bYear:
                            bYear = "19" + bYear
                            dYear = buried4Year
                        else:
                            finalVals.append("")
                            finalVals.append("")
                            finalVals.append("")
                            finalVals.append("")
                            return warFlag
                        birth = birth[:-2] + bYear
                        death = death[:-2] + dYear
                        finalVals.append(birth)
                        finalVals.append(int(bYear))
                        finalVals.append(death)
                        finalVals.append(int(dYear))
                    elif appYear:
                        if appYear[2:] < dYear:
                            if appYear[:2] == "20":
                                dYear = "19" + dYear
                            elif appYear[:2] == "19":
                                dYear = "18" + dYear
                            elif appYear[:2] == "18":
                                dYear = "17" + dYear
                        elif appYear[2:] >= dYear:
                            if appYear[:2] == "20":
                                dYear = "20" + dYear
                            elif appYear[:2] == "19":
                                dYear = "19" + dYear
                            elif appYear[:2] == "18":
                                dYear = "18" + dYear
                            elif appYear[:2] == "17":
                                dYear = "17" + dYear
                        if dYear[2:] < bYear:
                            if dYear[:2] == "20":
                                bYear = "19" + bYear
                            elif dYear[:2] == "19":
                                bYear = "18" + bYear
                            elif dYear[:2] == "18":
                                bYear = "17" + bYear
                            elif dYear[:2] == "17":
                                bYear = "16" + bYear
                        elif dYear[2:] > bYear:
                            if dYear[:2] == "20":
                                bYear = "20" + bYear
                            elif dYear[:2] == "19":
                                bYear = "19" + bYear
                            elif dYear[:2] == "18":
                                bYear = "18" + bYear
                            elif dYear[:2] == "17":
                                bYear = "17" + bYear
                        birth = birth[:-2] + bYear
                        death = death[:-2] + dYear
                        finalVals.append(birth)
                        finalVals.append(int(bYear))
                        finalVals.append(death)
                        finalVals.append(int(dYear))
                    elif warsFlag:
                        if dYear[0] == "0":
                            bYear = "19" + bYear
                            birth = birth[:-2] + bYear
                            dYear = "20" + dYear
                            death = death[:-2] + dYear
                            finalVals.append(birth)
                            finalVals.append(int(bYear))
                            finalVals.append(death)
                            finalVals.append(int(dYear))
                        elif bYear > dYear:
                            bYear = "18" + bYear
                            birth = birth[:-2] + bYear
                            dYear = "19" + dYear
                            death = death[:-2] + dYear
                            finalVals.append(birth)
                            finalVals.append(int(bYear))
                            finalVals.append(death)
                            finalVals.append(int(dYear))
                        elif bYear < dYear:
                            bYear = "19" + bYear
                            birth = birth[:-2] + bYear
                            dYear = "19" + dYear
                            death = death[:-2] + dYear
                            finalVals.append(birth)
                            finalVals.append(int(bYear))
                            finalVals.append(death)
                            finalVals.append(int(dYear))
                        elif war == "Spanish American War" and dYear < "98":
                            if bYear > dYear:
                                bYear = "18" + bYear
                                birth = birth[:-2] + bYear
                                dYear = "19" + dYear
                                death = death[:-2] + dYear
                                finalVals.append(birth)
                                finalVals.append(int(bYear)) 
                                finalVals.append(death)
                                finalVals.append(int(dYear))
                            elif bYear < dYear:
                                bYear = "19" + bYear
                                birth = birth[:-2] + bYear
                                dYear = "19" + dYear
                                death = death[:-2] + dYear
                                finalVals.append(birth)
                                finalVals.append(int(bYear)) 
                                finalVals.append(death)
                                finalVals.append(int(dYear))
                        else:
                            finalVals.append("")
                            finalVals.append("") 
                            finalVals.append("")
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
                    elif bYear[:2] == "17" and bYear[-2:] < dYear:
                        dYear = "17" + dYear
                    elif bYear[:2] == "17" and bYear[-2:] > dYear: 
                        dYear = "18" + dYear
                    death = death[:-2] + dYear
                    if birth:
                        finalVals.append(birth)
                    else:
                        finalVals.append("")
                    if bYear:
                        finalVals.append(int(bYear))
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
                    if dYear[:2] == "19" and dYear[2:] < bYear:
                        bYear = "18" + bYear
                    elif dYear[:2] == "19" and dYear[2:] > bYear:
                        bYear = "19" + bYear
                    elif dYear[:2] == "18" and dYear[2:] < bYear:
                        bYear = "19" + bYear
                    elif dYear[:2] == "18" and dYear[2:] > bYear: 
                        bYear = "18" + bYear
                    elif dYear[:2] == "20":
                        bYear = "19" + bYear
                    elif dYear[:2] == "17" and dYear[2:] > dYear:
                        bYear = "17" + bYear
                    elif dYear[:2] == "17" and dYear[2:] < dYear:
                        bYear = "18" + bYear
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
                if death:
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
                        if bYear[:2] == "19" and bYear[-2:] > dYear:
                            dYear = "20" + dYear
                        elif bYear[:2] == "19" and bYear[-2:] < dYear:
                            dYear = "19" + dYear
                        elif bYear[:2] == "18" and bYear[-2:] > dYear:
                            dYear = "19" + dYear
                        elif bYear[:2] == "18" and bYear[-2:] < dYear: 
                            dYear = "18" + dYear
                        elif bYear[:2] == "17" and bYear[-2:] > dYear:
                            dYear = "18" + dYear
                        elif bYear[:2] == "17" and bYear[-2:] < dYear:
                            dYear = "17" + dYear
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
                elif dYear:
                    finalVals.append(birth)
                    if bYear:
                        finalVals.append(int(bYear))
                    else:
                        finalVals.append("")
                    finalVals.append("")
                    finalVals.append(int(dYear))
                    return warFlag
                elif cent:
                    if cent[2:] > bYear[2:]:
                        bYear = "19" + bYear[2:]
                    elif cent[2:] < bYear[2:]:
                        bYear = "18" + bYear[2:]
                    birth = birth[:-4] + bYear
                    finalVals.append(birth)
                    finalVals.append(int(bYear))
                    finalVals.append("")
                    finalVals.append("")
                    return warFlag
                elif appYear:
                    if bYear[2:] < appYear[2:]:
                        bYear = "19" + bYear[2:]
                    else:
                        bYear = "20" + bYear[2:]
                    birth = birth[:-4] + bYear
                    finalVals.append(birth)
                    finalVals.append(int(bYear))
                    finalVals.append("")
                    finalVals.append(int(appYear))
                else:
                    finalVals.append(birth)
                    if bYear:
                        finalVals.append(int(bYear))
                    else:
                        finalVals.append("")
                    finalVals.append("")
                    finalVals.append("")
                    return warFlag
        else:
            finalVals.append("")
            finalVals.append("")
            finalVals.append("")
            finalVals.append("")
            return warFlag
        # elif dateparser.parse(death, settings={'STRICT_PARSING': True}) != None:
        #     finalVals.append("")
        #     finalVals.append("")
        #     death = dateparser.parse(death)
        #     finalVals.append(death.strftime("%m/%d/%Y"))
        #     finalVals.append(int(death.strftime("%Y")))
        #     return warFlag
        # else:
        #     birth, bYear, birthYYFlag = parseBirth(birth, bYear, birthYYFlag)
        #     death, dYear, deathYYFlag = parseDeath(death, dYear, deathYYFlag)
        #     if birthYYFlag and deathYYFlag:
        #         if cent:
        #             if bYear > dYear:
        #                 bYear = "18" + bYear
        #                 dYear = "19" + dYear
        #             elif bYear < dYear:
        #                 bYear = "19" + bYear
        #                 dYear = "19" + dYear
        #             birth = birth[:-2] + bYear
        #             death = death[:-2] + dYear
        #             finalVals.append(birth)
        #             finalVals.append(int(bYear))
        #             finalVals.append(death)
        #             finalVals.append(int(dYear))
        #             return warFlag
        #         elif dYear < bYear:
        #             finalVals.append("")
        #             finalVals.append(int("18" + bYear))
        #             finalVals.append("")
        #             finalVals.append(int("19" + dYear))
        #             warFlag = True
        #             return warFlag
        #         elif dYear > bYear:
        #             finalVals.append("")
        #             finalVals.append(int("19" + bYear))
        #             finalVals.append("")
        #             finalVals.append(int("19" + dYear))
        #             warFlag = True
        #             return warFlag
        #     elif birthYYFlag and not deathYYFlag:
        #         if dYear[:2] == "19" and dYear[-2:] < bYear:
        #             bYear = "18" + bYear
        #         elif dYear[:2] == "19" and dYear[-2:] > bYear:
        #             bYear = "19" + bYear
        #         elif dYear[:2] == "18" and dYear[-2:] < bYear:
        #             bYear = "19" + bYear
        #         elif dYear[:2] == "18" and dYear[-2:] > bYear: 
        #             bYear = "18" + bYear
        #         elif dYear[:2] == "20":
        #             bYear = "19" + bYear
        #         elif dYear[:2] == "17" and dYear[-2:] < bYear:
        #             bYear = "18" + bYear
        #         elif dYear[:2] == "17" and dYear[-2:] > bYear:
        #             bYear = "17" + bYear
        #         birth = birth[:-2] + bYear
        #         finalVals.append(birth)
        #         if bYear: 
        #             finalVals.append(int(bYear))
        #         else:
        #             finalVals.append("")
        #         finalVals.append(death)
        #         finalVals.append(int(dYear))
        #         return warFlag
        #     elif not birthYYFlag and deathYYFlag:
        #         if bYear[:2] == "19" and bYear[-2:] < dYear:
        #             dYear = "20" + dYear
        #         elif bYear[:2] == "19" and bYear[-2:] > dYear:
        #             dYear = "19" + dYear
        #         elif bYear[:2] == "18" and bYear[-2:] < dYear:
        #             dYear = "19" + dYear
        #         elif bYear[:2] == "18" and bYear[-2:] > dYear: 
        #             dYear = "18" + dYear
        #         elif bYear[:2] == "17" and bYear[-2:] < dYear:
        #             dYear = "18" + dYear
        #         elif bYear[:2] == "17" and bYear[-2:] > dYear:
        #             dYear = "17" + dYear
        #         death = death[:-2] + dYear
        #         finalVals.append(birth)
        #         finalVals.append(int(bYear))
        #         finalVals.append(death)
        #         finalVals.append(int(dYear))
        #         return warFlag
        #     else:
        #         finalVals.append("")
        #         finalVals.append("") 
        #         finalVals.append("")
        #         finalVals.append("") 
        #         return warFlag
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
                if bYear[:2] == "17" and bYear[-2:] < dYear:
                    dYear = "17" + dYear
                elif bYear[:2] == "17" and bYear[-2:] > dYear:
                    dYear = "18" + dYear
                elif bYear[:2] == "18" and bYear[-2:] < dYear:
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
            elif cent:
                finalVals.append("")
                finalVals.append("") 
                if len(cent) == 4:
                    dYear = cent
                elif len(cent) == 2:
                    dYear = "19" + dYear
                finalVals.append(death[:-2] + dYear)
                finalVals.append(int(dYear))
            elif warsFlag or (war == "Spanish American War" and dYear[:2] < "98"):
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
            if bYear:
                finalVals.append(birth)
                finalVals.append(int(bYear))
            else:
                finalVals.append("")
                finalVals.append("") 
            if dYear != "":
                finalVals.append("")
                finalVals.append(int(dYear))
                return warFlag
            elif buried4Year != "":
                dYear = buried4Year
                finalVals.append("")
                finalVals.append(int(dYear))
                return warFlag
            elif buried2Year != "":
                if bYear[:2] == "17" and bYear[2:] < buried2Year:
                    dYear = "17" + buried2Year
                elif bYear[:2] == "17" and bYear[2:] > buried2Year:
                    dYear = "18" + buried2Year
                elif bYear[:2] == "18" and bYear[2:] < buried2Year:
                    dYear = "18" + buried2Year
                elif bYear[:2] == "18" and bYear[2:] > buried2Year:
                    dYear = "19" + buried2Year
                elif bYear[:2] == "19" and bYear[2:] < buried2Year:
                    dYear = "19" + buried2Year
                elif bYear[:2] == "19" and bYear[2:] > buried2Year:
                    dYear = "20" + buried2Year
                finalVals.append("")
                finalVals.append(int(dYear))
                return warFlag
            elif cent:
                finalVals.append("")
                finalVals.append("") 
                finalVals.append("") 
                finalVals.append(int(cent))
            else:
                finalVals.append("")
                finalVals.append("") 
        elif buried4Year:
            if buried4Year[:2] == "19" and buried4Year[2:] > bYear:
                bYear = "19" + bYear
            elif buried4Year[:2] == "19" and buried4Year[2:] < bYear:
                bYear = "18" + bYear
            elif buried4Year[:2] == "20" and buried4Year[2:] < bYear:
                bYear = "19" + bYear
            elif buried4Year[:2] == "18" and buried4Year[2:] > bYear:
                bYear = "18" + bYear
            elif buried4Year[:2] == "18" and buried4Year[2:] < bYear:
                bYear = "17" + bYear
            birth = birth[:-2] + bYear
            finalVals.append(birth)
            finalVals.append(int(bYear))
            finalVals.append("")
            finalVals.append(int(buried4Year))
        else:
            if dYear:
                if dYear[:2] == "17" and dYear[-2:] > bYear:
                    bYear = "17" + bYear
                elif dYear[:2] == "18" and dYear[-2:] < bYear:
                    bYear = "17" + bYear
                elif dYear[:2] == "18" and dYear[-2:] > bYear:
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
        if dYear:
            pass
        elif buried4Year:
            dYear = buried4Year
        elif cent:
            if len(cent) == 4:
                dYear = cent
            elif len(cent) == 2:
                dYear = "19" + cent
        elif buried2Year:
            if war == "Spanish American War" and buried2Year < "98":
                dYear = "19" + buried2Year
            if warsFlag:
                if buried2Year > "14":
                    dYear = "19" + buried2Year
                else:
                    dYear = "20" + buried2Year
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