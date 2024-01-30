import dateparser
import re

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
            try:
                death = death.strftime("%m/%d/%Y")
            except AttributeError:
                dYear = ""
                death = ""
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


'''
Analyzes and formats date-related information based on various rules and conditions. 
It reconciles birth and death dates with other date-related information. Handles conditions
where missing birth and/or death dates and/or years. Attempts to correct 2-digit years
using cent date, else uses buried date, else uses war info, else uses "<" comparison and 
this comparison calls for the activation of warFlag signaling to highlight the row due to 
possible wrong century.

@param finalVals (list) - List where the processed date information will be appended
@param value (string) - Current value being processed (typically related to death date)
@param dob (string) - Date of birth
@param buried (string) - Date of burial
@param buriedYear (string) - Year of burial
@param cent (string) -  Century of the event (used for 2-digit years)
@param war (string) - War during which the person served, if applicable

@return warFlag (bool) - Flag indicating if a 2-digit year comparison was made to 
                         determine the century

@author Mike
''' 
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
    elif "at" in birth:
        birth = birth.split("at")
        birth = birth[0] 
    while birth and not birth[-1].isalnum():
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
                    elif dYear[:2] == "17" and dYear[-2:] > dYear:
                        bYear = "17" + bYear
                    elif dYear[:2] == "17" and dYear[-2:] < dYear:
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
                    elif bYear[:2] == "17" and bYear[-2:] < dYear:
                        dYear = "18" + dYear
                    elif bYear[:2] == "17" and bYear[-2:] > dYear:
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
                elif dYear[:2] == "17" and dYear[-2:] < bYear:
                    bYear = "18" + bYear
                elif dYear[:2] == "17" and dYear[-2:] > bYear:
                    bYear = "17" + bYear
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
                elif bYear[:2] == "17" and bYear[-2:] < dYear:
                    dYear = "18" + dYear
                elif bYear[:2] == "17" and bYear[-2:] > dYear:
                    dYear = "17" + dYear
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


'''
Processes and formats the burial year information. It adjusts the year based on the 
century and war involvement, and reconciles it with other date-related information.

@param value (str) - Extracted raw burial date
@param bYear (str) - Burial year, particularly if only the year is printed on the card
@param cent (str) - If not none, provides century data for death year
@param warsFlag (bool) - Flag indicating involvement in a 19th century war

@return buriedYear (str) - Processed and formatted burial year

@author Mike
'''
def buriedRule(value, bYear, cent, warsFlag):
    buried = value
    buried = buried.replace("I", "1").replace("-", "/")
    buriedYear = bYear
    if len(cent) == 2:
        buriedYear = "19" + cent
    elif len(cent) == 4:
        buriedYear = cent
    elif len(cent) == 3:
        buriedYear = "19" + cent[-2:]
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
                    print("buried is weird\n")
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

finalVals = []
birth = "1/2/1713"
death = "1/3/03"
cent = ""
buried = "2/4/03"
buriedYear = ""
war = ""
dateRule(finalVals, death, birth, buried, buriedYear, cent, war)
for x in finalVals:
    print(x)