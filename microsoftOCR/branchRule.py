'''
Determines and appends the military branch based on the extracted text. It handles various 
abbreviations and naming conventions.

@param finalVals (list) - A list where the processed military branch will be appended
@param value (str) - The raw branch name
@param war (str) - A net for if war name is caught by the branch key, or misinput on the
                   registration card
                   
@author Mike

'''
def branchRule(finalVals, value, war):
    value = value.replace("\n", " ")
    if value.count('.') >= 3:
        value = value.replace(" ", "")
    value = value.replace("USA", "")
    armys = ["co", "army", "inf", "infantry", "infan", "med", "cav", "div", \
             "sig", "art", "corps", "corp", "artillery", "army", "qmc", "q m c", \
             "ind"]
    navys = ["hospital", "navy", "naval", "avy", "usnr", "uss", "usn", "usnrf", \
             "u s n r", "u s s", "u s n", "u s n r f" ,"merchant"]
    guards = ["113", "102d", "114", "44", "181", "250", "112", "national"]
    branch = value
    branch = branch.replace("/", " ").replace(".", " ").replace("th", "").replace("-", "")\
             .replace(", ", " ")
    words = branch.split()
    if war in value and war != "":
        branch = ""
        value = ""
    if branch[-1:] == " ":
        branch = branch[:-1]
    if "air force" in branch.lower() or "air corp" in branch.lower() or "u s a f" in branch.lower():
        branch = "Air Force"
    elif "marine" in branch.lower():
        branch = "Marine Corps"
    elif branch.lower() in armys:
        branch = "Army"
    elif branch.lower() in navys:
        branch = "Navy"
    elif "u s m c " in branch.lower() or "usmc" in branch.lower():
        branch = "Marine Corps" 
    elif branch.lower() in guards:
        branch = "National Guard"
    elif "u s c g " in branch.lower():
        branch = "Coast Guard"
    else:
        for word in words:
            word = word.lower()
            if word in armys:
                branch = "Army"
                break
            elif "air" in word or "aero" in word or "aaf" in word:
                branch = "Air Force"
                break
            elif word in navys:
                branch = "Navy"
                break
            elif "coast" in word or "uscg" in word:
                branch = "Coast Guard"
                break
            elif word in guards:
                branch = "National Guard"
                break
            else:
                flag1 = False
                flag2 = False
                for x in armys:
                    if x in word:
                        branch = "Army"
                        flag1 = True
                        break
                if flag1:
                    break
                for x in navys:
                    if x in word:
                        branch = "Navy"
                        flag2 = True
                        break
                if flag2:
                    break
                branch = ""
    if value == "N/A":
        value = ""
    finalVals.append(value)   
    finalVals.append(branch)   