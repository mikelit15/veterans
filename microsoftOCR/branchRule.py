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
             "ind", "depot", "312", "usar", "ord", "engrs", "eng", "ordnance", \
             "detachment", "training regt", "field artillery", "btry", "service unit"]
    navys = ["hospital", "navy", "naval", "avy", "usnr", "uss", "usn", "usnrf", \
             "u s n r", "u s s", "u s n", "u s n r f", "merchant"]
    guards = ["113", "102d", "114", "44", "181", "250", "112", "national", "us guards", \
              "47th bn"]
    marines = ["marine", "usmc", "u s m c"]
    air_forces = ["air force", "air corp", "u s a f", "u s a a f", "aaf", "af"]
    coast_guard = ["coast guard", "uscg", "u s c g"]

    branch = value
    branch = branch.replace("/", " ").replace(".", " ").replace("th", "").replace("-", "")\
             .replace(", ", " ")
    words = branch.split()
    if war in value and war != "":
        branch = ""
        value = ""
    if branch[-1:] == " ":
        branch = branch[:-1]

    if any(af in branch.lower() for af in air_forces):
        branch = "Air Force"
    elif any(mar in branch.lower() for mar in marines):
        branch = "Marine Corps"
    elif branch.lower() in armys:
        branch = "Army"
    elif branch.lower() in navys:
        branch = "Navy"
    elif branch.lower() in guards:
        branch = "National Guard"
    elif any(cg in branch.lower() for cg in coast_guard):
        branch = "Coast Guard"
    else:
        for word in words:
            word = word.lower()
            if word in armys:
                branch = "Army"
                break
            elif any(af in word for af in air_forces):
                branch = "Air Force"
                break
            elif word in navys:
                branch = "Navy"
                break
            elif any(cg in word for cg in coast_guard):
                branch = "Coast Guard"
                break
            elif word in guards:
                branch = "National Guard"
                break
            elif any(mar in word for mar in marines):
                branch = "Marine Corps"
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
    if value == "N/A" or value == "Not Listed":
        value = ""
    finalVals.append(value)   
    finalVals.append(branch)