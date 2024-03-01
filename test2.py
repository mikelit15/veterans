
value = "19 42-"
if value:
    try:
        tempCent = value.replace(",", "").replace(".", "").replace(":", "").replace(";", "").replace("/", "")\
            .replace(" ", "").replace("\n", "").replace("_", "").replace("in", "").replace("...", "")
        while tempCent and not tempCent[-1].isnumeric():
            tempCent = tempCent[:-1]
        if tempCent.count("19") == 2:
            tempCent = tempCent.split("19")
            if all(item == "" for item in tempCent):
                cent = "1919"
            else:
                for x in tempCent:
                    if x != "":
                        cent = "19" + x
        else:
            if tempCent[2:] == "19":
                cent = tempCent[2:] + tempCent[:2]
            else:
                cent = tempCent
    except IndexError:
        pass
else:
    pass

print(cent)