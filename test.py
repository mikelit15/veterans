import re

value = "195919"
value = value.replace(" ", "").replace(".", "")
if value.count("19") == 2:
    value = value.split("19")
    if all(item == "" for item in value):
        value = "1919"
    else:
        for x in value:
            if x != "":
                value = "19" + x
print(value)