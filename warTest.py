import re


def warRule(value, world):
    value = value.replace("\n", " ").replace("7", "1").replace("J", "1")
    ww1years = ["1914", "1915", "1916", "1917", "1918"]
    war = value.strip()
    identified_wars = []  
    war = war.replace("N.", "W").replace("N", "W").replace(" ", "").replace("-", " ")\
             .replace(".", "").replace("#", "").replace(",", "").replace("L", "1").replace("I", "1")
    ww1_pattern = re.compile(
        r'WW1|'  # Matches "WW1"
        r'WWI|'  # Matches "WWI"
        r'\b1\b|'  # Matches standalone "1"
        r'W\.?W\.?\s?(1|I|i|l|L|T|ONE)|'  # Matches "WW 1", "WW I", with optional periods and space
        r'WORLD\s*WAR\s*(1|I|i|l|L|T|ONE)',  # Matches "World War 1", "World War I", with optional spaces
        re.IGNORECASE
    ) 
    ww2_pattern = re.compile(
        r'WW2|'  # Matches "WW2"
        r'WWII|'  # Matches "WWII"
        r'\b2\b|'  # Matches standalone "2"
        r'W\.?W\.?\s?(2|II|ii|ll|LL|TT|TWO|11|TI|IT|IL|LI|LT|TL|T1|1T|1L|L1|1I|I1)|'  # Matches "WW 2", "WW II", with optional periods and space
        r'WORLD\s*WAR\s*(2|II|ii|ll|LL|TT|TWO|11|TI|IT|IL|LI|LT|TL|T1|1T|1L|L1|1I|I1)',  # Matches "World War 2", "World War II", with optional spaces
        re.IGNORECASE
    )
    ww1_and_ww2_pattern = re.compile(
        r'(WW|WORLD\s*WAR|WORLD\s*WARS)\s*'  # Starts with "WW" or "World War" with optional spaces
        r'(1|I|i|l|L|T|ONE)'  # First part of the pattern for World War 1
        r'(\s*(and|&)\s*|\sand\s*|\s*&\s*|)'  # Optional connector: "and", "&", with or without spaces
        r'(2|II|ii|ll|LL|TT|TWO|11|TI|IT|IL|LI|LT|TL|WW2|WWII)|'  # Second part of the pattern for World War 2, optional to allow single mentions
        r'WW1andWW11',  # Explicitly matches "WW1andWW11"
        re.IGNORECASE
    )
    korean_war_pattern = re.compile(r'Korea', re.IGNORECASE)
    vietnam_war_pattern = re.compile(r'Vietnam', re.IGNORECASE)
    simple_world_war_pattern = re.compile(r'\bWorld\s*War\b', re.IGNORECASE) 
    if ww1_and_ww2_pattern.search(war):
        identified_wars.extend(["World War 1", "World War 2"])
    else:
        if ww2_pattern.search(war):
            identified_wars.append("World War 2")
        elif simple_world_war_pattern.search(war):
            identified_wars.append("World War 1")
        elif not identified_wars and ww1_pattern.search(war):
            identified_wars.append("World War 1")
    if korean_war_pattern.search(war):
        identified_wars.append("Korean War")
    if vietnam_war_pattern.search(war):
        identified_wars.append("Vietnam War")
    war = war.replace(".", "").replace("Calv", "").replace("Vols", "")
    if "Civil" in war or "Citil" in war or "Gettysburg" in war \
        or "Fredericksburg" in war or "bull" in war:
        identified_wars.append("Civil War")
    if "Spanish" in war or "Amer" in war or "American" in war or "SpAm" in war:
        identified_wars.append("Spanish American War")
    if "Mexican" in war:
        identified_wars.append("Mexican Border War")
    if "Rebellion" in war:
        identified_wars.append("War of the Rebellion")
    if "Revolution" in war:
        identified_wars.append("Revolutionary War")
    if "1812" in war:
        identified_wars.append("War of 1812")
    words = war.split()
    for x in words:
        if x in ww1years:
            war = "World War 1"
            break
    if identified_wars:
        war = ' and '.join(identified_wars)
    else:
        war = ""
    if world != "" and war != "World War 1":
        war = "World War 1"
    if "Army" in war or "US" in war:
        war = ""
    return war

print(warRule("WWT1", ""))
