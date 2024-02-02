import re

def warRule(value, world):
    ww1years = ["1914", "1915", "1916", "1917", "1918"]
    war = value.strip()
    identified_wars = []  
    war = war.replace("N.", "W").replace("N", "W").replace(" ", "").replace("-", "").replace(".", "").replace("#", "").replace(",", "")
    war = re.sub(r'\bT\b', '2', war)  
    war = re.sub(r'\bTT\b', '2', war)  
    ww1_pattern = re.compile(r'WW1|WWI|WWl|\b1\b|W\.?W\.?\s?(1|I|l|One|ONE)|World\s*War\s*(1|I|l|One|ONE)|WorldWar\s*(1|I|l|One|ONE)', re.IGNORECASE)
    ww2_pattern = re.compile(r'WW2|WWII|WWll|WW11|\b2\b|W\.?W\.?\s?(2|II|ll|Two|TWO)|World\s*War\s*(2|II|ll|Two|TWO)|WorldWar\s*(2|II|ll|Two|TWO)', re.IGNORECASE)
    korean_war_pattern = re.compile(r'Korea', re.IGNORECASE)
    vietnam_war_pattern = re.compile(r'Vietnam', re.IGNORECASE)
    if ww2_pattern.search(war):
        identified_wars.append("World War 2")
    if not identified_wars and ww1_pattern.search(war):
        identified_wars.append("World War 1")
    if korean_war_pattern.search(war):
        identified_wars.append("Korean War")
    if vietnam_war_pattern.search(war):
        identified_wars.append("Vietnam War")
    if not war and not identified_wars:
        war = war.replace(".", "").replace("Calv", "").replace("Vols", "")
        if "Korea" in war:
            war = "Korean War"
        elif "Vietnam" in war:
            war = "Vietnam War"
        elif "Civil" in war or "Citil" in war or "Gettysburg" in war:
            war = "Civil War"
        elif "Spanish" in war or "Amer" in war or "American" in war:
            war = "Spanish American War"
        elif "Mexican" in war:
            war = "Mexican Border War"
        elif "Rebellion" in war:
            war = "War of the Rebellion"
        elif "Revolution" in war:
            war = "Revolutionary War"
        elif "1812" in war:
            war = "War of 1812"
        else:
            war = ""
        words = war.split()
        for x in words:
            if x in ww1years:
                war = "World War 1"
                break
    war = ' and '.join(identified_wars)
    if not war and world:
        war = world
    if "Army" in war or "US" in war:
        war = ""
    return war

value = "WWTT Korea"
print(warRule(value, ""))