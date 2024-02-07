import re

def warRule(value, world):
    ww1years = ["1914", "1915", "1916", "1917", "1918"]
    war = value.strip()
    identified_wars = []  
    war = war.replace("N.", "W").replace("N", "W").replace(" ", "").replace("-", "").replace(".", "").replace("#", "").replace(",", "")
    war = re.sub(r'\bT\b', '2', war)  
    war = re.sub(r'\bTT\b', '2', war)  
    ww1_pattern = re.compile(r'WW1|WWI|WWl|\b1\b|W\.?W\.?\s?(1|I|l|T|One|ONE)|World\s*War\s*(1|I|l|T|One|ONE)|WorldWar\s*(1|I|l|T|One|ONE)', re.IGNORECASE)
    ww2_pattern = re.compile(r'WW2|WWII|WWll|WW11|\b2\b|W\.?W\.?\s?(2|II|ll|TT|Two|TWO)|World\s*War\s*(2|II|ll|TT|Two|TWO)|WorldWar\s*(2|II|ll|TT|Two|TWO)', re.IGNORECASE)    
    ww1_and_ww2_combined_pattern = re.compile(r'(WW\s*I\s*&\s*II|World\s*War\s*I\s*&\s*II|World\s*War\s*1\s*and\s*2)', re.IGNORECASE)
    korean_war_pattern = re.compile(r'Korea', re.IGNORECASE)
    vietnam_war_pattern = re.compile(r'Vietnam', re.IGNORECASE)
    simple_world_war_pattern = re.compile(r'\bWorld\s*War\b', re.IGNORECASE) 

    if ww1_and_ww2_combined_pattern.search(war):
        identified_wars.extend(["World War 1", "World War 2"])
    elif simple_world_war_pattern.search(war):
        identified_wars.append("World War 1")
    else:
        if ww2_pattern.search(war):
            identified_wars.append("World War 2")
        if not identified_wars and ww1_pattern.search(war):
            identified_wars.append("World War 1")
    if korean_war_pattern.search(war):
        identified_wars.append("Korean War")
    if vietnam_war_pattern.search(war):
        identified_wars.append("Vietnam War")
    if identified_wars:
        war = ' and '.join(identified_wars)
    if war and not identified_wars:
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
    if not war and world:
        war = world
    if "Army" in war or "US" in war:
        war = ""
    return war

# "WW. #1"
# World Wars I & II
# .WW. IT
# WW IT
# WW TT


value = "WW 1 & 2"
world = ""
war = ""
# if "&" in value:
#     wars = value.split("&")
#     war1 = warRule(wars[0], world)
#     war2 = warRule(wars[1], world)
#     if war2:
#         war = war1 + " and " + war2
#     else:
#         war = war1
# elif "and" in value.lower():
#     wars = value.split("and")
#     war1 = warRule(wars[0], world)
#     war2 = warRule(wars[1], world)
#     if war2:
#         war = war1 + " and " + war2
#     else:
#         war = war1
# elif "-" in value:
#     wars = value.split("-")
#     war1 = warRule(wars[0], world)
#     war2 = warRule(wars[1], world)
#     if war2:
#         war = war1 + " and " + war2
#     else:
#         war = war1
# else:
war = warRule(value, world)
war = war.strip().strip('and').strip()

print(war)