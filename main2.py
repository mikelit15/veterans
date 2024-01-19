import os
import openpyxl
import re

network_folder = r"\\ucclerk\pgmdoc\Veterans\Cemetery"
cemeterys = os.listdir(network_folder)
miscPath = os.path.join(network_folder, "Misc")
miscs = os.listdir(miscPath)
print(miscs)