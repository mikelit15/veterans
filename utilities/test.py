import re
import duplicates
import pandas as pd

columnWidths = {
    "VID": 10,
    "VLNAME": 20,
    "VFNAME": 20,
    "VDOB": 20,
    "VDOBY": 10,
    "VDOD": 20,
    "VDODY": 10,
    "VWARREC": 20,
    "SHEET NAME": 10
}
# df = duplicates.main()
# for index, row in df.iterrows():
#     for col in df.columns:
#         print(f"Row {index}, Column {col}: {row[col]}")

d = {'col1': [1], 'col2': [4, 5, 6]}
max_length = max(len(d['col1']), len(d['col2']))
d['col1'].extend([None] * (max_length - len(d['col1'])))
df = pd.DataFrame(d)
print(df.to_string())