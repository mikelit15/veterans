import pandas as pd

columns = "B:N"
df = pd.read_excel(r"\\ucclerk\pgmdoc\Veterans\Veterans.xlsx", usecols=columns)
duplicates = df[df.duplicated(keep=False)]
print("Duplicate Rows:")
print(duplicates)
print(f"Count of duplicate rows: {duplicates.shape[0]}")
