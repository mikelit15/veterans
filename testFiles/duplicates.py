import pandas as pd
import numpy as np


'''
Loads and processes an Excel spreadsheet containing information about veterans to identify 
duplicate records. The process includes several key steps: 
    - load data from a specified sheet and columns
    - adjust index for alignment with Excel's row numbers
    - identify potential duplicates based last name, first name, and date of death year 
      fields
    - organize and display the confirmed duplicates for review

@author: Mike
'''

def loadAndProcessSheet(sheetName, excelFile):
    df = pd.read_excel(excelFile, sheet_name=sheetName, usecols="A:K")
    df.reset_index(inplace=True, drop=True)
    df.index = df.index + 2
    df['VDOBY'] = df['VDOBY'].apply(lambda x: str(int(x)) if pd.notna(x) else np.nan)
    df['VDOBY'] = df['VDOBY'].astype(object).where(pd.notna(df['VDOBY']), np.nan)
    df['VDODY'] = df['VDODY'].apply(lambda x: str(int(x)) if pd.notna(x) else np.nan)
    df['VDODY'] = df['VDODY'].astype(object).where(pd.notna(df['VDODY']), np.nan)
    df['VID'] = df['VID'].apply(lambda x: int(x) if pd.notna(x) else np.nan)
    df['VID'] = df['VID'].astype(object).where(pd.notna(df['VID']), np.nan)
    df['SHEET NAME'] = sheetName  
    return df

def findDuplicates(df, excludeList):
    df = df[~df['VID'].isin(excludeList)]
    confirmedDuplicates = pd.DataFrame()
    potentialDuplicates = df[df.duplicated(subset=['VLNAME', 'VFNAME', 'VDODY'], keep=False)]
    for _, group in potentialDuplicates.groupby(['VLNAME', 'VFNAME', 'VDODY']):
        if not group['VDOBY'].isna().any():
            confirmed_group = group[group.duplicated(subset=['VLNAME', 'VFNAME', 'VDODY', 'VDOBY'], keep=False)]
            confirmedDuplicates = pd.concat([confirmedDuplicates, confirmed_group])
        else:
            confirmedDuplicates = pd.concat([confirmedDuplicates, group])
    confirmedDuplicates['MaxVIDinPair'] = confirmedDuplicates.groupby(['VLNAME', 'VFNAME', 'VDODY', 'VDOBY'])['VID'].transform('max')
    confirmedDuplicates.sort_values(by=['VLNAME', 'VFNAME', 'VDODY', 'MaxVIDinPair', 'VID'], inplace=True)
    confirmedDuplicates.drop(columns=['MaxVIDinPair'], inplace=True)
    return confirmedDuplicates

def main():
    filePath = r"\\ucclerk\pgmdoc\Veterans\VeteransTest.xlsx"
    combinedDF = pd.DataFrame()
    confirmedDuplicates = pd.DataFrame()
    
    excludeList = [8465, 424, 9494, 2123, 4661, 4662]
    
    try:
        excelFile = pd.ExcelFile(filePath)
        for sheetName in excelFile.sheetNames:
            df = loadAndProcessSheet(sheetName, filePath)
            print(f"Processing sheet: {sheetName}")
            print("Columns:", df.columns)
            requiredColumns = {'VLNAME', 'VFNAME', 'VDODY', 'VDOBY', 'VID'}
            if not requiredColumns.issubset(df.columns):
                print(f"### Missing required columns in sheet {sheetName} ###")
                continue
            combinedDF = pd.concat([combinedDF, df], ignore_index=True)
        if not combinedDF.empty:
            confirmedDuplicates = findDuplicates(combinedDF, excludeList)
        if confirmedDuplicates.empty:
            print("### No Duplicates ###")
        else:
            print(confirmedDuplicates)
    except Exception as e:
        print("### No Duplicates ###")
        print(f"Error: {e}")
    return confirmedDuplicates

# def main(cemetery):
#     confirmedDuplicates = pd.DataFrame()
#     try:
#         df = pd.read_excel(r"\\ucclerk\pgmdoc\Veterans\Veterans.xlsx", sheetName= cemetery, usecols="A:K")
#         df.reset_index(inplace=True, drop=True)  
#         df.index = df.index + 2
#         df['VDOBY'] = df['VDOBY'].apply(lambda x: str(int(x)) if pd.notna(x) else np.nan)
#         df['VDODY'] = df['VDODY'].apply(lambda x: str(int(x)) if pd.notna(x) else np.nan)
#         df['VDOBY'] = df['VDOBY'].astype(object).where(pd.notna(df['VDOBY']), np.nan)
#         df['VDODY'] = df['VDODY'].astype(object).where(pd.notna(df['VDODY']), np.nan)
#         potentialDuplicates = df[df.duplicated(subset=['VLNAME', 'VFNAME', 'VDODY'], keep=False)]
#         for _, group in potentialDuplicates.groupby(['VLNAME', 'VFNAME', 'VDODY']):
#             if not group['VDOBY'].isna().any():
#                 confirmedGroup = group[group.duplicated(subset=['VLNAME', 'VFNAME', 'VDODY', 'VDOBY'], keep=False)]
#                 confirmedDuplicates = pd.concat([confirmedDuplicates, confirmedGroup])
#             else:
#                 confirmedDuplicates = pd.concat([confirmedDuplicates, group])
#         confirmedDuplicates['MaxVIDinPair'] = confirmedDuplicates.groupby(['VLNAME', 'VFNAME', 'VDODY', 'VDOBY'])['VID'].transform('max')
#         confirmedDuplicates.sort_values(by=['MaxVIDinPair', 'VID'], inplace=True)
#         confirmedDuplicates.drop(columns=['MaxVIDinPair'], inplace=True)
#         if confirmedDuplicates.empty:
#             print("### No Duplicates ###")
#         else:
#             print(confirmedDuplicates)
#     except Exception as e:
#         print("### No Duplicates ###")
#         print(f"Error: {e}")
#     return confirmedDuplicates


if __name__ == "__main__":
    main()