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

@param sheetName (str) - The name of the sheet in the Excel file to be processed
@param excelFile (str) - The path to the Excel file to be loaded

@return df (DataFrame) - Processed DataFrame with relevant columns loaded and adjusted

@author Mike
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

'''
Identifies duplicate records based on specific fields such as last name, first name, 
and date of death year. Handles confirmed duplicates and organizes them for review.

@param df (DataFrame) - The DataFrame to search for duplicates
@param excludeList (list) - A list of IDs to exclude from the duplicate check

@return confirmedDuplicates (DataFrame) - DataFrame containing confirmed duplicates

@author Mike
'''
def findDuplicates(df, excludeList):
    df = df[~df['VID'].isin(excludeList)]
    confirmedDuplicates = pd.DataFrame()
    potentialDuplicates = df[df.duplicated(subset=['VLNAME', 'VFNAME', 'VDODY'], keep=False)]
    for _, group in potentialDuplicates.groupby(['VLNAME', 'VFNAME', 'VDODY']):
        if not group['VDOBY'].isna().any():
            confirmedGroup = group[group.duplicated(subset=['VLNAME', 'VFNAME', 'VDODY', 'VDOBY'], keep=False)]
            confirmedDuplicates = pd.concat([confirmedDuplicates, confirmedGroup])
        else:
            confirmedDuplicates = pd.concat([confirmedDuplicates, group])
    confirmedDuplicates['MaxVIDinPair'] = confirmedDuplicates.groupby(['VLNAME', 'VFNAME', 'VDODY', 'VDOBY'])['VID'].transform('max')
    confirmedDuplicates.sort_values(by=['VLNAME', 'VFNAME', 'VDODY', 'MaxVIDinPair', 'VID'], inplace=True)
    confirmedDuplicates.drop(columns=['MaxVIDinPair'], inplace=True)
    return confirmedDuplicates


'''
Main function that processes the Excel file, identifies duplicates, and saves the results.
Iterates through all sheets in the Excel file and checks for duplicates across the entire dataset.

@return confirmedDuplicates (DataFrame) - DataFrame containing all confirmed duplicates 
                                          across the sheets processed
                                                                      
@author Mike
'''
def main():
    filePath = r"\\ucclerk\pgmdoc\Veterans\Veterans.xlsx"
    outputFilePath = r"\\ucclerk\pgmdoc\Veterans\confirmed_duplicates.txt"
    combinedDF = pd.DataFrame()
    confirmedDuplicates = pd.DataFrame()
    
    excludeList = []
    
    try:
        excelFile = pd.ExcelFile(filePath)
        for sheetName in excelFile.sheet_names:
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
            with open(outputFilePath, 'w') as file:
                confirmedDuplicates.to_string(file, index=False)
    except Exception as e:
        print("### No Duplicates ###")
        print(f"Error: {e}")
    return confirmedDuplicates

if __name__ == "__main__":
    main()