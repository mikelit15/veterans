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
def main(cemetery):
    confirmed_duplicates = pd.DataFrame()
    try:
        df = pd.read_excel(r"\\ucclerk\pgmdoc\Veterans\Veterans.xlsx", sheet_name= cemetery, usecols="A:K")
        df.reset_index(inplace=True, drop=True)  
        df.index = df.index + 2
        df['VDOBY'] = df['VDOBY'].apply(lambda x: str(int(x)) if pd.notna(x) else np.nan)
        df['VDODY'] = df['VDODY'].apply(lambda x: str(int(x)) if pd.notna(x) else np.nan)
        df['VDOBY'] = df['VDOBY'].astype(object).where(pd.notna(df['VDOBY']), np.nan)
        df['VDODY'] = df['VDODY'].astype(object).where(pd.notna(df['VDODY']), np.nan)
        potential_duplicates = df[df.duplicated(subset=['VLNAME', 'VFNAME', 'VDODY'], keep=False)]
        for _, group in potential_duplicates.groupby(['VLNAME', 'VFNAME', 'VDODY']):
            if not group['VDOBY'].isna().any():
                confirmed_group = group[group.duplicated(subset=['VLNAME', 'VFNAME', 'VDODY', 'VDOBY'], keep=False)]
                confirmed_duplicates = pd.concat([confirmed_duplicates, confirmed_group])
            else:
                confirmed_duplicates = pd.concat([confirmed_duplicates, group])
        confirmed_duplicates['MaxVIDinPair'] = confirmed_duplicates.groupby(['VLNAME', 'VFNAME', 'VDODY', 'VDOBY'])['VID'].transform('max')
        confirmed_duplicates.sort_values(by=['MaxVIDinPair', 'VID'], inplace=True)
        confirmed_duplicates.drop(columns=['MaxVIDinPair'], inplace=True)
        if confirmed_duplicates.empty:
            print("### No Duplicates ###")
        else:
            print(confirmed_duplicates)
    except Exception as e:
        print("### No Duplicates ###")
        print(f"Error: {e}")
    return confirmed_duplicates


if __name__ == "__main__":
    cemetery = "Evergreen"
    main(cemetery)