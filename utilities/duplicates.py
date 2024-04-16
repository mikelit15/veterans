import pandas as pd


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
def main():
    try:
        df = pd.read_excel(r"\\ucclerk\pgmdoc\Veterans\Veterans.xlsx", sheet_name="Graceland", usecols="A:K")
        df.reset_index(inplace=True, drop=True)  
        df.index = df.index + 2
        potential_duplicates = df[df.duplicated(subset=['VLNAME', 'VFNAME', 'VDODY'], keep=False)]
        confirmed_duplicates = pd.DataFrame()
        for _, group in potential_duplicates.groupby(['VLNAME', 'VFNAME', 'VDODY']):
            if not group['VDOBY'].isna().any():
                confirmed_group = group[group.duplicated(subset=['VLNAME', 'VFNAME', 'VDODY', 'VDOBY'], keep=False)]
                confirmed_duplicates = pd.concat([confirmed_duplicates, confirmed_group])
            else:
                confirmed_duplicates = pd.concat([confirmed_duplicates, group])
        confirmed_duplicates['MaxVIDinPair'] = confirmed_duplicates.groupby(['VLNAME', 'VFNAME', 'VDODY', 'VDOBY'])['VID'].transform('max')
        confirmed_duplicates.sort_values(by=['MaxVIDinPair', 'VID'], inplace=True)
        confirmed_duplicates.drop(columns=['MaxVIDinPair'], inplace=True)
        if df.empty:
            print("### No Duplicates ###")
        else:
            print(confirmed_duplicates)
    except Exception:
        print("### No Duplicates ###")


if __name__ == "__main__":
    main()