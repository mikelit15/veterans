import pandas as pd

def main():
    df = pd.read_excel(r"\\ucclerk\pgmdoc\Veterans\Veterans.xlsx", usecols="A:K")
    potential_duplicates = df[df.duplicated(subset=['VLNAME', 'VFNAME', 'VDODY'], keep=False)]
    confirmed_duplicates = pd.DataFrame()
    for _, group in potential_duplicates.groupby(['VLNAME', 'VFNAME', 'VDODY']):
        if not group['VDOBY'].isna().any():
            confirmed_group = group[group.duplicated(subset=['VLNAME', 'VFNAME', 'VDODY', 'VDOBY'], keep=False)]
            confirmed_duplicates = pd.concat([confirmed_duplicates, confirmed_group])
        else:
            confirmed_duplicates = pd.concat([confirmed_duplicates, group])
    confirmed_duplicates = confirmed_duplicates.set_index(df.columns[0])
    confirmed_duplicates.sort_values(by=df.columns[0], inplace=True)
    unique_first_ids = confirmed_duplicates.drop_duplicates(subset=['VLNAME', 'VFNAME', 'VDODY'], keep='first')
    print("Confirmed Duplicate Rows:\n")
    print(f"{confirmed_duplicates}\n")

if __name__ == "__main__":
    main()