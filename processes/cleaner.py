import pandas as pd

def standardize(df,mapping):
    df = df.copy()
    for standard_col, possible_cols in mapping.items():
        for col in possible_cols:
            if col in df.columns:
                df.rename(columns={col: standard_col}, inplace=True)
                break
    return df


def remove_duplicate_rows_and_columns(df):
    """
Removes duplicate rows and columns with all duplicate data.
    """
    clean_df = df.copy()
    clean_df = clean_df.drop_duplicates()
    clean_df = clean_df.loc[: , ~clean_df.T.duplicated()] 

# df.loc[rows, columns] returns the DataFrame of the selected rows and columns.
# : selects all columns, and ~clean_df.T.duplicated returns a boolean series of the non-duplicate columns for selection.
# Note: .duplicated() always checks for identical rows, and what we do is transpose with .T to check for identical columns instead.

    return clean_df

def standardize_dates(df):
    """
    Takes a DataFrame and standardizes the date format to "dd/mm/yyyy". 
    It also handles errors by coercing invalid dates to NaT (Not a Time).
    """
    
    df = df.copy()

    def parse_date(x):
        s = str(x).strip()

        s = s.replace("/", "-")
        
        if "-"  in s[2] and "-" in s[5]: # if the format is "dd-mm-yyyy" or "mm-dd-yyyy"
            
            if len(s) == 10 and s[0:2] > "12" and s[0:2] < "32" and s[3:5] > "12": # if the format is "dd-mm-yyyy", d>=12 and d<32
                print(s,1)
                return pd.to_datetime(s, format="%d-%m-%Y")
            
            if len(s) == 10 and s[3:5] > "12" and s[3:5] < "32" and s[0:2] < "12": # if the format is "mm-dd-yyyy, d>=12 and d<32
                return pd.to_datetime(s, format="%m-%d-%Y")
            
            if len(s) == 10 and s[2] == "-": # If the format is "dd-mm-yyyy"
                return pd.to_datetime(s, format="%d-%m-%Y")
        else:
             
            if len(s) == 10 and s[4] == "-": # If the format is "yyyy-mm-dd"
                return pd.to_datetime(s, format="%Y-%m-%d")

        return pd.NaT 

    df["date"] = df["date"].apply(parse_date)
    df["date"] = df["date"].dt.strftime("%d/%m/%Y")

    return df