import pandas as pd
import os

def dataframes_from(folder,type="csv"):
    """This function reads all the files on the folder with the specified type 
    and returns a list of dataframes. Also it separates the weekly dataframes from the 
    products dataframe, which is the last one in the list."""
    
    dfs_weeks = []
    df_product = None
    
    files = [f for f in os.listdir(folder) if f.endswith(f".{type}")]
    
    for f in files: 
        file = os.path.join(folder, f)
        if f.endswith(".csv"):
            
            if "week" in f:
                df = pd.read_csv(file)        
                dfs_weeks.append(df)
            
            elif "product" in f:
                df_product = pd.read_csv(file)
        
        elif f.endswith(".xlsx"):
            
            if "week" in f:
                df = pd.read_excel(file)
                dfs_weeks.append(df)
                
            elif "product" in f:
                df_product = pd.read_excel(file)
        else:
            continue

    if df_product is None:
        raise ValueError("No product file found in the folder.")
    
    return dfs_weeks,df_product
