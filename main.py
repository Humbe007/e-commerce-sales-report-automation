import pandas as pd
import os
import processes.charger as ch
import processes.cleaner as cl
import processes.sales_metrics as sales
import processes.export as ex

#Part 1: Data Collection from Datasets 

dfs_weeks, df_products = ch.dataframes_from("Datasets","csv")

#Part 2: Data Cleaning: Standardization of dates and column names

mapping = {"order_id" : ["Order ID","order id", "orderId","order_id","OrderID"],
           "date": ["Created at","date","Created At","created_at"],
           "product": ["Product Name","product","product_name"],
           "quantity": ["qty","Quantity","quantity"],}


dfs_weeks = [cl.standardize(df, mapping) for df in dfs_weeks]
df_merged = pd.concat(dfs_weeks, ignore_index=True)

df_merged = cl.remove_duplicate_rows_and_columns(df_merged)
df_merged = cl.standardize_dates(df_merged) # This function standarize the dates to this format: dd-mm-yyyy and convert the date column to datetime format. 
# Also, df_merged is the sheet4 of the resultant excel file.

#Part 3: Sheet 2: Top products by profit
df_grouped = df_merged.groupby('product')['quantity'].sum().reset_index()

df_top_products = pd.DataFrame([
    sales.get_product_stats(df_grouped, df_products, p) for p in df_products['product'].unique()
])

df_top_products = df_top_products.sort_values(by='total_profit', ascending=False).reset_index(drop=True)

#Part 4: Sheet 1
df_sheet1 = pd.DataFrame(sales.get_totals(df_merged,df_top_products), index=[0])

df_sheet1 = df_sheet1.set_index('metric').T.reset_index()
df_sheet1.columns = ['metric', 'quantity']

#Part 5: Sheet 3: Sales trend

df_sheet3 = sales.sheet3(df_merged,df_products)

#Part 6: Exporting to Excel:

os.makedirs('output', exist_ok=True)
with pd.ExcelWriter('output//sales_report.xlsx',engine='openpyxl') as writer:
    df_sheet1.to_excel(writer, sheet_name='Resume', index=False)
    df_top_products.to_excel(writer, sheet_name='Top products', index=False)
    df_sheet3.to_excel(writer, sheet_name='Daily sales', index=False)
    df_merged.to_excel(writer, sheet_name='Merged Data', index=False)
    
    for sheet_name in writer.sheets:
        worksheet = writer.sheets[sheet_name]
    
        # Ajustar columnas
        ex.adjust_column_dimensions(worksheet)
    
        # Freeze panes (congelar fila superior)
        worksheet.freeze_panes = "A2"