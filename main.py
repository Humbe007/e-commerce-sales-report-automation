import pandas as pd
import openpyxl as op
import os

import processes.charger as ch
import processes.cleaner as cl
import processes.sales_metrics as sales
import processes.export as ex
import processes.formatter as fm

#Part 1: Data Collection from Datasets 

dfs_weeks, df_products = ch.dataframes_from("Datasets","csv")
print("The data has been loaded successfully.")

#Part 2: Data Cleaning: Standardization of dates and column names

mapping = {"order_id" : ["Order ID","order id", "orderId","order_id","OrderID"],
           "date": ["Created at","date","Created At","created_at"],
           "product": ["Product Name","product","product_name"],
           "quantity": ["qty","Quantity","quantity"],}


dfs_weeks = [cl.standardize(df, mapping) for df in dfs_weeks]
df_merged = pd.concat(dfs_weeks, ignore_index=True)

df_merged = cl.remove_duplicate_rows_and_columns(df_merged)
print("The data has been cleaned successfully.")


df_merged = cl.standardize_dates(df_merged) # This function standarize the dates to this format: dd-mm-yyyy and convert the date column to datetime format. 
# Also, df_merged is the sheet4 of the resultant excel file.
print("The dates and columns have been standardized successfully.")

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
print("The sales metrics have been calculated successfully.")

#Part 6: Exporting to Excel:

os.makedirs('output', exist_ok=True)
try:
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
except PermissionError as e:
    input(f"An error has been occurred: {e} \n  ---Please close or remove the file 'sales_report.xlsx' at the folder and try again---\nPress Enter to finish.")

workbook = op.load_workbook('output//sales_report.xlsx')

#Part 7: Formatting
resume_sheet = workbook['Resume']
fm.format_table(resume_sheet)

top_products_sheet = workbook['Top products']
fm.format_table(top_products_sheet)
fm.apply_color_scale(worksheet=top_products_sheet, column_letter='C', green_min=100, yellow_min=90)
col_letters = ("B", "D", "E", "F","G")
for col in col_letters:
    fm.apply_gradient_column(worksheet=top_products_sheet, column_letter=col, start_color='FFFF00', end_color='66FF33')

daily_sales_sheet = workbook['Daily sales']
fm.format_table(daily_sales_sheet)
fm.apply_color_scale(worksheet=daily_sales_sheet, column_letter='C', green_min=4, yellow_min =False)

merged_data_sheet = workbook['Merged Data']
fm.format_table(merged_data_sheet)
fm.apply_color_scale(worksheet=merged_data_sheet, column_letter='D', green_min=10, yellow_min=5)

workbook.save('output//sales_report.xlsx')
print("The sales report has been generated and formatted successfully.")
input("Press Enter to finish.")