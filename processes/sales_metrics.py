import pandas as pd
def get_product_stats(df_grouped, df_products, product):
    
    qty_series = df_grouped.loc[df_grouped['product'] == product, 'quantity']  
    qty = qty_series.iloc[0] if not qty_series.empty else 0
    
    cost_series = df_products.loc[df_products['product'] == product, 'cost']
    cost = cost_series.iloc[0] if not cost_series.empty else 0
    
    price_series = df_products.loc[df_products['product'] == product, 'price']
    price = price_series.iloc[0] if not price_series.empty else 0
    
    return {
        'product': product,
        'total_profit': qty * price - qty * cost,
        'quantity': qty,
        'total_cost': qty * cost,
        'total_income': qty * price,
        'individual_cost': cost,
        'individual_price': price,  
    }

def get_totals(df_merged,df_top_products):
    df_merged = df_merged.copy()
    df_top_products = df_top_products.copy()
    
    total_costs = df_top_products['total_cost'].sum()
    total_profits = df_top_products['total_profit'].sum()
    total_incomes = df_top_products['total_income'].sum()
    total_qty_sells = df_merged.shape[0]
    average_order_value = total_incomes / total_qty_sells if total_qty_sells > 0 else 0
    return {
        'metric': "quantity",
        'total_profits': total_profits,
        'total_costs': total_costs,
        'total_incomes': total_incomes,
        'number_of_sales': total_qty_sells,
        'average_order_value': average_order_value.round(2)
    }
    
def sheet3(df_merged, df_products):
    df_merged = df_merged.copy()
    df_products = df_products.copy()
    
    df = df_merged.merge(df_products[['product', 'price']], on='product', how='left')
    df['revenue'] = df['quantity'] * df['price']
    
    return df.groupby('date').agg(
        revenue=('revenue', 'sum'),
        number_of_sales=('quantity', 'count')
        ).reset_index()
    

