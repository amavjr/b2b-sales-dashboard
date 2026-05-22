import pandas as pd
import numpy as np
from datetime import datetime, timedelta

np.random.seed(42)

# Create date range
dates = pd.date_range('2022-01-01', '2024-12-31', freq='D')

# Define dimensions
regions = ['Northeast', 'Southeast', 'Midwest', 'Southwest', 'West']
products = ['Laptop Pro', 'Wireless Mouse', 'Mechanical Keyboard', '27" Monitor', 'USB-C Dock']
categories = ['Hardware', 'Accessories', 'Accessories', 'Hardware', 'Accessories']
sales_reps = {
    'Northeast': ['Marcus Webb', 'Olivia Chen'],
    'Southeast': ['James Okoro', 'Diana Torres'],
    'Midwest': ['Sarah Nowak', 'Tyler Hughes'],
    'Southwest': ['Rachel Kim', 'Carlos Mendez'],
    'West': ['Naomi Sato', 'Derek Patel']
}

# Generate orders
orders = []
order_id = 10000

for date in dates:
    for region in regions:
        num_orders = np.random.randint(2, 10)
        for _ in range(num_orders):
            rep = np.random.choice(sales_reps[region])
            product_idx = np.random.randint(0, len(products))
            product = products[product_idx]
            category = categories[product_idx]
            quantity = np.random.randint(1, 10)
            unit_price = np.random.choice([29.99, 49.99, 99.99, 199.99, 399.99, 799.99, 1299.99])
            discount = np.random.choice([0, 0.05, 0.10, 0.15, 0.20], p=[0.6, 0.15, 0.1, 0.1, 0.05])
            revenue = quantity * unit_price * (1 - discount)
            cogs = revenue * np.random.uniform(0.35, 0.65)
            profit = revenue - cogs
            deal_stage = np.random.choice(['Closed Won', 'Closed Lost'], p=[0.62, 0.38])
            
            orders.append({
                'Order_ID': order_id,
                'Order_Date': date,
                'Region': region,
                'Sales_Rep': rep,
                'Product': product,
                'Category': category,
                'Quantity': quantity,
                'Unit_Price': unit_price,
                'Discount_Pct': discount * 100,
                'Revenue': round(revenue, 2),
                'COGS': round(cogs, 2),
                'Profit': round(profit, 2),
                'Deal_Stage': deal_stage
            })
            order_id += 1

df = pd.DataFrame(orders)

# Add calculated columns
df['Year'] = df['Order_Date'].dt.year
df['Month'] = df['Order_Date'].dt.month_name()
df['Quarter'] = df['Order_Date'].dt.quarter
df['Profit_Margin'] = (df['Profit'] / df['Revenue'] * 100).round(1)

# Save to Excel for Tableau
df.to_excel('data/sales_data.xlsx', sheet_name='Raw Orders', index=False)

# Create summary sheets for Tableau
summary_by_rep = df[df['Deal_Stage'] == 'Closed Won'].groupby(['Sales_Rep', 'Region']).agg({
    'Revenue': 'sum',
    'Profit': 'sum',
    'Order_ID': 'count'
}).rename(columns={'Order_ID': 'Wins'}).round(2)

summary_by_product = df[df['Deal_Stage'] == 'Closed Won'].groupby(['Product', 'Category']).agg({
    'Revenue': 'sum',
    'Profit': 'sum',
    'Quantity': 'sum'
}).round(2)

summary_by_month = df[df['Deal_Stage'] == 'Closed Won'].groupby(['Year', 'Month']).agg({
    'Revenue': 'sum',
    'Profit': 'sum',
    'Order_ID': 'count'
}).round(2)

# Save summaries to same Excel file
with pd.ExcelWriter('data/sales_data.xlsx', engine='openpyxl', mode='a') as writer:
    summary_by_rep.to_excel(writer, sheet_name='Rep Summary')
    summary_by_product.to_excel(writer, sheet_name='Product Summary')
    summary_by_month.to_excel(writer, sheet_name='Monthly Trend')

print("Data generation complete!")
print(f"Total orders: {len(df):,}")
print(f"Total revenue (won deals): ")
print(f"File saved to: data/sales_data.xlsx")
print("")
print("Ready to open in Tableau!")
