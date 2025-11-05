# generate_sales_data.py
# Generates sample sales data (2 months) and saves to Sales_Dashboard.xlsx

import pandas as pd
import numpy as np
from datetime import datetime

np.random.seed(42)

start_date = datetime(2024,6,1)
end_date = datetime(2024,7,31)
dates = pd.date_range(start_date, end_date, freq='D')

products = ['Laptop', 'Smartphone', 'Tablet', 'Headphones', 'Smartwatch']
regions = ['North', 'South', 'East', 'West']

data = []
for _ in range(80):
    date = np.random.choice(dates)
    product = np.random.choice(products)
    region = np.random.choice(regions)
    quantity = np.random.randint(1, 10)
    price = np.random.randint(5000, 50000)
    sales = quantity * price
    profit = round(sales * np.random.uniform(0.1, 0.25), 2)
    data.append([date, product, region, quantity, sales, profit])

df = pd.DataFrame(data, columns=['Date','Product','Region','Quantity','Sales','Profit'])
df.to_excel('Sales_Dashboard.xlsx', index=False)
print('Saved Sales_Dashboard.xlsx with', len(df), 'rows')
