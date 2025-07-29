import pandas as pd
import sqlite3

# Step 3A: Load the CSV
df = pd.read_csv('superstore_sales.csv')
print("First 5 rows of data:")
print(df.head(), "\n")

# Step 3B: Connect to SQLite and save the DataFrame as a table
conn = sqlite3.connect('sales_analysis.db')      # Creates (or opens) the database file
df.to_sql('sales_data', conn, if_exists='replace', index=False)
print("Data saved to SQLite database 'sales_analysis.db'.\n")

# Step 3C: Run a sample SQL query using pandas
query = """
SELECT Category, 
       Region,
       ROUND(SUM(Sales), 2) AS Total_Sales,
       ROUND(SUM(Profit), 2) AS Total_Profit
FROM sales_data
GROUP BY Category, Region
ORDER BY Total_Sales DESC
"""
result = pd.read_sql(query, conn)
print("Sample SQL query – Total Sales & Profit by Category and Region:")
print(result)
# Step 4A: Top 5 Orders by Sales
top5_query = """
SELECT "Order ID", Category, Region, Sales
FROM sales_data
ORDER BY Sales DESC
LIMIT 5;
"""
top5 = pd.read_sql(top5_query, conn)
print("\nTop 5 Orders by Sales:")
print(top5)

# Step 4B: Total Sales by Category (reuse if you like)
# Already have 'result', but let's export it too.
# result is the Category/Region query from Step 3.

# Step 4C: Export to Excel
with pd.ExcelWriter('sales_report.xlsx') as writer:
    result.to_excel(writer, sheet_name='Sales_by_Cat_Region', index=False)
    top5.to_excel(writer, sheet_name='Top5_Orders', index=False)

print("\nExported query results to 'sales_report.xlsx'.")

