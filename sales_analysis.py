import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os

# Load the data
df = pd.read_excel("C:\\Users\\haroo\\OneDrive\\Dokumen\\Desktop\\Retail-Sales-Analysis\\superstore.xlsx")

# Clean the data
df.dropna(inplace=True)
df['Order Date'] = pd.to_datetime(df['Order Date'])

# Create outputs folder
if not os.path.exists('outputs'):
    os.makedirs('outputs')

# Top-selling categories
# Top-selling categories
top_categories = df.groupby('Category')['Sales'].sum().sort_values(ascending=False)
plt.figure(figsize=(8, 5))
sns.barplot(
    x=top_categories.values,
    y=top_categories.index,
    hue=top_categories.index,
    palette="viridis",
    legend=False
)
plt.title("Top-Selling Product Categories")
plt.xlabel("Sales")
plt.ylabel("Category")
plt.tight_layout()
plt.savefig("outputs/top_categories.png")
plt.close()

# Monthly revenue trend
df['Month'] = df['Order Date'].dt.to_period('M').astype(str)
monthly_sales = df.groupby('Month')['Sales'].sum()
plt.figure(figsize=(10, 5))
monthly_sales.plot(marker='o', color='teal')
plt.title("Monthly Sales Trend")
plt.xlabel("Month")
plt.ylabel('Sales')
plt.xticks(rotation=45)
plt.grid(True)
plt.tight_layout()
plt.savefig("outputs/monthly_trend.png")
plt.close()

# Sales by region
# Sales by region
regional_sales = df.groupby('Region')['Sales'].sum().sort_values()
plt.figure(figsize=(7, 5))
sns.barplot(
    x=regional_sales.values,
    y=regional_sales.index,
    hue=regional_sales.index,
    palette='magma',
    legend=False
)
plt.title("Sales by Region")
plt.xlabel("Sales")
plt.ylabel("Region")
plt.tight_layout()
plt.savefig("outputs/regional_sales.png")
plt.close()


print("✅ Charts generated and saved in the outputs/ folder.")
