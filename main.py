# import libraries
import pandas as pd, numpy as np, seaborn as sns, matplotlib.pyplot as plt, datetime as dt
# Read Excel file with multiple sheets into a dataframe
dfs = pd.ExcelFile('/Users/sreeshareddy/Desktop/Raw Data.xlsx')

df_orders = pd.read_excel(dfs, sheet_name='orders')
df_customers = pd.read_excel(dfs, sheet_name='customers')
df_products = pd.read_excel(dfs, sheet_name='products')


#Display number of rows and columns
print(df_orders.shape)
print(df_customers.shape)
print(df_products.shape)
# copy and get first five row information
summary = df_orders.head()
summary.to_clipboard()

summary = df_customers.head()
summary.to_clipboard()

summary =df_products.head()
summary.to_clipboard()

#copy and get statistics for columns

summary = df_orders.describe()
summary.to_clipboard()

summary = df_customers.describe()
summary.to_clipboard()

summary = df_products.describe()
summary.to_clipboard()

# drop empty columns in df_orders
drop_columns = ["Customer Name", "Email", "Country", "Coffee Type", "Roast Type","Size", "Unit Price", "Sales"]
df_orders.drop(drop_columns, axis=1, inplace=True)

# Merge Sheets
df_1 = pd.merge(df_orders, df_customers, on="Customer ID")
df1 = pd.merge(df_1, df_products, on="Product ID")

# re-run the analysis
print(df1.shape)
summary = df1.head()
summary.to_clipboard()

summary = df1.describe()
summary.to_clipboard()

# Check for NULL values
print(df1.isna().sum())

# Drop columns with NULL values
df1 = df1.drop(["Email", "Phone Number"], axis =1)

# finding unique values in the categorical column
print("Unique Values in the column Country:", df1['Country'].unique())
print("Unique Values in the column Coffee Type:", df1['Coffee Type'].unique())
print("Unique Values in the column Roast Type:", df1['Roast Type'].unique())
print("Unique Values in the column Size:", df1['Size'].unique())
print("Unique Values in the column Loyalty Card:", df1['Loyalty Card'].unique())

# count for values in categorical columns
sns.countplot(x='Country', data=df1)
plt.show()
sns.countplot(x='Coffee Type', data=df1)
plt.show()
sns.countplot(x='Roast Type', data=df1)
plt.show()
sns.countplot(x='Loyalty Card', data=df1)
plt.show()

# Drop columns that are not useful for analysis
df1 = df1.drop(['Order ID', 'Customer ID', 'Product ID', 'Customer Name','City', 'Postcode', 'Address Line 1'], axis=1)

# Convert categorical data to numerical
Country_map = {'United States':1, 'United Kingdom':2, 'Ireland':3}
CoffeeType_map = {'Ara':1, 'Exc':2, 'Lib':3, 'Rob':4}
RoastType_map = {'M':1, 'L':2, 'D':3}
LoyaltyCard_map = {'Yes':1, 'No':0}

df1['Country'] = df1['Country'].map(Country_map)
df1['Coffee Type'] = df1['Coffee Type'].map(CoffeeType_map)
df1['Roast Type'] = df1['Roast Type'].map(RoastType_map)
df1['Loyalty Card'] = df1['Loyalty Card'].map(LoyaltyCard_map)

# Re-run descriptive analysis
print(df1.info())

# Identifying Outliers
sns.boxplot(x='Profit', data=df1, color='b')
plt.title("Boxplot for 'Profit'")
plt.show()

# Generate heatmap
sns.heatmap(df1.corr(), annot = True)
plt.title("Correlation heatmap for Coffee Sales")
plt.show()

# Linear Regression plots for independent and dependent variables
sns.set(color_codes=True)
sns.regplot(x="Unit Price", y="Profit", data=df1, scatter_kws={"alpha": 0.3, "color":'b'}, color='orange', label='Unit Price')
plt.title("Unit Price and Profit")
plt.legend(loc='upper right')
plt.show()

sns.set(color_codes=True)
sns.regplot(x="Size", y="Unit Price", data=df1, scatter_kws={"alpha": 0.3, "color":'b'}, color='orange', label='Unit Price & Size')
plt.title("Size and Unit Price")
plt.legend(loc='upper right')
plt.show()

sns.set(color_codes=True)
sns.regplot(x="Size", y="Profit", data=df1, scatter_kws={"alpha": 0.3, "color":'b'}, color='orange', label='Size')
plt.title("Size and Profit")
plt.legend(loc='upper right')
plt.show()

















