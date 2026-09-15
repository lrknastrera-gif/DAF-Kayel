import pandas as pd

products = pd.read_csv("MotorPH_Products_List_2025.csv")
print("PRODUCTS COLUMNS:")
print(products.columns.tolist())
print("\nPRODUCTS FIRST ROW:")
print(products.head(1))

sales = pd.read_csv("MotorPH_Sales Data-3rd Quarter-Year 2025.csv")
print("\n\nSALES COLUMNS:")
print(sales.columns.tolist())
print("\nSALES FIRST ROW:")
print(sales.head(1))
import pandas as pd
import numpy as np

products = pd.read_csv("MotorPH_Products_List_2025.csv")
print(f"Original shape: {products.shape}")
print(f"Columns: {products.columns.tolist()}")

# Remove missing & duplicates
products = products.dropna(subset=['EntrName'])
products = products.drop_duplicates()

# RENAME - one by one to catch errors
products = products.rename(columns={
    'EntrNo': 'Product ID Number',
    'EntrName': 'Product Name',
    'EntrDetails': 'Product Type'
})

print(f"\nAfter rename: {products.columns.tolist()}")

# Reset index and set Product ID
products = products.reset_index(drop=True)
products['Product ID Number'] = range(1, len(products) + 1)

# Add Unit Price if missing
if 'Unit Price' not in products.columns:
    products['Unit Price'] = 0.0

# Add Manufacturing Date if it exists, rename it
if 'Manufacturing Date' in products.columns:
    products = products.rename(columns={'Manufacturing Date': 'Date of Manufacturing'})
else:
    products['Date of Manufacturing'] = None

# Add Acquisition Date - check exact column name
if 'Acquisition' in products.columns:
    products = products.rename(columns={'Acquisition': 'Date of Acquisition'})
else:
    products['Date of Acquisition'] = None

print(f"Final columns before select: {products.columns.tolist()}")

# Now select only the required columns that exist
final_columns = ['Product ID Number', 'Product Name', 'Product Type', 'Unit Price', 'Date of Manufacturing', 'Date of Acquisition']
existing_columns = [col for col in final_columns if col in products.columns]

products = products[existing_columns]

print(f"Final shape: {products.shape}")
print(products.head())

products.to_csv("MotorPH_Products_Preprocessed.csv", index=False)
print("✅ Exported!")

import pandas as pd
import numpy as np

# ===== SALES PREPROCESSING =====
print("="*60)
print("PROCESSING SALES DATA")
print("="*60)

sales = pd.read_csv("MotorPH_Sales Data-3rd Quarter-Year 2025.csv")
print(f"Original shape: {sales.shape}")
print(f"Columns: {sales.columns.tolist()}")

# STEP 1: Check for missing values
print(f"\nMissing values BEFORE cleaning:")
print(sales.isnull().sum())

# STEP 2: Drop rows with missing critical values
sales = sales.dropna(subset=['product', 'date'])
print(f"After dropna: {sales.shape}")

# STEP 3: Remove duplicate records
sales = sales.drop_duplicates()
print(f"After drop_duplicates: {sales.shape}")

# STEP 4: Rename columns to cleaner format
sales = sales.rename(columns={
    'date': 'Date',
    'client_type': 'Client Type',
    'product': 'Product',
    'unitprice': 'Unit Price',
    'quantity': 'Quantity',
    'total': 'Total',
    'payment': 'Payment Method'
})

print(f"\nColumns after rename: {sales.columns.tolist()}")

# STEP 5: Fix data types
# Convert date to datetime
sales['Date'] = pd.to_datetime(sales['Date'], errors='coerce')

# Convert numeric columns
sales['Unit Price'] = pd.to_numeric(sales['Unit Price'], errors='coerce')
sales['Quantity'] = pd.to_numeric(sales['Quantity'], errors='coerce')
sales['Total'] = pd.to_numeric(sales['Total'], errors='coerce')

# STEP 6: Standardize text fields (clean casing, trim whitespace)
sales['Client Type'] = sales['Client Type'].str.strip().str.title()
sales['Product'] = sales['Product'].str.strip().str.title()
sales['Payment Method'] = sales['Payment Method'].str.strip().str.title()

# STEP 7: Final check
print(f"\nMissing values AFTER cleaning:")
print(sales.isnull().sum())

print(f"\nFinal shape: {sales.shape}")
print(f"Duplicate rows: {sales.duplicated().sum()}")
print(f"\nData types:")
print(sales.dtypes)

# EXPORT
sales.to_csv("MotorPH_Sales_Preprocessed.csv", index=False)
print("\n✅ Sales exported successfully!")
print("\nFirst 5 rows:")
print(sales.head())


products = pd.read_csv("MotorPH_Products_Preprocessed.csv")
sales = pd.read_csv("MotorPH_Sales_Preprocessed.csv")

print(f"Products: {products.shape}")
print(f"Sales: {sales.shape}")
print(products.head())
print(sales.head())