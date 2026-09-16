import pandas as pd
import numpy as np

# Load the preprocessed dataset
products = pd.read_csv("MotorPH_Products_Preprocessed.csv")

# Display first few rows
print("="*70)
print("MOTORPH PRODUCT ANALYSIS")
print("="*70)
print("\nFirst 5 rows of dataset:")
print(products.head())

print(f"\nDataset shape: {products.shape}")
print(f"Columns: {products.columns.tolist()}")
print(f"\nData types:")
print(products.dtypes)

# STEP 1: Count total number of products
total_products = len(products)

print("\n" + "="*70)
print("ANALYSIS 1: TOTAL NUMBER OF PRODUCTS")
print("="*70)

print(f"\nTotal number of products in inventory: {total_products}")
# STEP 2: Count products by type
product_type_counts = products['Product Type'].value_counts()

print("\n" + "="*70)
print("ANALYSIS 2: PRODUCT COUNTS BY TYPE")
print("="*70)

print("\nNumber of products in each category:")
print(product_type_counts)

print(f"\nTotal product types: {len(product_type_counts)}")
# STEP 3: Unit Price Statistics
print("\n" + "="*70)
print("ANALYSIS 3: UNIT PRICE STATISTICS")
print("="*70)

# Convert Unit Price to numeric (in case of any formatting issues)
products['Unit Price'] = pd.to_numeric(products['Unit Price'], errors='coerce')

# Calculate statistics
average_price = products['Unit Price'].mean()
min_price = products['Unit Price'].min()
max_price = products['Unit Price'].max()

print(f"\nAverage (Mean) Unit Price: ₱{average_price:,.2f}")
print(f"Minimum (Cheapest) Unit Price: ₱{min_price:,.2f}")
print(f"Maximum (Most Expensive) Unit Price: ₱{max_price:,.2f}")

# Find cheapest products
cheapest = products[products['Unit Price'] == min_price][['Product ID Number', 'Product Name', 'Product Type', 'Unit Price']]
print("\nCheapest Product(s):")
print(cheapest)

# Find most expensive products
most_expensive = products[products['Unit Price'] == max_price][['Product ID Number', 'Product Name', 'Product Type', 'Unit Price']]
print("\nMost Expensive Product(s):")
print(most_expensive)

# STEP 4: Total Inventory Cost
print("\n" + "="*70)
print("ANALYSIS 4: TOTAL INVENTORY COST")
print("="*70)

# Calculate total cost
total_inventory_cost = products['Unit Price'].sum()

print(f"\nTotal cost of all items in inventory: ₱{total_inventory_cost:,.2f}")