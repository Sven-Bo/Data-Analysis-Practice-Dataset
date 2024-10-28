"""
Author: Sven Bosau
Website: https://pythonandvba.com
YouTube Channel: https://youtube.com/@codingisfun

Description:
This script generates a realistic sales dataset for data analysis practice, spanning a configurable number of years 
with seasonal trends and customer behavior patterns.

Analysis Ideas:
1. Sales Performance: Analyze total sales by product, category, or store location to identify top performers.
2. Customer Segmentation: Use customer types (Loyal vs. Occasional) and segments to understand purchase frequency and customer value.
3. Seasonal Trends: Examine monthly sales patterns, especially around peak months, to plan inventory and marketing strategies.
4. Discount Impact: Assess how discounts affect sales volume and total revenue. Do higher discounts correlate with increased sales?
5. Supplier Analysis: Identify which suppliers support the highest revenue-generating products.
"""

# Import necessary libraries
import pandas as pd
import numpy as np
from faker import Faker
from random import randint, choice, uniform
import datetime
from pathlib import Path

# SETTINGS
dataset_name = "Sales_Data.xlsx"  # Name of the output Excel file
num_sales_records = 5000  # Number of rows in Sales Data
num_customers = 100  # Number of unique customers
num_suppliers = 10  # Number of suppliers
store_locations = [
    "Berlin",
    "Munich",
    "Hamburg",
    "Cologne",
    "Frankfurt",
]  # Store locations for sales data
starting_year = 2023  # Starting year for the dataset
years_span = 2  # Number of years to cover in the dataset

# Initialize Faker for generating realistic names, emails, etc.
fake = Faker()

# Define product details with categories and price ranges
product_categories = {
    "Laptop": {"Category": "Electronics", "Price Range": (800, 1500)},
    "Smartphone": {"Category": "Electronics", "Price Range": (300, 1000)},
    "Tablet": {"Category": "Electronics", "Price Range": (200, 600)},
    "Headphones": {"Category": "Accessories", "Price Range": (30, 150)},
    "Monitor": {"Category": "Electronics", "Price Range": (100, 300)},
    "Keyboard": {"Category": "Accessories", "Price Range": (20, 100)},
    "Mouse": {"Category": "Accessories", "Price Range": (10, 50)},
    "Printer": {"Category": "Office Supplies", "Price Range": (100, 400)},
    "Wireless Charger": {"Category": "Accessories", "Price Range": (20, 40)},
    "Bluetooth Speaker": {"Category": "Electronics", "Price Range": (70, 90)},
    "Office Chair": {"Category": "Furniture", "Price Range": (140, 160)},
    "LED Lamp": {"Category": "Office Supplies", "Price Range": (15, 35)},
    "Backpack": {"Category": "Apparel", "Price Range": (30, 50)},
    "Desk Organizer": {"Category": "Office Supplies", "Price Range": (10, 20)},
    "USB-C Hub": {"Category": "Electronics", "Price Range": (40, 50)},
    "Portable SSD": {"Category": "Electronics", "Price Range": (110, 130)},
    "Smartwatch": {"Category": "Electronics", "Price Range": (180, 220)},
    "Digital Camera": {"Category": "Electronics", "Price Range": (290, 310)},
}

# Define supplier data without contact and rating info
supplier_data = {
    "Supplier ID": [f"SUP{i+1}" for i in range(num_suppliers)],
    "Supplier Name": [fake.company() for _ in range(num_suppliers)],
    "Location": [
        choice(["Germany", "Netherlands", "France", "Poland", "Italy", "Spain", "UK"])
        for _ in range(num_suppliers)
    ],
    "Specialization": [
        choice(
            ["Electronics", "Accessories", "Office Supplies", "Furniture", "Apparel"]
        )
        for _ in range(num_suppliers)
    ],
}
df_suppliers = pd.DataFrame(supplier_data)


# Generate product data, ensuring prices end in .99 or .95
def adjust_price(price):
    return round(price) - 0.01 if price % 1 < 0.5 else round(price) - 0.05


product_data = {
    "Product ID": [f"PROD{i+1}" for i in range(len(product_categories))],
    "Product Name": list(product_categories.keys()),
    "Category": [details["Category"] for details in product_categories.values()],
    "Unit Price": [
        adjust_price(uniform(*details["Price Range"]))
        for details in product_categories.values()
    ],
    "Supplier ID": [
        choice(df_suppliers["Supplier ID"]) for _ in range(len(product_categories))
    ],
}
df_products = pd.DataFrame(product_data)

# Generate customer data with zip codes in locations
customer_data = {
    "Customer ID": [f"CUST{1000 + i}" for i in range(num_customers)],
    "Name": [fake.name() for _ in range(num_customers)],
    "Email": [fake.email() for _ in range(num_customers)],
    "Location": [f"{fake.zipcode()} {fake.city()}" for _ in range(num_customers)],
    "Customer Segment": [
        choice(["Individual", "Business", "Enterprise"]) for _ in range(num_customers)
    ],
}
df_customers = pd.DataFrame(customer_data)

# Define seasonal trends for each product category
seasonal_trends = {
    "Electronics": [12],
    "Accessories": [6, 11],
    "Office Supplies": [1, 9],
    "Furniture": [3, 10],
    "Apparel": [8, 12],
}


# Generate dates based on seasonal trends within the specified years
def apply_seasonal_trend(category):
    month = np.random.choice(seasonal_trends.get(category, [randint(1, 12)]))
    year = np.random.choice(range(starting_year, starting_year + years_span))
    day = randint(1, 28)
    return datetime.date(year, month, day)


# Generate sales data with sequential Order IDs and Purchase Dates
sales_data = []
current_date = datetime.date(starting_year, 1, 1)  # Starting point for sales
for i in range(num_sales_records):
    product_id = choice(df_products["Product ID"])
    customer_id = choice(df_customers["Customer ID"])
    customer_type = (
        "Loyal"
        if customer_id
        in np.random.choice(
            df_customers["Customer ID"], size=num_customers // 4, replace=False
        )
        else "Occasional"
    )
    quantity = randint(1, 5)
    unit_price = df_products[df_products["Product ID"] == product_id][
        "Unit Price"
    ].values[0]
    total_price = round(quantity * unit_price, 2)
    discount = np.random.choice([0, 0, 5, 10, 15]) if customer_type == "Loyal" else 0
    discount_amount = round(total_price * (discount / 100), 2)
    final_price = total_price - discount_amount

    # Increment date within the year range
    current_date += datetime.timedelta(days=randint(0, 1))
    if current_date.year >= starting_year + years_span:
        break

    sales_data.append(
        {
            "Order ID": f"ORD{1000 + i}",
            "Purchase Date": current_date,
            "Customer ID": customer_id,
            "Customer Type": customer_type,
            "Store Location": choice(store_locations),
            "Product ID": product_id,
            "Quantity": quantity,
            "Unit Price": unit_price,
            "Total Price": total_price,
            "Discount Applied (%)": discount,
            "Discount Amount": discount_amount,
            "Final Price": final_price,
        }
    )

df_sales = pd.DataFrame(sales_data)

# Set up the output path with branding sheet, auto-fit columns, and saving to Excel
output_path = Path.cwd() / dataset_name
with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
    workbook = writer.book
    about_sheet = workbook.add_worksheet("About This Dataset")
    bold = workbook.add_format({"bold": True})
    link_format = workbook.add_format({"color": "blue", "underline": True})

    # Add branding and links
    about_sheet.write("A1", "Created by Sven Bosau", bold)
    about_sheet.write("A3", "Website:", bold)
    about_sheet.write_url(
        "B3", "https://pythonandvba.com", link_format, "https://pythonandvba.com"
    )
    about_sheet.write("A4", "YouTube:", bold)
    about_sheet.write_url(
        "B4",
        "https://youtube.com/@codingisfun",
        link_format,
        "https://youtube.com/@codingisfun",
    )
    about_sheet.write("A5", "Explore all Excel Solutions:", bold)
    about_sheet.write_url(
        "B5",
        "https://pythonandvba.com/solutions",
        link_format,
        "https://pythonandvba.com/solutions",
    )
    about_sheet.write("A6", "Pandas & Data Analysis Playlist:", bold)
    about_sheet.write_url(
        "B6",
        "https://www.youtube.com/playlist?list=PL7QI8ORyVSCYgVLeZ2JA3WDG9SbbloBdM",
        link_format,
        "YouTube Playlist",
    )
    about_sheet.write("A8", "Analysis Ideas:", bold)
    about_sheet.write(
        "A9", "1. Sales Performance by product, category, or store location"
    )
    about_sheet.write("A10", "2. Customer Segmentation by frequency and value")
    about_sheet.write("A11", "3. Seasonal Trends for monthly sales")
    about_sheet.write("A12", "4. Discount Impact on volume and revenue")
    about_sheet.write("A13", "5. Supplier Analysis for top revenue products")

    # Saving datasets with header-based column widths and coloring tabs
    df_sales.to_excel(writer, sheet_name="Sales Data", index=False)
    df_customers.to_excel(writer, sheet_name="Customers", index=False)
    df_products.to_excel(writer, sheet_name="Products", index=False)
    df_suppliers.to_excel(writer, sheet_name="Suppliers", index=False)

    # Set tab colors
    writer.sheets["About This Dataset"].set_tab_color("yellow")
    writer.sheets["Sales Data"].set_tab_color("blue")
    writer.sheets["Customers"].set_tab_color("green")
    writer.sheets["Products"].set_tab_color("orange")
    writer.sheets["Suppliers"].set_tab_color("purple")

    # Set column widths based on header length with padding
    for sheet_name, df in zip(
        writer.sheets.keys(), [df_sales, df_customers, df_products, df_suppliers]
    ):
        worksheet = writer.sheets[sheet_name]
        for i, column in enumerate(df.columns):
            header_length = len(column) + 10  # Add padding
            worksheet.set_column(i, i, header_length)

print("Dataset generated and saved at:", output_path)
