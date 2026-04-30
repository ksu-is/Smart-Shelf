import pandas as pd
import os
from datetime import datetime

# Path to your inventory file
INVENTORY_FILE = r"C:\Users\Admin\Desktop\Darsh\smallbiz_inventory_software\Data\inventory.xlsx"

# Expected columns
COLUMNS = [
    "SKU",
    "Product Name",
    "Category",
    "Quantity",
    "Cost Price",
    "Unit Price",
    "Profit Per Unit",
    "Supplier",
    "Last Updated",
    "Sales Count",
    "Low Stock Threshold",
    "Description"
]

def load_inventory():
    """Load inventory Excel into a DataFrame. Create empty if missing."""
    if os.path.exists(INVENTORY_FILE):
        try:
            df = pd.read_excel(INVENTORY_FILE)
            # Ensure all expected columns are present
            for col in COLUMNS:
                if col not in df.columns:
                    df[col] = None
            return df
        except Exception as e:
            print(f"[ERROR] Could not load inventory: {e}")
            return pd.DataFrame(columns=COLUMNS)
    else:
        print("[INFO] Inventory file not found. Creating empty inventory.")
        return pd.DataFrame(columns=COLUMNS)

def save_inventory(df):
    """Save the DataFrame back to Excel."""
    try:
        df.to_excel(INVENTORY_FILE, index=False)
        print(f"[INFO] Inventory saved. {len(df)} records.")
    except Exception as e:
        print(f"[ERROR] Could not save inventory: {e}")

def initialize_inventory_file():
    """Create empty file if missing."""
    if not os.path.exists(INVENTORY_FILE):
        df = pd.DataFrame(columns=COLUMNS)
        save_inventory(df)
        print("[INFO] Created empty inventory file.")
    else:
        print("[INFO] Inventory file already exists.")

def add_product(product_dict):
    """Add a new product if SKU doesn't exist."""
    df = load_inventory()
    if product_dict["SKU"] in df["SKU"].values:
        print(f"[WARN] SKU {product_dict['SKU']} already exists. Skipping.")
        return
    product_dict["Last Updated"] = datetime.now()
    product_dict["Profit Per Unit"] = round(
        product_dict["Unit Price"] - product_dict["Cost Price"], 2
    )
    df = pd.concat([df, pd.DataFrame([product_dict])], ignore_index=True)
    save_inventory(df)
    print(f"[INFO] Added product {product_dict['Product Name']} (SKU {product_dict['SKU']})")

def update_product_quantity(sku, quantity_change):
    """Adjust quantity of a product (positive or negative)."""
    df = load_inventory()
    idx = df.index[df["SKU"] == sku].tolist()
    if not idx:
        print(f"[WARN] SKU {sku} not found.")
        return
    i = idx[0]
    df.at[i, "Quantity"] = max(0, df.at[i, "Quantity"] + quantity_change)
    df.at[i, "Last Updated"] = datetime.now()
    if quantity_change < 0:
        df.at[i, "Sales Count"] += abs(quantity_change)
    save_inventory(df)
    print(f"[INFO] Updated quantity for SKU {sku}. New Qty: {df.at[i, 'Quantity']}")

def delete_product(sku):
    """Delete product by SKU."""
    df = load_inventory()
    new_df = df[df["SKU"] != sku]
    if len(new_df) == len(df):
        print(f"[WARN] SKU {sku} not found. Nothing deleted.")
    else:
        save_inventory(new_df)
        print(f"[INFO] Deleted product with SKU {sku}.")

def update_product_info(sku, updates):
    """Update fields of a product."""
    df = load_inventory()
    idx = df.index[df["SKU"] == sku].tolist()
    if not idx:
        print(f"[WARN] SKU {sku} not found.")
        return
    i = idx[0]
    for key, value in updates.items():
        if key in df.columns:
            df.at[i, key] = value
    # Recalculate profit if prices updated
    if "Cost Price" in updates or "Unit Price" in updates:
        cost = df.at[i, "Cost Price"]
        unit = df.at[i, "Unit Price"]
        df.at[i, "Profit Per Unit"] = round(unit - cost, 2)
    df.at[i, "Last Updated"] = datetime.now()
    save_inventory(df)
    print(f"[INFO] Updated product {sku}.")

def get_low_stock_items():
    """Return DataFrame of items below threshold."""
    df = load_inventory()
    low_stock_df = df[df["Quantity"] <= df["Low Stock Threshold"]]
    print(f"[INFO] Found {len(low_stock_df)} low stock items.")
    return low_stock_df

def get_product_by_sku(sku):
    """Return details of a product."""
    df = load_inventory()
    product = df[df["SKU"] == sku]
    if product.empty:
        print(f"[WARN] SKU {sku} not found.")
        return None
    return product.iloc[0].to_dict()

def generate_inventory_report():
    """Print a summary report of current inventory."""
    df = load_inventory()
    total_items = len(df)
    total_stock_value = (df["Quantity"] * df["Unit Price"]).sum()
    potential_profit = (df["Quantity"] * df["Profit Per Unit"]).sum()
    print(f"[REPORT] Total Products: {total_items}")
    print(f"[REPORT] Total Stock Value: ₹{total_stock_value:,.2f}")
    print(f"[REPORT] Potential Profit (if sold): ₹{potential_profit:,.2f}")

def update_inventory_after_sale(items):
    """
    Deduct sold quantities from inventory and save.
    
    Args:
        items (list): List of dicts with 'Product Name' and 'Quantity' to deduct
    """
    file_path = r"C:\Users\Admin\Desktop\Darsh\smallbiz_inventory_software\Data\inventory.xlsx"
    
    if not os.path.exists(file_path):
        print(f"[ERROR] Inventory file not found at {file_path}")
        return

    df = load_inventory()

    for item in items:
        product = item["Product Name"]
        qty_sold = item["Quantity"]

        idx = df[df["Product Name"].str.lower() == product.lower()].index
        if not idx.empty:
            current_qty = df.at[idx[0], "Quantity"]
            new_qty = max(current_qty - qty_sold, 0)
            df.at[idx[0], "Quantity"] = new_qty
            print(f"[INFO] Updated {product}: {current_qty} -> {new_qty}")
        else:
            print(f"[WARN] Product '{product}' not found in inventory.")

    # Save the updated inventory
    try:
        df.to_excel(file_path, index=False)
        print(f"[INFO] Inventory updated and saved to {file_path}")
    except Exception as e:
        print(f"[ERROR] Failed to save updated inventory: {e}")
