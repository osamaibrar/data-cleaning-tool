"""
generate_sample_data.py
Creates a sample messy Excel file to test the data_cleaner script.
"""

import pandas as pd
import os

os.makedirs("data", exist_ok=True)

data = {
    "Product ID": [
        "P001", "P002", "P003", "P002", "P004",
        "P005", "P006", "P006", "P007", None
    ],
    "Product Name": [
        "  apple juice ", "MANGO DRINK", "orange soda", "MANGO DRINK",
        "lemon water  ", "  grape juice", "PEACH TEA", "PEACH TEA",
        "berry blend", "  watermelon juice"
    ],
    "Category": [
        "drinks", "DRINKS", "Drinks", "DRINKS", "drinks",
        "DRINKS", "drinks", "drinks", "DRINKS", None
    ],
    "Price (PKR)": [
        150, 200, 180, 200, 90,
        220, 170, 170, 195, 210
    ],
    "Stock": [
        50, 30, None, 30, 75,
        20, 45, 45, 60, 15
    ],
    "Entry Date": [
        "1/5/2025", "2025-05-02", "03-05-2025", "2025-05-02",
        "4 May 2025", "5/5/2025", "2025-05-06", "2025-05-06",
        "07/05/2025", "8-05-2025"
    ],
    "Entered By": [
        "ali khan", "SARA AHMED", "  Bilal  ", "SARA AHMED",
        "usman", "FATIMA MALIK", "ali khan", "ali khan",
        "  ZARA  ", "bilal"
    ]
}

df = pd.DataFrame(data)
df.to_excel("data/sample_data.xlsx", index=False)
print("[OK] Sample file created → data/sample_data.xlsx")
print(f"     Rows: {len(df)} (includes duplicates, missing values, messy text)")
