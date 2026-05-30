"""
Data Cleaning & Report Generator
=================================
Author  : [Your Name]
Skills  : Python, pandas, openpyxl
Purpose : Automatically cleans messy Excel/CSV files and generates
          a summary report — reducing manual data-entry effort.

How to run:
    python data_cleaner.py --input data/sample_data.xlsx --output data/cleaned_output.xlsx
"""

import pandas as pd
import argparse
import os
from datetime import datetime


# ──────────────────────────────────────────────
#  1. Load the file
# ──────────────────────────────────────────────

def load_file(filepath: str) -> pd.DataFrame:
    """Load an Excel or CSV file into a DataFrame."""
    ext = os.path.splitext(filepath)[1].lower()
    if ext in [".xlsx", ".xls"]:
        df = pd.read_excel(filepath)
    elif ext == ".csv":
        df = pd.read_csv(filepath)
    else:
        raise ValueError(f"Unsupported file type: {ext}. Use .xlsx or .csv")
    print(f"[OK] Loaded '{filepath}' — {len(df)} rows, {len(df.columns)} columns")
    return df


# ──────────────────────────────────────────────
#  2. Clean the data
# ──────────────────────────────────────────────

def clean_data(df: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    """
    Apply a series of cleaning steps and return the cleaned
    DataFrame plus a report dictionary summarising what changed.
    """
    report = {}
    original_rows = len(df)

    # Step 1: Strip leading/trailing whitespace from all text columns
    text_cols = df.select_dtypes(include="object").columns
    for col in text_cols:
        df[col] = df[col].str.strip()
    report["whitespace_fixed"] = len(text_cols)

    # Step 2: Standardise text columns to Title Case
    for col in text_cols:
        df[col] = df[col].str.title()
    report["text_standardised"] = len(text_cols)

    # Step 3: Remove fully duplicate rows
    before = len(df)
    df = df.drop_duplicates()
    report["duplicates_removed"] = before - len(df)

    # Step 4: Count missing values per column
    missing = df.isnull().sum()
    report["missing_values"] = missing[missing > 0].to_dict()

    # Step 5: Flag rows that still have any missing value
    df["has_missing"] = df.isnull().any(axis=1).map({True: "Yes", False: "No"})

    # Step 6: Standardise date columns (any column with "date" in the name)
    date_cols = [c for c in df.columns if "date" in c.lower()]
    for col in date_cols:
        try:
            df[col] = pd.to_datetime(df[col], dayfirst=True, errors="coerce")
            df[col] = df[col].dt.strftime("%d-%m-%Y")
        except Exception:
            pass
    report["date_columns_fixed"] = len(date_cols)

    report["original_rows"] = original_rows
    report["final_rows"] = len(df)
    report["rows_removed"] = original_rows - len(df)

    return df, report


# ──────────────────────────────────────────────
#  3. Save the cleaned file
# ──────────────────────────────────────────────

def save_output(df: pd.DataFrame, output_path: str):
    """Save the cleaned DataFrame to an Excel file with basic formatting."""
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Cleaned Data")

        # Auto-fit column widths
        ws = writer.sheets["Cleaned Data"]
        for col_cells in ws.columns:
            max_len = max(
                len(str(cell.value)) if cell.value is not None else 0
                for cell in col_cells
            )
            ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 4, 40)

    print(f"[OK] Cleaned file saved → '{output_path}'")


# ──────────────────────────────────────────────
#  4. Print the summary report
# ──────────────────────────────────────────────

def print_report(report: dict, input_file: str, output_file: str):
    """Print a human-readable cleaning summary to the console."""
    line = "─" * 48
    print(f"\n{line}")
    print("  DATA CLEANING REPORT")
    print(f"  {datetime.now().strftime('%d %b %Y  %H:%M')}")
    print(line)
    print(f"  Input file   : {input_file}")
    print(f"  Output file  : {output_file}")
    print(line)
    print(f"  Original rows    : {report['original_rows']}")
    print(f"  Duplicates removed: {report['duplicates_removed']}")
    print(f"  Final rows       : {report['final_rows']}")
    print(f"  Text cols cleaned : {report['text_standardised']}")
    print(f"  Date cols fixed  : {report['date_columns_fixed']}")

    if report["missing_values"]:
        print(f"\n  Missing values found:")
        for col, count in report["missing_values"].items():
            print(f"    • {col}: {count} missing")
    else:
        print(f"\n  No missing values found.")

    print(f"\n  Rows flagged 'has_missing' added to output.")
    print(line)
    print("  Cleaning complete.\n")


# ──────────────────────────────────────────────
#  5. Main entry point
# ──────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Clean a messy Excel/CSV file.")
    parser.add_argument("--input",  required=True, help="Path to the input file (.xlsx or .csv)")
    parser.add_argument("--output", required=True, help="Path for the cleaned output file (.xlsx)")
    args = parser.parse_args()

    df = load_file(args.input)
    df_clean, report = clean_data(df)
    save_output(df_clean, args.output)
    print_report(report, args.input, args.output)


if __name__ == "__main__":
    main()
