# Data Cleaning & Report Generator

A Python tool that automatically cleans messy Excel/CSV files and generates a summary report.  
Built as a portfolio project to demonstrate real-world data skills.

---

## What it does

- Removes duplicate rows
- Standardises text (Title Case, trims whitespace)
- Fixes and formats date columns automatically
- Flags rows with missing values
- Outputs a clean, formatted Excel file
- Prints a full cleaning summary report

---

## Skills demonstrated

| Skill | Used for |
|-------|----------|
| Python | Core scripting logic |
| pandas | Data loading, cleaning, transformation |
| openpyxl | Excel file creation and formatting |
| argparse | Command-line interface |
| File I/O | Reading .xlsx / .csv input files |

---

## How to run

### 1. Install requirements
```bash
pip install pandas openpyxl
```

### 2. Generate sample data (optional — for testing)
```bash
python generate_sample_data.py
```

### 3. Run the cleaner
```bash
python data_cleaner.py --input data/sample_data.xlsx --output data/cleaned_output.xlsx
```

---

## Sample output

```
────────────────────────────────────────────────
  DATA CLEANING REPORT
  01 Jun 2025  14:32
────────────────────────────────────────────────
  Input file   : data/sample_data.xlsx
  Output file  : data/cleaned_output.xlsx
────────────────────────────────────────────────
  Original rows    : 10
  Duplicates removed: 3
  Final rows       : 7
  Text cols cleaned : 4
  Date cols fixed  : 1

  Missing values found:
    • Product ID: 1 missing
    • Category: 1 missing
    • Stock: 1 missing

  Rows flagged 'has_missing' added to output.
────────────────────────────────────────────────
  Cleaning complete.
```

---

## Project background

Built to automate the kind of manual data cleaning work I do daily as a Data Entry Associate.  
This script turns a task that takes 30–60 minutes into a 5-second automated process.

---

## Author

**[Your Name]**  
Data Entry & Business Development Professional, Lahore  
Certified: Python Programming (Coursera) | Google Sheets (Coursera)  
[LinkedIn profile link]
