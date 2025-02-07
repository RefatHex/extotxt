# Excel to Text and Formatted Excel Converter

## Overview

This script processes an Excel file containing transaction data, extracts relevant information, and generates:

- A **fixed-width text file** (`.txt`) with formatted transaction records.
- A **new Excel file** (`.xlsx`) that follows the CSR file column structure.

## Features

✅ Extracts **transaction type** from the filename.  
✅ Removes **dashes from NDC numbers**.  
✅ **Dynamically determines** the last date of the month.  
✅ **Creates both a text file and an Excel file** with correct formats.  
✅ Ensures the **Excel output follows CSR file column names**.

## Requirements

### 1. Install Python (if not installed)

Download and install Python from [Python.org](https://www.python.org/downloads/).  
Make sure Python is added to your system's PATH.

### 2. Install Dependencies

Run the following command to install the required Python packages:

```bash
pip install pandas openpyxl
```

## How to Use

1. **Place your Excel file** (e.g., `JANUARY 2025SALES.xlsx`) in the same directory as the script.
2. **Run the script** using the following command:

   ```bash
   python script.py
   ```

The script will generate two output files:

- `JANUARY 2025SALES.txt` → Fixed-width text file
- `JANUARY 2025SALES_formatted.xlsx` → Formatted Excel file
