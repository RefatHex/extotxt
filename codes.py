import pandas as pd
import os

def clean_text(value):
    """Removes unwanted newline characters and extra spaces."""
    return str(value).strip().replace("\n", "")

def detect_transaction_code(filename):
    """Determines the transaction type (S, P, I) based on the filename."""
    filename = filename.lower()
    if "sale" in filename:
        return "S"  # Sale
    elif "purchase" in filename:
        return "P"  # Purchase
    elif "inventory" in filename:
        return "I"  # Inventory
    return "S"  # Default to Sale if nothing matches

def get_last_day_of_month(df):
    """Determines the last date from the 'DATE' column in MMDDYYYY format."""
    df["DATE"] = pd.to_datetime(df["DATE"], errors='coerce')
    last_date = df["DATE"].max()
    return last_date.strftime('%m%d%Y') if pd.notnull(last_date) else " " * 8

def format_header(last_day):
    """Creates the control record (header) at the beginning of the file."""
    reporting_dea = "RY0658940"  # Fixed reporting registrant
    asterisk = "*"
    report_freq = "M"  # Monthly report
    central_reporter_dea = " " * 9  # 9 spaces (not used)
    return f"{reporting_dea}{asterisk}{last_day}{report_freq}{central_reporter_dea}"

def get_csr_columns():
    """Returns the expected column structure for the CSR file."""
    return [
        "YAVARI DEA", "Transaction Code", "ACTION INDICATOR", "NDC NUMBER (NO DASHES)",
        "QUANTITY", "UNIT", "ASSOCIATED REGISTRATION NUMBER", "ORDER FORM NUMBER",
        "TRANSACTION DATE", "CORRECTION NUMBER", "STRENGTH", "Transaction Number"
    ]

def format_transaction(row, transaction_code):
    """Formats a row of transaction data into an 80-character fixed-width format."""
    reporting_dea = "RY0658940"  # Fixed first column
    action_indicator = " "  # Blank (1 char)
    ndc_number = clean_text(row["NDC"]).replace("-", "").ljust(11)[:11]  # 11 chars, no dashes
    quantity = str(row["QUANTITY"]).rjust(8, '0')[:8]  # 8 chars, zero-padded
    unit = " "  # Blank (1 char)
    associate_dea = clean_text(row["DEA"]).ljust(9)[:9]  # 9 chars
    order_form = " " * 9  # 9 blank spaces (not used)
    transaction_date = pd.to_datetime(row["DATE"]).strftime('%m%d%Y')  # MMDDYYYY (8 chars)
    correction_number = " " * 8  # 8 blank spaces (not used)
    strength = " " * 4  # 4 blank spaces (not used)
    transaction_id = str(row.name + 1).rjust(10, '0')[:10]  # Unique ID (10 chars)
    filler = " "  # 1 blank space
    return f"{reporting_dea}{transaction_code}{action_indicator}{ndc_number}{quantity}{unit}{associate_dea}{order_form}{transaction_date}{correction_number}{strength}{transaction_id}{filler}"

def excel_to_text_and_excel(input_excel):
    df = pd.read_excel(input_excel, sheet_name="Report")
    
    # Clean text fields
    df = df.applymap(clean_text)
    
    # Detect transaction type from filename
    transaction_code = detect_transaction_code(input_excel)
    
    # Get the last day dynamically
    last_day = get_last_day_of_month(df)
    
    # Define output filenames based on user input
    base_name = os.path.splitext(input_excel)[0]
    output_text = f"{base_name}.txt"
    output_excel = f"{base_name}_formatted.xlsx"
    
    formatted_data = []
    
    with open(output_text, 'w') as txt_file:
        # Write header record
        header_line = format_header(last_day)
        txt_file.write(header_line + '\n')

        # Write transaction records
        for _, row in df.iterrows():
            formatted_line = format_transaction(row, transaction_code)
            txt_file.write(formatted_line + '\n')
            formatted_data.append([
                "RY0658940", transaction_code, "",  # DEA, Transaction Code, Action Indicator
                clean_text(row["NDC"]).replace("-", ""),  # NDC (no dashes)
                str(row["QUANTITY"]).rjust(8, '0')[:8],  # Quantity
                "",  # Unit
                clean_text(row["DEA"]),  # Associated DEA
                "",  # Order Form Number
                pd.to_datetime(row["DATE"]).strftime('%m%d%Y'),  # Transaction Date
                "", "",  # Correction Number, Strength
                str(row.name + 1).rjust(10, '0')[:10]  # Transaction Number
            ])
    
    # Create structured DataFrame and save to Excel
    formatted_df = pd.DataFrame(formatted_data, columns=get_csr_columns())
    formatted_df.to_excel(output_excel, index=False)
    
    print(f"Files successfully created: {output_text}, {output_excel}")

# Example usage
input_excel_file = "JANUARY 2025SALES.xlsx"
excel_to_text_and_excel(input_excel_file)
