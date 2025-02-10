import pandas as pd
import os

def clean_text(value):
    """Removes unwanted newline characters and extra spaces."""
    return str(value).strip().replace("\n", "")

def detect_transaction_code(row):
    """Determines the transaction type based on the column 'Transaction Code'."""
    return clean_text(row.get("Transaction Code", "S"))  # Default to Sale (S) if missing

def get_last_day_of_month(df):
    """Determines the last date from the 'TRANSACTION DATE' column in MMDDYYYY format."""
    df["TRANSACTION DATE (MMDDYYYY) 8 DIGITS"] = pd.to_datetime(df["TRANSACTION DATE (MMDDYYYY) 8 DIGITS"], errors='coerce')
    last_date = df["TRANSACTION DATE (MMDDYYYY) 8 DIGITS"].max()
    return last_date.strftime('%m%d%Y') if pd.notnull(last_date) else " " * 8

def format_header(last_day):
    """Creates the control record (header) at the beginning of the file."""
    reporting_dea = "RY0658940"
    asterisk = "*"
    report_freq = "M"  # Monthly report
    central_reporter_dea = " " * 9
    return f"{reporting_dea}{asterisk}{last_day}{report_freq}{central_reporter_dea}"

def format_transaction(row):
    """Formats a row of transaction data into an 80-character fixed-width format."""
    reporting_dea = "RY0658940"
    transaction_code = detect_transaction_code(row)
    action_indicator = " "
    ndc_number = clean_text(row["NDC NUMBER (NO DASHES) (11 digits)"]).replace("-", "").ljust(11)[:11]
    quantity = str(row["QUANTITY"]).rjust(8, '0')[:8]
    unit = " "
    associate_dea = clean_text(row["DEA (9 digits)"]).ljust(9)[:9]
    order_form = " " * 9
    transaction_date = pd.to_datetime(row["TRANSACTION DATE (MMDDYYYY) 8 DIGITS"]).strftime('%m%d%Y')
    correction_number = " " * 8
    strength = " " * 4
    transaction_id = str(row.name + 1).rjust(10, '0')[:10]
    filler = " "
    return f"{reporting_dea}{transaction_code}{action_indicator}{ndc_number}{quantity}{unit}{associate_dea}{order_form}{transaction_date}{correction_number}{strength}{transaction_id}{filler}"

def excel_to_text_and_excel(input_excel):
    try:
        df = pd.read_excel(input_excel, sheet_name="Sheet1")
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return
    
    df = df.applymap(clean_text)
    last_day = get_last_day_of_month(df)
    
    base_name = os.path.splitext(input_excel)[0]
    output_text = f"{base_name}.txt"
    output_excel = f"{base_name}_formatted.xlsx"
    
    formatted_data = []
    
    with open(output_text, 'w') as txt_file:
        header_line = format_header(last_day)
        txt_file.write(header_line + '\n')

        for _, row in df.iterrows():
            formatted_line = format_transaction(row)
            txt_file.write(formatted_line + '\n')
            formatted_data.append([
                "RY0658940", detect_transaction_code(row), "",
                clean_text(row["NDC NUMBER (NO DASHES) (11 digits)"]).replace("-", ""),
                str(row["QUANTITY"]).rjust(8, '0')[:8],
                "",
                clean_text(row["DEA (9 digits)"]),
                "",
                pd.to_datetime(row["TRANSACTION DATE (MMDDYYYY) 8 DIGITS"]).strftime('%m%d%Y'),
                "", "",
                str(row.name + 1).rjust(10, '0')[:10]
            ])
    
    formatted_df = pd.DataFrame(formatted_data, columns=[
        "YAVARI DEA", "Transaction Code", "ACTION INDICATOR", "NDC NUMBER (NO DASHES)",
        "QUANTITY", "UNIT", "ASSOCIATED REGISTRATION NUMBER", "ORDER FORM NUMBER",
        "TRANSACTION DATE", "CORRECTION NUMBER", "STRENGTH", "Transaction Number"
    ])
    formatted_df.to_excel(output_excel, index=False)
    
    print(f"Files successfully created: {output_text}, {output_excel}")

if __name__ == "__main__":
    input_excel_file = input("Enter the Excel filename (including .xlsx extension): ").strip()
    if os.path.exists(input_excel_file):
        excel_to_text_and_excel(input_excel_file)
    else:
        print("File not found. Please check the filename and try again.")
