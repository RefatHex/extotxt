import pandas as pd
import os

def clean_text(value):
    return str(value).strip().replace("\n", "")

def detect_transaction_code(row):
    return clean_text(row.iloc[1] if len(row) > 1 else "S")  # Assuming Transaction Code is in column 1

def get_last_day_of_month(df):
    # Ensure TRANSACTION DATE (column 8) is properly converted to datetime
    df.iloc[:, 8] = pd.to_datetime(df.iloc[:, 8], errors='coerce')
    
    # Drop rows where TRANSACTION DATE is NaT
    valid_dates = df.iloc[:, 8].dropna()
    
    if not valid_dates.empty:
        last_date = valid_dates.max()
        return last_date.strftime('%m%d%Y') if pd.notnull(last_date) else " " * 8
    else:
        return " " * 8

def format_header(last_day):
    reporting_dea = "RY0658940"
    asterisk = "*"
    report_freq = "M"
    central_reporter_dea = " " * 9
    return f"{reporting_dea}{asterisk}{last_day}{report_freq}{central_reporter_dea}"

def format_transaction(row):
    reporting_dea = "RY0658940"
    transaction_code = detect_transaction_code(row)
    action_indicator = " "
    ndc_number = clean_text(row.iloc[3]).replace("-", "").ljust(11)[:11]  # NDC in column 3
    quantity = str(row.iloc[4]).rjust(8, '0')[:8]  # QUANTITY in column 4
    unit = " "
    associate_dea = clean_text(row.iloc[6]).ljust(9)[:9]  # DEA in column 6
    order_form = " " * 9

    # Convert TRANSACTION DATE safely
    transaction_date = pd.to_datetime(row.iloc[8], errors='coerce')  # TRANSACTION DATE in column 8
    transaction_date_str = transaction_date.strftime('%m%d%Y') if pd.notnull(transaction_date) else " " * 8

    correction_number = " " * 8
    strength = " " * 4
    transaction_id = str(row.name + 1).rjust(10, '0')[:10]
    filler = " "

    return f"{reporting_dea}{transaction_code}{action_indicator}{ndc_number}{quantity}{unit}{associate_dea}{order_form}{transaction_date_str}{correction_number}{strength}{transaction_id}{filler}"

def excel_to_text_and_excel(input_excel):
    try:
        # Read Excel file, skipping the first row (header)
        df = pd.read_excel(input_excel, sheet_name="Sheet1", header=None, skiprows=1)
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return
    
    df = df.applymap(clean_text)  # Clean text fields
    last_day = get_last_day_of_month(df)  # Get last day of month
    
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
                clean_text(row.iloc[3]).replace("-", ""),  # NDC column
                str(row.iloc[4]).rjust(8, '0')[:8],  # Quantity column
                "",
                clean_text(row.iloc[6]),  # DEA column
                "",
                pd.to_datetime(row.iloc[8], errors='coerce').strftime('%m%d%Y') if pd.notnull(pd.to_datetime(row.iloc[8], errors='coerce')) else " " * 8,
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
