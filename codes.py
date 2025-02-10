import pandas as pd
import os

def clean_text(value):
    return str(value).strip().replace("\n", "")

def detect_transaction_code(row):
    return clean_text(row.iloc[1] if len(row) > 1 else "S")  

def get_last_day_of_month(df):
    df.iloc[:, 8] = pd.to_datetime(df.iloc[:, 8], errors='coerce')
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
    
    transaction_date = pd.to_datetime(row.iloc[8], errors='coerce')  # TRANSACTION DATE in column 8
    transaction_date_str = transaction_date.strftime('%m%d%Y') if pd.notnull(transaction_date) else " " * 8

    correction_number = " " * 8
    strength = " " * 4
    transaction_id = str(row.name + 1).rjust(10, '0')[:10]
    filler = " "
    
    return f"{reporting_dea}{transaction_code}{action_indicator}{ndc_number}{quantity}{unit}{associate_dea}{order_form}{transaction_date_str}{correction_number}{strength}{transaction_id}{filler}"

def excel_to_text(input_excel):
    try:
        # Read the only available sheet dynamically
        df = pd.read_excel(input_excel, sheet_name=None, header=None, skiprows=1)
        sheet_name = list(df.keys())[0]  # Get the first available sheet
        df = df[sheet_name]
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return
    
    df = df.applymap(clean_text)
    last_day = get_last_day_of_month(df)
    
    base_name = os.path.splitext(input_excel)[0]
    output_text = f"{base_name}.txt"
    
    with open(output_text, 'w') as txt_file:
        header_line = format_header(last_day)
        txt_file.write(header_line + '\n')
        
        for _, row in df.iterrows():
            formatted_line = format_transaction(row)
            txt_file.write(formatted_line + '\n')
    
    print(f"File successfully created: {output_text}")

if __name__ == "__main__":
    input_excel_file = input("Enter the Excel filename (including .xlsx extension): ").strip()
    if os.path.exists(input_excel_file):
        excel_to_text(input_excel_file)
    else:
        print("File not found. Please check the filename and try again.")
