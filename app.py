import streamlit as st
import pandas as pd
import os

# Function Definitions (same as your provided code)

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
    reporting_dea = "RY0658940"
    action_indicator = " "
    ndc_number = clean_text(row["NDC"]).replace("-", "").ljust(11)[:11]
    quantity = str(row["QUANTITY"]).rjust(8, '0')[:8]
    unit = " "
    associate_dea = clean_text(row["DEA"]).ljust(9)[:9]
    order_form = " " * 9
    transaction_date = pd.to_datetime(row["DATE"]).strftime('%m%d%Y')
    correction_number = " " * 8
    strength = " " * 4
    transaction_id = str(row.name + 1).rjust(10, '0')[:10]
    filler = " "
    return f"{reporting_dea}{transaction_code}{action_indicator}{ndc_number}{quantity}{unit}{associate_dea}{order_form}{transaction_date}{correction_number}{strength}{transaction_id}{filler}"

def excel_to_text_and_excel(input_excel):
    try:
        df = pd.read_excel(input_excel, sheet_name="Report")
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
        return
    
    df = df.applymap(clean_text)
    transaction_code = detect_transaction_code(input_excel)
    last_day = get_last_day_of_month(df)
    
    base_name = os.path.splitext(input_excel)[0]
    output_text = f"{base_name}.txt"
    output_excel = f"{base_name}_formatted.xlsx"
    
    formatted_data = []
    
    with open(output_text, 'w') as txt_file:
        header_line = format_header(last_day)
        txt_file.write(header_line + '\n')

        for _, row in df.iterrows():
            formatted_line = format_transaction(row, transaction_code)
            txt_file.write(formatted_line + '\n')
            formatted_data.append([
                "RY0658940", transaction_code, "",
                clean_text(row["NDC"]).replace("-", ""),
                str(row["QUANTITY"]).rjust(8, '0')[:8],
                "",
                clean_text(row["DEA"]),
                "",
                pd.to_datetime(row["DATE"]).strftime('%m%d%Y'),
                "", "",
                str(row.name + 1).rjust(10, '0')[:10]
            ])
    
    formatted_df = pd.DataFrame(formatted_data, columns=get_csr_columns())
    formatted_df.to_excel(output_excel, index=False)
    
    return output_text, output_excel

# Streamlit App

def main():
    st.title("Excel to Text and Excel Converter")

    # Upload Excel File Section
    uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

    # File selection dropdown for previous files
    col1, col2 = st.columns([1, 4])  # Adjusted columns for better layout

    with col2:
        st.subheader("Select Previous Files")
        if os.path.exists("./temp"):
            # List of previously uploaded files (Text or Excel)
            existing_files = os.listdir("./temp")
            previous_files = [f for f in existing_files if f.endswith(".txt") or f.endswith(".xlsx")]
            previous_file = st.selectbox("Choose a file", previous_files)

            # Show the selected previous file with download options
            if previous_file:
                file_path = os.path.join("./temp", previous_file)
                with open(file_path, "rb") as file:
                    st.download_button(f"Download {previous_file}", file, file_name=previous_file)

    if uploaded_file is not None:
        input_excel_path = f"./temp/{uploaded_file.name}"
        os.makedirs(os.path.dirname(input_excel_path), exist_ok=True)

        with open(input_excel_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        st.success(f"File '{uploaded_file.name}' uploaded successfully.")

        # Convert file
        output_text, output_excel = excel_to_text_and_excel(input_excel_path)

        # Display generated files immediately below
        st.subheader("Generated Files")
        st.write(f"Text File: {output_text}")
        with open(output_text, "rb") as file:
            st.download_button("Download Text File", file, file_name=output_text)

        st.write(f"Excel File: {output_excel}")
        with open(output_excel, "rb") as file:
            st.download_button("Download Excel File", file, file_name=output_excel)

if __name__ == "__main__":
    main()
