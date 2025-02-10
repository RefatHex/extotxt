import streamlit as st
import pandas as pd
import os
import shutil

# Function Definitions

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
    ndc_number = clean_text(row.iloc[3]).replace("-", "").ljust(11)[:11]
    quantity = str(row.iloc[4]).rjust(8, '0')[:8]
    unit = " "
    associate_dea = clean_text(row.iloc[6]).ljust(9)[:9]
    order_form = " " * 9
    
    transaction_date = pd.to_datetime(row.iloc[8], errors='coerce')
    transaction_date_str = transaction_date.strftime('%m%d%Y') if pd.notnull(transaction_date) else " " * 8

    correction_number = " " * 8
    strength = " " * 4
    transaction_id = str(row.name + 1).rjust(10, '0')[:10]
    filler = " "
    
    return f"{reporting_dea}{transaction_code}{action_indicator}{ndc_number}{quantity}{unit}{associate_dea}{order_form}{transaction_date_str}{correction_number}{strength}{transaction_id}{filler}"

def excel_to_text(input_excel):
    try:
        df_dict = pd.read_excel(input_excel, sheet_name=None, header=None, skiprows=1)
        first_sheet = list(df_dict.keys())[0]
        df = df_dict[first_sheet]
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
        return None
    
    df = df.applymap(clean_text)
    last_day = get_last_day_of_month(df)
    
    output_text = f"{input_excel}.txt"

    with open(output_text, 'w') as txt_file:
        header_line = format_header(last_day)
        txt_file.write(header_line + '\n')

        for _, row in df.iterrows():
            formatted_line = format_transaction(row)
            txt_file.write(formatted_line + '\n')
    
    return output_text

# Streamlit App

def main():
    st.title("Excel to Text Converter")

    uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

    if uploaded_file is not None:
        os.makedirs("./temp", exist_ok=True)
        input_excel_path = f"./temp/{uploaded_file.name}"

        with open(input_excel_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        st.success(f"File '{uploaded_file.name}' uploaded successfully.")

        output_text = excel_to_text(input_excel_path)

        if output_text:
            st.subheader("Generated File")
            with open(output_text, "rb") as file:
                st.download_button("Download Text File", file, file_name=os.path.basename(output_text))

    st.subheader("Previously Uploaded and Generated Files")
    if os.path.exists("./temp"):
        previous_files = [f for f in os.listdir("./temp")]

        if previous_files:
            selected_file = st.selectbox("Choose a file", previous_files)

            if selected_file:
                file_path = os.path.join("./temp", selected_file)

                col1, col2 = st.columns(2)
                with col1:
                    with open(file_path, "rb") as file:
                        st.download_button(f"Download {selected_file}", file, file_name=selected_file)

                with col2:
                    if st.button(f"Delete {selected_file}"):
                        try:
                            if os.path.exists(file_path):
                                shutil.rmtree(file_path) if os.path.isdir(file_path) else os.remove(file_path)
                                st.success(f"File '{selected_file}' deleted successfully.")
                                st.rerun()
                            else:
                                st.error("File not found.")
                        except Exception as e:
                            st.error(f"Error deleting file: {e}")
        else:
            st.write("No previously uploaded or generated files found.")

if __name__ == "__main__":
    main()
