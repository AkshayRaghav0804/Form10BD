import streamlit as st
import pandas as pd
import re
from pathlib import Path
from io import BytesIO

# ✅ Must be first
st.set_page_config(page_title="Form10BD Data Cleaner", layout="wide")

# ---- Logo Display ----
ASSETS_DIR = Path('./assets')
logo_path = ASSETS_DIR / 'kkc logo.png'

if logo_path.exists():
    st.image(str(logo_path), width=375)
else:
    st.warning("Logo file not found. Please place 'kkc logo.png' in the 'assets' directory.")

# ---- Validation Function ----
def validate_and_correct(row):
    uid = row['Unique Identification Number']
    id_code = str(row['ID Code']).strip().title()
    change_note = ''
    
    if pd.isna(uid) or str(uid).strip().lower() == 'not available':
        uid_clean = 'NNNNN0000N'
        correct_code = 'Permanent Account Number'
        change_note = "Filled default PAN for missing UID"
    else:
        uid = str(uid).strip()
        uid_clean = re.sub(r'[^A-Za-z0-9]', '', uid)
        correct_code = id_code
        is_valid = False

        if uid_clean.isdigit() and len(uid_clean) == 12:
            correct_code = 'Aadhaar Number'
            is_valid = True
            # Convert to numeric for Aadhaar
            uid_clean = int(uid_clean)
        elif re.fullmatch(r'[A-Z]{5}[0-9]{4}[A-Z]', uid_clean, re.IGNORECASE):
            correct_code = 'Permanent Account Number'
            is_valid = True
            # Keep as string for PAN
        elif re.fullmatch(r'[A-Za-z0-9]{8,10}', uid_clean):
            correct_code = 'Passport Number'
            is_valid = True
            # Keep as string for Passport
        elif re.fullmatch(r'[A-Za-z]{2}[0-9]{11,13}', uid_clean):
            correct_code = 'Driving Licence'
            is_valid = True
            # Keep as string for Driving License
        else:
            # Try to convert to numeric if it's a pure number
            if uid_clean.isdigit():
                uid_clean = int(uid_clean)

        if not is_valid:
            change_note = "Invalid UID Format - Needs Review"
        elif id_code != correct_code:
            change_note = "ID Code mismatch"
        elif uid != uid_clean:
            change_note = "Formatted UID"

    # Do not change the ID Code here, just flag it in the change note
    row['Change Note'] = change_note
    row['Unique Identification Number'] = uid_clean

    return pd.Series([row['ID Code'], row['Unique Identification Number'], change_note])

# ---- Helpers ----
def clean_text(text):
    if pd.isna(text):
        return ''
    return re.sub(r'[^\w\s]', '', str(text)).strip()[:100]

def format_date(date_value):
    if pd.isna(date_value):
        return ''
    try:
        parsed = pd.to_datetime(date_value, errors='coerce')
        return parsed.strftime('%d-%b-%Y') if not pd.isna(parsed) else ''
    except:
        return ''

def convert_to_numeric(value):
    if pd.isna(value):
        return ''
    try:
        return int(float(str(value).replace(',', '').strip()))
    except:
        return value

# --- Improved Excel Writer with better handling of numeric values ---
def to_excel_download(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Clone the dataframe to avoid modifying the original
        export_df = df.copy()
        
        # Convert numeric strings to actual numbers for Excel
        uid_col = export_df['Unique Identification Number']
        for i, val in enumerate(uid_col):
            if isinstance(val, str) and val.isdigit():
                export_df.loc[export_df.index[i], 'Unique Identification Number'] = int(val)
        
        # Write to Excel
        export_df.to_excel(writer, sheet_name='CleanedData', index=False, startrow=1, header=False)
        workbook = writer.book
        worksheet = writer.sheets['CleanedData']
        
        # Write headers
        for col_num, value in enumerate(export_df.columns.values):
            worksheet.write(0, col_num, value)
        
        # Get the column index for UID
        uid_col_index = export_df.columns.get_loc('Unique Identification Number')
        
        # Apply custom formatting for each cell in the UID column
        number_format = workbook.add_format({'num_format': '0'})
        text_format = workbook.add_format()
        
        for row_num, value in enumerate(export_df['Unique Identification Number'], start=1):
            try:
                if isinstance(value, (int, float)) or (isinstance(value, str) and value.isdigit()):
                    # Convert to integer and write as number
                    numeric_value = int(float(str(value).replace(',', '')))
                    worksheet.write_number(row_num, uid_col_index, numeric_value, number_format)
                else:
                    # Keep as text for alphanumeric values
                    worksheet.write(row_num, uid_col_index, value, text_format)
            except Exception as e:
                # Fallback to plain text if conversion fails
                worksheet.write(row_num, uid_col_index, str(value), text_format)

    output.seek(0)
    return output.read()

# ---- Main Processor ----
def process_dataframe(df):
    df.columns = df.columns.str.strip()

    required_cols = ['ID Code', 'Unique Identification Number']
    for col in required_cols:
        if col not in df.columns:
            st.error(f"Missing required column: {col}")
            st.stop()

    df[['ID Code', 'Unique Identification Number', 'Change Note']] = df.apply(validate_and_correct, axis=1)

    # Clean special characters from all columns except exclusions
    excluded_cols = ['Date of Issuance of Unique Registration Number', 'Mode of receipt']
    for col in df.columns:
        if col not in excluded_cols and df[col].dtype == object:
            df[col] = df[col].apply(clean_text)

    if 'Date of Issuance of Unique Registration Number' in df.columns:
        df['Date of Issuance of Unique Registration Number'] = df['Date of Issuance of Unique Registration Number'].apply(format_date)

    if 'Amount of donation (Indian rupees)' in df.columns:
        df['Amount of donation (Indian rupees)'] = df['Amount of donation (Indian rupees)'].apply(convert_to_numeric)

    return df

# ---- Streamlit UI ----
st.title("Form10BD Data Cleaner Tool")

uploaded_file = st.file_uploader("Upload your Form10BD Excel file (.xlsx)", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        st.success("✅ File uploaded successfully!")

        if st.button("Process File"):
            with st.spinner("Processing..."):
                processed_df = process_dataframe(df)
                st.success("✅ Processing completed!")

                st.write("### Full Cleaned Data:")
                st.dataframe(processed_df, use_container_width=True)

                excel_data = to_excel_download(processed_df)

                st.download_button(
                    label="📥 Download Cleaned Excel File",
                    data=excel_data,
                    file_name="Form10BD_Cleaned.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    except Exception as e:
        st.error(f"❌ Error: {str(e)}")

st.info("Designed by KKC & ASSOCIATES LLP - IT Team")
