import streamlit as st
import pandas as pd
import re
from io import BytesIO

# ✅ Must be first
st.set_page_config(
    page_title="Form 10BD | Data Cleaner",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ---- Custom CSS ----
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap');

html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }
.stApp { background: #F4F7F0; }
#MainMenu, footer, header { visibility: hidden; }
.block-container { padding: 0 !important; max-width: 100% !important; }

/* ── NAVBAR ── */
.navbar {
    background: #2D4A1E;
    padding: 0 48px;
    height: 64px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    position: sticky;
    top: 0;
    z-index: 999;
    box-shadow: 0 2px 12px rgba(0,0,0,0.20);
}
.navbar-brand { display: flex; align-items: center; gap: 14px; }
.navbar-logo-box {
    width: 38px; height: 38px;
    background: #7cb452;
    border-radius: 8px;
    display: flex; align-items: center; justify-content: center;
    font-size: 19px; font-weight: 700; color: #fff;
    font-family: 'DM Mono', monospace;
    box-shadow: 0 2px 8px rgba(124,180,82,0.4);
}
.navbar-title  { color: #FFFFFF; font-size: 17px; font-weight: 600; }
.navbar-subtitle { color: #9DC878; font-size: 12px; font-weight: 400; letter-spacing: 0.5px; text-transform: uppercase; }
.navbar-badge {
    background: #7cb452; color: #fff;
    font-size: 11px; font-weight: 600;
    padding: 4px 14px; border-radius: 20px; letter-spacing: 0.3px;
}

/* ── BODY ── */
.page-body { padding: 40px 48px 60px 48px; }

.section-label {
    font-size: 11px; font-weight: 600;
    letter-spacing: 1.5px; text-transform: uppercase;
    color: #5a9a30; margin-bottom: 8px;
}
.section-title  { font-size: 26px; font-weight: 600; color: #1E3510; margin-bottom: 4px; }
.section-desc   { font-size: 14px; color: #6B7A60; margin-bottom: 28px; }

/* ── INFO BOX ── */
.info-box {
    background: #EFF6E8;
    border: 1px solid #B8D99A;
    border-radius: 10px;
    padding: 14px 20px;
    font-size: 13px; color: #2D4A1E;
    margin-bottom: 24px; line-height: 1.7;
}

/* ── STAT CARDS ── */
.stats-row { display: flex; gap: 16px; margin-bottom: 28px; flex-wrap: wrap; }
.stat-card {
    background: #fff; border-radius: 12px; padding: 20px 26px;
    flex: 1; min-width: 150px;
    border: 1px solid #D8E8CC;
    box-shadow: 0 1px 4px rgba(0,0,0,0.05);
}
.stat-label { font-size: 11px; font-weight: 600; letter-spacing: 1px; text-transform: uppercase; color: #9DB890; margin-bottom: 6px; }
.stat-value { font-size: 30px; font-weight: 600; font-family: 'DM Mono', monospace; line-height: 1; color: #1E3510; }
.stat-value.green  { color: #7cb452; }
.stat-value.dgreen { color: #4A8A20; }
.stat-value.orange { color: #D4831E; }
.stat-value.red    { color: #C94040; }

/* ── UPLOAD PLACEHOLDER ── */
.upload-card {
    background: #fff; border: 2px dashed #B8D99A; border-radius: 16px;
    padding: 52px 32px; text-align: center;
}
.upload-icon  { font-size: 44px; margin-bottom: 12px; }
.upload-title { font-size: 17px; font-weight: 600; color: #1E3510; margin-bottom: 6px; }
.upload-hint  { font-size: 13px; color: #9DB890; }

/* ── BUTTONS ── */
.stButton > button {
    background: #7cb452 !important; color: #fff !important;
    border: none !important; border-radius: 10px !important;
    padding: 14px 40px !important;
    font-family: 'DM Sans', sans-serif !important;
    font-size: 15px !important; font-weight: 600 !important;
    box-shadow: 0 4px 14px rgba(124,180,82,0.35) !important;
    width: 100% !important; transition: background 0.2s !important;
}
.stButton > button:hover { background: #5f9a36 !important; }

[data-testid="stDownloadButton"] > button {
    background: #4A8A20 !important; color: #fff !important;
    border: none !important; border-radius: 10px !important;
    padding: 14px 40px !important;
    font-family: 'DM Sans', sans-serif !important;
    font-size: 15px !important; font-weight: 600 !important;
    box-shadow: 0 4px 14px rgba(74,138,32,0.30) !important;
    width: 100% !important; transition: background 0.2s !important;
}
[data-testid="stDownloadButton"] > button:hover { background: #3A6E18 !important; }

/* ── DATAFRAME ── */
[data-testid="stDataFrame"] {
    border-radius: 12px !important; overflow: hidden !important;
    border: 1px solid #D8E8CC !important;
    box-shadow: 0 2px 8px rgba(0,0,0,0.05) !important;
}

/* ── ALERTS ── */
.stSuccess > div { background: #EFF6E8 !important; border-left: 4px solid #7cb452 !important; border-radius: 8px !important; color: #1E3510 !important; }
.stError   > div { background: #FDECEA !important; border-left: 4px solid #C94040 !important; border-radius: 8px !important; }
.stWarning > div { background: #FEF3E7 !important; border-left: 4px solid #D4831E !important; border-radius: 8px !important; }
.stInfo    > div { background: #EFF6E8 !important; border-left: 4px solid #5a9a30 !important; border-radius: 8px !important; color: #2D4A1E !important; }

/* ── DIVIDER ── */
.divider { height: 1px; background: #D8E8CC; margin: 28px 0; }

/* ── RESULTS HEADER ── */
.results-header { display: flex; align-items: center; justify-content: space-between; margin-bottom: 16px; }
.results-title  { font-size: 18px; font-weight: 600; color: #1E3510; }
.results-count  { font-size: 13px; color: #6B7A60; background: #EFF6E8; border-radius: 20px; padding: 4px 14px; font-family: 'DM Mono', monospace; }

/* ── EXPORT BOX ── */
.export-box {
    background: #EFF6E8; border: 1px solid #B8D99A; border-radius: 10px;
    padding: 16px 20px; font-size: 13px; color: #2D4A1E; line-height: 2;
}

/* ── COLUMNS INFO ── */
.col-info-box {
    background: #fff; border: 1px solid #D8E8CC; border-radius: 12px;
    padding: 20px; font-size: 13px; color: #3A4A30; line-height: 2.3;
}

/* ── FOOTER ── */
.footer {
    background: #2D4A1E; padding: 20px 48px;
    display: flex; align-items: center; justify-content: space-between;
    margin-top: 60px;
}
.footer-text  { color: #6A8A58; font-size: 13px; }
.footer-brand { color: #9DC878; font-size: 13px; font-weight: 500; }

[data-testid="column"] { padding: 0 8px !important; }
</style>
""", unsafe_allow_html=True)


# ============================================================
#  CORE LOGIC
# ============================================================

def strip_special_chars(text):
    """
    Removes ALL characters except letters (A-Z a-z), digits (0-9),
    and spaces. This applies to every column except the ones that
    have their own dedicated handlers (dates, amounts, UID).
    """
    if pd.isna(text):
        return ''
    return re.sub(r'[^A-Za-z0-9 ]', '', str(text)).strip()


def clean_uid(uid_str):
    """Strip everything except alphanumeric from a UID string."""
    return re.sub(r'[^A-Za-z0-9]', '', str(uid_str))


def validate_and_correct(row):
    uid      = row['Unique Identification Number']
    id_code  = str(row['ID Code']).strip().title()
    change_note = ''

    if pd.isna(uid) or str(uid).strip().lower() in ('not available', 'na', 'n/a', ''):
        uid_clean    = 'NNNNN0000N'
        correct_code = 'Permanent Account Number'
        change_note  = 'Filled default PAN for missing UID'
    else:
        uid       = str(uid).strip()
        uid_clean = clean_uid(uid)
        correct_code = id_code
        is_valid  = False

        if uid_clean.isdigit() and len(uid_clean) == 12:
            correct_code = 'Aadhaar Number'
            is_valid     = True
            uid_clean    = int(uid_clean)          # numeric for Aadhaar
        elif re.fullmatch(r'[A-Z]{5}[0-9]{4}[A-Z]', uid_clean, re.IGNORECASE):
            correct_code = 'Permanent Account Number'
            is_valid     = True
            uid_clean    = uid_clean.upper()
        elif re.fullmatch(r'[A-Za-z0-9]{8,10}', uid_clean):
            correct_code = 'Passport Number'
            is_valid     = True
        elif re.fullmatch(r'[A-Za-z]{2}[0-9]{11,13}', uid_clean):
            correct_code = 'Driving Licence'
            is_valid     = True
        else:
            if uid_clean.isdigit():
                uid_clean = int(uid_clean)

        if not is_valid:
            change_note = 'Invalid UID Format - Needs Review'
        elif id_code != correct_code:
            change_note = 'ID Code mismatch'
        elif str(uid) != str(uid_clean):
            change_note = 'Formatted UID'

    return pd.Series([id_code, uid_clean, change_note])


def format_date(date_value):
    if pd.isna(date_value):
        return ''
    try:
        parsed = pd.to_datetime(date_value, errors='coerce')
        return parsed.strftime('%d-%b-%Y') if not pd.isna(parsed) else ''
    except Exception:
        return ''


def convert_to_numeric(value):
    if pd.isna(value):
        return ''
    try:
        return int(float(str(value).replace(',', '').strip()))
    except Exception:
        return value


def to_excel_download(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        export_df = df.copy()

        # Ensure Aadhaar stored as integer
        for i, val in enumerate(export_df['Unique Identification Number']):
            if isinstance(val, str) and val.isdigit():
                export_df.loc[export_df.index[i], 'Unique Identification Number'] = int(val)

        export_df.to_excel(writer, sheet_name='CleanedData', index=False, startrow=1, header=False)
        workbook  = writer.book
        worksheet = writer.sheets['CleanedData']

        # Styled header
        hdr_fmt = workbook.add_format({
            'bold': True, 'bg_color': '#2D4A1E', 'font_color': '#FFFFFF',
            'font_name': 'Calibri', 'font_size': 10, 'border': 0
        })
        for col_num, col_name in enumerate(export_df.columns.values):
            worksheet.write(0, col_num, col_name, hdr_fmt)

        uid_idx    = export_df.columns.get_loc('Unique Identification Number')
        num_fmt    = workbook.add_format({'num_format': '0'})
        txt_fmt    = workbook.add_format()

        for row_num, val in enumerate(export_df['Unique Identification Number'], start=1):
            try:
                if isinstance(val, (int, float)) or (isinstance(val, str) and val.isdigit()):
                    worksheet.write_number(row_num, uid_idx, int(float(str(val).replace(',', ''))), num_fmt)
                else:
                    worksheet.write(row_num, uid_idx, val, txt_fmt)
            except Exception:
                worksheet.write(row_num, uid_idx, str(val), txt_fmt)

    output.seek(0)
    return output.read()


def process_dataframe(df):
    df.columns = df.columns.str.strip()

    for col in ['ID Code', 'Unique Identification Number']:
        if col not in df.columns:
            st.error(f"Missing required column: **{col}**")
            st.stop()

    # ── Step 1: UID validation ──
    df[['ID Code', 'Unique Identification Number', 'Change Note']] = \
        df.apply(validate_and_correct, axis=1)

    # ── Step 2: Date formatting ──
    date_col = 'Date of Issuance of Unique Registration Number'
    if date_col in df.columns:
        df[date_col] = df[date_col].apply(format_date)

    # ── Step 3: Amount conversion ──
    amt_col = 'Amount of donation (Indian rupees)'
    if amt_col in df.columns:
        df[amt_col] = df[amt_col].apply(convert_to_numeric)

    # ── Step 4: Strip ALL special characters from every remaining text column ──
    # Columns that must NOT go through strip_special_chars:
    protected = {
        'Unique Identification Number',   # already cleaned
        'ID Code',                         # already cleaned
        'Change Note',                     # internal flag
        date_col,                          # formatted date string
        amt_col,                           # numeric
    }
    for col in df.columns:
        if col not in protected and df[col].dtype == object:
            df[col] = df[col].apply(strip_special_chars)

    return df


# ============================================================
#  NAVBAR
# ============================================================
st.markdown("""
<div class="navbar">
    <div class="navbar-brand">
        <div class="navbar-logo-box">K</div>
        <div>
            <div class="navbar-title">KKC &amp; Associates LLP</div>
            <div class="navbar-subtitle">Form 10BD Data Cleaner</div>
        </div>
    </div>
    <div class="navbar-badge">IT Tool v2.0</div>
</div>
""", unsafe_allow_html=True)

# ============================================================
#  PAGE BODY
# ============================================================
st.markdown('<div class="page-body">', unsafe_allow_html=True)

st.markdown("""
<div class="section-label">Compliance Tool</div>
<div class="section-title">Form 10BD Data Cleaner</div>
<div class="section-desc">
    Upload your donor data Excel file. The tool will validate all UIDs, correct ID codes,
    strip special characters from every column, and export a clean filing-ready sheet.
</div>
<div class="info-box">
    <strong>Cleaning rules applied:</strong> &nbsp;
    All special characters (, . / - ' &amp; @ # etc.) are removed from every text column &nbsp;·&nbsp;
    Aadhaar (12-digit), PAN (AAAAA0000A), Passport &amp; Driving Licence validated &nbsp;·&nbsp;
    ID code mismatches flagged &nbsp;·&nbsp;
    Dates formatted as DD-Mon-YYYY &nbsp;·&nbsp;
    Donation amounts converted to integers.
</div>
""", unsafe_allow_html=True)

# ---- Upload Row ----
col_up, col_info = st.columns([2, 1], gap="large")

with col_up:
    st.markdown('<div class="section-label">Step 1 — Upload File</div>', unsafe_allow_html=True)
    uploaded_file = st.file_uploader(
        "Drop your Form 10BD Excel file here",
        type=["xlsx"],
        label_visibility="collapsed"
    )

with col_info:
    st.markdown("""
    <div class="section-label">Expected Columns</div>
    <div class="col-info-box">
        ✦ &nbsp;<b>ID Code</b><br>
        ✦ &nbsp;<b>Unique Identification Number</b><br>
        ✦ &nbsp;Name of donor<br>
        ✦ &nbsp;Address of donor<br>
        ✦ &nbsp;Date of Issuance of Unique Registration Number<br>
        ✦ &nbsp;Amount of donation (Indian rupees)<br>
        ✦ &nbsp;Mode of receipt &nbsp;/&nbsp; Donation Type
    </div>
    """, unsafe_allow_html=True)

# ============================================================
#  MAIN LOGIC
# ============================================================
if uploaded_file:
    try:
        df_raw     = pd.read_excel(uploaded_file)
        total_rows = len(df_raw)
        total_cols = len(df_raw.columns)

        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
        st.markdown('<div class="section-label">Step 2 — Preview &amp; Process</div>', unsafe_allow_html=True)

        st.markdown(f"""
        <div class="stats-row">
            <div class="stat-card">
                <div class="stat-label">Total Rows</div>
                <div class="stat-value green">{total_rows:,}</div>
            </div>
            <div class="stat-card">
                <div class="stat-label">Columns Detected</div>
                <div class="stat-value">{total_cols}</div>
            </div>
            <div class="stat-card">
                <div class="stat-label">Uploaded File</div>
                <div class="stat-value" style="font-size:13px;padding-top:6px;
                     font-family:'DM Sans',sans-serif;color:#7cb452;">
                     {uploaded_file.name}
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        col_btn, _ = st.columns([1, 1], gap="large")
        with col_btn:
            process_clicked = st.button("⚙️  Process & Clean Data", use_container_width=True)

        if process_clicked:
            with st.spinner("Validating UIDs and stripping special characters…"):
                processed_df = process_dataframe(df_raw.copy())
                st.session_state["processed_df"] = processed_df
                st.session_state["total_rows"]    = total_rows

        if "processed_df" in st.session_state:
            processed_df = st.session_state["processed_df"]
            t_rows       = st.session_state.get("total_rows", total_rows)
            change_col   = 'Change Note' if 'Change Note' in processed_df.columns else None
            flagged      = int((processed_df[change_col] != '').sum()) if change_col else 0
            invalid      = int(processed_df[change_col].str.contains('Invalid', na=False).sum()) if change_col else 0
            clean_rows   = t_rows - flagged

            st.markdown(f"""
            <div class="stats-row">
                <div class="stat-card">
                    <div class="stat-label">Clean Records</div>
                    <div class="stat-value dgreen">{clean_rows:,}</div>
                </div>
                <div class="stat-card">
                    <div class="stat-label">Flagged / Corrected</div>
                    <div class="stat-value orange">{flagged:,}</div>
                </div>
                <div class="stat-card">
                    <div class="stat-label">Needs Manual Review</div>
                    <div class="stat-value red">{invalid:,}</div>
                </div>
                <div class="stat-card">
                    <div class="stat-label">Total Processed</div>
                    <div class="stat-value green">{t_rows:,}</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            st.success(f"✅  {t_rows:,} records processed — special characters removed from all columns.")

            # ---- Export ----
            st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
            st.markdown('<div class="section-label">Step 3 — Export</div>', unsafe_allow_html=True)

            col_dl1, col_dl2 = st.columns([1, 1], gap="large")
            with col_dl1:
                excel_bytes = to_excel_download(processed_df)
                st.download_button(
                    label="📥  Download Cleaned Excel File",
                    data=excel_bytes,
                    file_name="Form10BD_Cleaned.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            with col_dl2:
                st.markdown("""
                <div class="export-box">
                    <b>What's inside the export:</b><br>
                    ✔ &nbsp;All special characters removed from every column<br>
                    ✔ &nbsp;Validated &amp; corrected UID numbers<br>
                    ✔ &nbsp;Corrected ID code labels<br>
                    ✔ &nbsp;Dates formatted as DD-Mon-YYYY<br>
                    ✔ &nbsp;Change notes column for audit trail
                </div>
                """, unsafe_allow_html=True)

            # ---- Data Table ----
            st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
            st.markdown(f"""
            <div class="results-header">
                <div class="results-title">Cleaned Data Preview</div>
                <div class="results-count">{len(processed_df):,} records</div>
            </div>
            """, unsafe_allow_html=True)

            f_col, _ = st.columns([1, 3], gap="large")
            with f_col:
                show_filter = st.selectbox(
                    "Filter view",
                    ["All Records", "Flagged / Corrected Only", "Needs Review Only"],
                    label_visibility="collapsed"
                )

            display_df = processed_df.copy()
            if show_filter == "Flagged / Corrected Only" and change_col:
                display_df = display_df[display_df[change_col] != '']
            elif show_filter == "Needs Review Only" and change_col:
                display_df = display_df[display_df[change_col].str.contains('Invalid', na=False)]

            st.dataframe(display_df, use_container_width=True, height=500)

    except Exception as e:
        st.error(f"❌  Error reading file: {str(e)}")

else:
    st.markdown("""
    <div class="upload-card">
        <div class="upload-icon">📂</div>
        <div class="upload-title">No file uploaded yet</div>
        <div class="upload-hint">Use the uploader above to select your Form 10BD Excel file (.xlsx)</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)

# ============================================================
#  FOOTER
# ============================================================
st.markdown("""
<div class="footer">
    <div class="footer-text">© 2025 KKC &amp; Associates LLP · All rights reserved</div>
    <div class="footer-brand">Designed by IT Team &nbsp;·&nbsp; Form 10BD Compliance Suite</div>
</div>
""", unsafe_allow_html=True)