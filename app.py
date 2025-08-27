import streamlit as st
import pandas as pd
import re
from unidecode import unidecode
from ftfy import fix_text
import requests
from io import BytesIO
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill

# --- Google Sheets URL for unknown companies ---
UNKNOWN_COMPANY_LOG_URL = "https://script.google.com/macros/s/AKfycbxj8iwsHuSw3mmnsm0s72DsY51cKVy3K54DQOgcWaOgrhK706ZjFS_GlvPTdA8k-66N/exec"

# --- Global Styles ---
st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;700&display=swap');

        html, body, [class*="css"] {
            font-family: 'Montserrat', sans-serif;
            background-color: #ffffff;
            color: #0a2342;
        }
        .title-text {
            text-align: center;
            font-size: 3em;
            font-weight: 700;
            margin-bottom: 0.1em;
            color: #0a2342;
        }
        .subtitle-text {
            text-align: center;
            font-size: 1.2em;
            margin-bottom: 2em;
            color: #0a2342;
        }
        .rounded-box {
            background-color: #F0F0F0;
            border-radius: 16px;
            padding: 1.5rem;
            margin-bottom: 1.5rem;
            color: #0a2342;
        }
        .feedback-container {
            background-color: #F0F0F0;
            border-radius: 16px;
            padding: 1.5rem;
            margin-top: 2rem;
            color: #0a2342;
        }
        .feedback-box textarea {
            background-color: #112340;
            color: #ffffff;
            border-radius: 12px;
        }
        .download-button button, .stButton>button, .stDownloadButton>button {
            background-color: #112340 !important;
            color: #ffffff !important;
            border-radius: 12px;
        }
        .section-header {
            font-size: 1.3em;
            font-weight: 600;
            color: #0a2342;
            margin-bottom: 0.5em;
        }
    </style>
""", unsafe_allow_html=True)

# --- Load Company Directory ---
try:
    company_df = pd.read_csv("company_directory.csv")
    company_dict = dict(zip(company_df['Raw Company'].str.lower().str.strip(), company_df['Cleaned Company'].str.strip()))
except:
    company_dict = {}

# --- Cleaning Functions ---
COMMON_SUFFIXES = ['ltd', 'inc', 'group', 'brands', 'company', 'companies', 'incorporation', 'corporation']

def format_mc_name(name):
    return re.sub(r'\bMc([a-z])', lambda m: 'Mc' + m.group(1).upper(), name, flags=re.IGNORECASE)

def clean_company(name):
    if pd.isna(name): return ''
    name = str(name).strip()
    name_key = name.lower()

    if name_key in company_dict:
        return company_dict[name_key]

    # Fallback cleaning logic
    try:
        name = name.encode('latin1').decode('utf-8')
    except: pass
    name = fix_text(name)
    name = unidecode(name)
    name = re.sub(r'\b(?:' + '|'.join(COMMON_SUFFIXES) + r')\b', '', name, flags=re.IGNORECASE)
    name = re.sub(r'[^A-Za-z0-9\s\-]', '', name)
    name = re.sub(r'\s{2,}', ' ', name).strip()

    # Capitalize based on length
    cleaned_name = name.upper() if len(name) <= 4 else name.title()

    # Log unknown name to Google Sheets
    try:
        requests.post(
            UNKNOWN_COMPANY_LOG_URL,
            json={"sheet": "UnknownCompanies", "raw_company": name}
        )
    except:
        pass

    return cleaned_name

def clean_name(name, is_first=True):
    if pd.isna(name): return ''
    try:
        name = name.encode('latin1').decode('utf-8')
    except: pass
    name = fix_text(name)
    name = unidecode(str(name)).strip()
    name_parts = name.split()
    cleaned = name_parts[0] if is_first else name_parts[-1] if name_parts else ''
    return format_mc_name(cleaned.title())

def infer_from_email(first, last, email):
    if pd.isna(email): return first, last
    user = email.split('@')[0].lower()
    if len(last) <= 1:
        pattern = re.escape(first.lower()) + r'[._]?([a-z]+)'
        match = re.match(pattern, user)
        if match:
            guessed_last = match.group(1).title()
            return first, format_mc_name(guessed_last)
        if user.startswith(first[0].lower()):
            guess = user[len(first[0]):]
            return first, format_mc_name(guess.title()) if guess else last
    return first, last

def clean_data(df):
    cleaned_df = df.copy()
    changed_mask = pd.DataFrame(False, index=df.index, columns=df.columns)
    changes = 0

    for i, row in df.iterrows():
        orig_first, orig_last = str(row.get('First Name', '')).strip(), str(row.get('Last Name', '')).strip()
        orig_company = str(row.get('Company', '')).strip()
        email = str(row.get('Email', '')).strip() if 'Email' in df.columns else ''

        first, last = clean_name(orig_first, True), clean_name(orig_last, False)
        first, last = infer_from_email(first, last, email)
        company = clean_company(orig_company)

        if first != orig_first:
            cleaned_df.at[i, 'First Name'] = first
            changed_mask.at[i, 'First Name'] = True
        if last != orig_last:
            cleaned_df.at[i, 'Last Name'] = last
            changed_mask.at[i, 'Last Name'] = True
        if company != orig_company:
            cleaned_df.at[i, 'Company'] = company
            changed_mask.at[i, 'Company'] = True
            
        if first != orig_first or last != orig_last or company != orig_company:
            changes += 1

    pct = (changes / len(df)) * 100 if len(df) else 0
    return cleaned_df, pct, changed_mask

def generate_highlighted_excel(df, mask):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='openpyxl')
    df.to_excel(writer, index=False, sheet_name='CleanedData')
    workbook = writer.book
    worksheet = writer.sheets['CleanedData']

    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    for r_idx in range(2, len(df)+2):
        for c_idx, col in enumerate(df.columns):
            if mask.at[r_idx - 2, col]:
                cell = worksheet.cell(row=r_idx, column=c_idx+1)
                cell.fill = yellow_fill

    writer.save()
    output.seek(0)
    return output

# --- UI Logic ---
# (You can reinsert your existing Streamlit interface and usage tracking code here, unchanged.)
