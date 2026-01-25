import streamlit as st
import pandas as pd
import re
from unidecode import unidecode
from ftfy import fix_text
import os
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils.dataframe import dataframe_to_rows
import random

# --- Page Config ---
st.set_page_config(page_title="Cleanr", layout="wide", page_icon="🧹")

# --- Custom UI Styling (Canva Re-creation) ---
st.markdown("""
    <style>
        /* Import Montserrat */
        @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600;700&display=swap');

        /* Main Grid Background */
        .stApp {
            background-color: #ffffff;
            background-image: 
                linear-gradient(#f0f0f0 1px, transparent 1px),
                linear-gradient(90deg, #f0f0f0 1px, transparent 1px);
            background-size: 30px 30px;
            font-family: 'Montserrat', sans-serif;
        }

        /* Sidebar Styling */
        section[data-testid="stSidebar"] {
            background-color: #f8f9fa !important;
            border-right: 1px solid #e0e0e0;
        }

        /* Red Accents for Checkboxes and Slider */
        input[type="checkbox"]:checked {
            background-color: #e63946 !important;
        }
        
        .stSlider [data-baseweb="slider"] [role="slider"] {
            background-color: #e63946;
        }
        
        /* Cleanr Logo Style */
        .logo-text {
            font-size: 80px;
            font-weight: 700;
            color: #112340;
            text-align: center;
            margin-top: 20px;
            margin-bottom: 40px;
        }

        /* Headers */
        h1, h2, h3 {
            color: #112340;
            font-weight: 700;
        }

        /* Feedback Box */
        .feedback-box {
            background-color: #ffffff;
            padding: 20px;
            border-radius: 12px;
            border: 1px solid #e0e0e0;
            margin-top: 20px;
        }
    </style>
""", unsafe_allow_html=True)

# --- Constants & Helpers ---
TITLES = {'dr', 'dr.', 'prof', 'prof.', 'sir', 'mr', 'mrs', 'ms', 'mx', 'hon'}
COMMON_SUFFIXES = ['ltd', 'inc', 'group', 'brands', 'company', 'companies', 'incorporation', 'corporation']

def clean_first_name(name):
    if pd.isna(name) or not str(name).strip(): return ''
    name = fix_text(unidecode(str(name))).strip()
    name = re.sub(r'\(.*?\)|\[.*?\]', '', name)
    parts = name.split()
    if not parts: return ''
    # Smart Title Handling (e.g., "Dr. Rian" -> "Rian")
    if parts[0].lower() in TITLES and len(parts) > 1:
        first = parts[1]
    else:
        first = parts[0]
    return first.title()

def clean_last_name(name):
    if pd.isna(name) or not str(name).strip(): return ''
    name = fix_text(unidecode(str(name))).strip()
    parts = name.split()
    processed = []
    for p in parts:
        if p.lower().startswith("mc") and len(p) > 2:
            processed.append("Mc" + p[2:].title())
        else:
            processed.append(p.title())
    return " ".join(processed)

def clean_company(name):
    if pd.isna(name) or not str(name).strip(): return ''
    name = fix_text(unidecode(str(name)))
    name = re.split(r'\s[-|:|–]\s', name)[0]
    name = re.sub(r'\b(?:' + '|'.join(COMMON_SUFFIXES) + r')\b', '', name, flags=re.IGNORECASE)
    name = re.sub(r'[^A-Za-z0-9\s\-]', '', name)
    name = re.sub(r'\s{2,}', ' ', name).strip()
    return name.title()

def detect_email_pattern(row):
    email = str(row.get('Email', '')).lower()
    first = str(row.get('First Name', '')).lower()
    last = str(row.get('Last Name', '')).lower()
    if not email or '@' not in email: return 'unknown'
    local = email.split('@')[0]
    if f"{first}.{last}" == local: return "first.last"
    if f"{first[0]}{last}" == local: return "finitiallast"
    return "other"

def infer_last_from_email(first, email):
    if not email or pd.isna(email) or not first: return ''
    user = email.split('@')[0].lower()
    first_l = first.lower()
    m = re.match(rf"{re.escape(first_l)}[._-]?([a-z]+)", user)
    if m: return m.group(1).title()
    if user.startswith(first_l[0]):
        guess = user[len(first_l[0]):]
        if guess: return guess.title()
    return ''

# --- Main App ---
st.markdown('<div class="logo-text">Cleanr.</div>', unsafe_allow_html=True)

with st.sidebar:
    st.header("⚙️ Cleaning Options")
    # Updated Defaults: Email Patterns ON, others OFF
    opt_names = st.checkbox("Clean Names", value=True)
    opt_company = st.checkbox("Clean Company Names", value=True)
    opt_infer = st.checkbox("Infer Last Names from Email", value=True)
    opt_patterns = st.checkbox("Check Company Email Patterns", value=True)
    opt_phone = st.checkbox("Clean Mobile Numbers", value=False)
    opt_titles = st.checkbox("Clean Job Titles", value=False)
    opt_split = st.checkbox("Split by Company", value=False)
    
    if opt_split:
        max_lists = st.slider("Max Lists", 1, 10, 4)
    
    st.divider()
    uploaded_file = st.file_uploader("Upload CSV File", type=["csv"])

if uploaded_file:
    df = pd.read_csv(uploaded_file, encoding='latin1')
    # Standardize column headers
    df.columns = [col.strip().title().replace('_', ' ') for col in df.columns]
    
    # Auto-map Company Name
    if 'Company Name' in df.columns and 'Company' not in df.columns:
        df.rename(columns={'Company Name': 'Company'}, inplace=True)
    
    cleaned_df = df.copy()
    
    # Process Data
    with st.spinner('Cleaning...'):
        for i, row in df.iterrows():
            orig_first = str(row.get('First Name', '')).strip()
            orig_last = str(row.get('Last Name', '')).strip()
            orig_co = str(row.get('Company', '')).strip()
            email = str(row.get('Email', '')).strip()

            if opt_names:
                first = clean_first_name(orig_first)
                cleaned_df.at[i, 'First Name'] = first
                
                if orig_last and orig_last.lower() != 'nan' and orig_last != '':
                    cleaned_df.at[i, 'Last Name'] = clean_last_name(orig_last)
                elif opt_infer:
                    inferred = infer_last_from_email(first, email)
                    cleaned_df.at[i, 'Last Name'] = clean_last_name(inferred)
            
            if opt_company:
                cleaned_df.at[i, 'Company'] = clean_company(orig_co)
        
        if opt_patterns:
            cleaned_df['Email Pattern'] = cleaned_df.apply(detect_email_pattern, axis=1)

    st.success("Data cleaned successfully!")
    
    # Excel Download
    wb = Workbook()
    ws = wb.active
    ws.title = "Cleaned Leads"
    for r in dataframe_to_rows(cleaned_df, index=False, header=True):
        ws.append(r)
    
    output = BytesIO()
    wb.save(output)
    
    st.download_button(
        label="📥 Download Cleaned Excel",
        data=output.getvalue(),
        file_name="cleaned_outreach.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    st.dataframe(cleaned_df.head(20))

# Feedback Section
st.markdown('<div class="feedback-box">', unsafe_allow_html=True)
st.markdown("### 💬 Leave Feedback")
feedback_text = st.text_area("Have any suggestions, bugs, or feature ideas?", placeholder="Enter your feedback here...")
if st.button("Submit Feedback"):
    st.toast("Feedback received! Thanks for helping improve Cleanr.")
st.markdown('</div>', unsafe_allow_html=True)
