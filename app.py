import streamlit as st
import pandas as pd
import re
from unidecode import unidecode

# --- Company Cleaning ---
COMMON_SUFFIXES = [
    'ltd', 'inc', 'group', 'brands', 'company', 'companies', 'incorporation'
]

def clean_company(name):
    if pd.isna(name): return ''
    try:
        name = name.encode('latin1').decode('utf-8')
    except (UnicodeEncodeError, UnicodeDecodeError):
        pass
    name = unidecode(str(name))
    name = re.sub(r'\b(?:' + '|'.join(COMMON_SUFFIXES) + r')\b', '', name, flags=re.IGNORECASE)
    name = re.sub(r'[^A-Za-z0-9\s\-]', '', name)
    name = re.sub(r'\s{2,}', ' ', name)
    return name.strip().title()

# --- Name Cleaning ---
def clean_name(name, is_first=True):
    if pd.isna(name): return ''
    try:
        name = name.encode('latin1').decode('utf-8')
    except (UnicodeEncodeError, UnicodeDecodeError):
        pass
    name = unidecode(str(name)).strip()
    name_parts = name.split()
    if not name_parts:
        return ''
    return name_parts[0].title() if is_first else name_parts[-1].title()

# --- Main Cleaning Function ---
def clean_data(df):
    cleaned_df = df.copy()
    changes = 0

    for i, row in df.iterrows():
        original_first = str(row.get('First Name', '')).strip()
        original_last = str(row.get('Last Name', '')).strip()
        original_company = str(row.get('Company', '')).strip()

        first = clean_name(original_first, is_first=True)
        last = clean_name(original_last, is_first=False)
        company = clean_company(original_company)

        if (first != original_first) or (last != original_last) or (company != original_company):
            changes += 1

        cleaned_df.at[i, 'First Name'] = first
        cleaned_df.at[i, 'Last Name'] = last
        cleaned_df.at[i, 'Company'] = company

    percent_cleaned = (changes / len(df)) * 100 if len(df) else 0
    return cleaned_df, percent_cleaned

# --- Streamlit UI ---
st.set_page_config(page_title="Excel Cleaner", layout="centered")

st.title("🧼 Excel Cleaner for Mail Merge")
st.markdown("Upload your Cognism or LinkedIn CSV export and get a cleaned version ready for mail merge.")

uploaded_file = st.file_uploader("📤 Upload CSV File", type=["csv"])

if uploaded_file:
    df = pd.read_csv(uploaded_file, encoding='latin1')
    df.columns = [col.strip().title().replace('_', ' ') for col in df.columns]
    column_map = {'Company Name': 'Company'}
    df.rename(columns=column_map, inplace=True)

    st.write("📋 Detected columns:", df.columns.tolist())

    cleaned_df, percent_cleaned = clean_data(df)

    st.success("✅ Done! Your data is cleaned and ready to download.")
    st.info(f"🧮 {percent_cleaned:.1f}% of rows were cleaned or updated.")

    st.download_button(
        label="📥 Download Cleaned CSV",
        data=cleaned_df.to_csv(index=False).encode("utf-8-sig"),
        file_name="cleaned_output.csv",
        mime="text/csv"
    )

    st.subheader("🔍 Before Cleaning")
    st.dataframe(df.head(10))

    st.subheader("✨ After Cleaning")
    st.dataframe(cleaned_df.head(10))
