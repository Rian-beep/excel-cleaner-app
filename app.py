import streamlit as st
import pandas as pd
import re
import emoji
from unidecode import unidecode

# --- Company Cleaning ---
COMMON_SUFFIXES = ['ltd', 'inc', 'group', 'company', 'brands']
NAME_MAP = {
    'colgate-palmolive': 'colgate',
    'loreal': 'loreal',
}

def fix_mojibake(text):
    if pd.isna(text): return ''
    try:
        return text.encode('latin1').decode('utf-8')
    except (UnicodeEncodeError, UnicodeDecodeError):
        return text

def clean_company(name):
    if pd.isna(name): return ''
    name = fix_mojibake(str(name))
    name = unidecode(name)
    name = re.sub(r'\b(?:' + '|'.join(COMMON_SUFFIXES) + r')\b', '', name, flags=re.IGNORECASE)
    name = re.sub(r'[^A-Za-z0-9\s\-]', '', name)
    name = name.strip().lower()
    for k, v in NAME_MAP.items():
        if k in name:
            return v.capitalize()
    return name.capitalize()

# --- Name Cleaning ---
def clean_name(name):
    if pd.isna(name): return ''
    name = fix_mojibake(str(name))
    name = unidecode(name).strip().title()
    return name

def infer_missing_name(first, last, email):
    if pd.isna(email):
        return first, last

    email_user = email.split('@')[0].lower()
    parts = re.split(r'[._\-]', email_user)

    # Common patterns like john.smith, smith.john
    if len(parts) >= 2:
        if len(first) <= 2:
            first = parts[0].capitalize()
        if len(last) <= 2:
            last = parts[-1].capitalize()
        return first, last

    # jsmith → First: J, Last: Smith
    if len(parts) == 1 and len(first) <= 2 and len(last) <= 2:
        match = re.match(r'^([a-zA-Z])([a-zA-Z]+)$', email_user)
        if match:
            first_guess, last_guess = match.groups()
            if len(first) <= 2:
                first = first_guess.capitalize()
            if len(last) <= 2:
                last = last_guess.capitalize()
            return first, last

    # NEW: johnsmith@... or smithjohn@...
    if len(first) > 1 and len(last) <= 2:
        if first.lower() in email_user:
            guess = email_user.replace(first.lower(), '', 1)
            if guess and (last.lower() == guess[0] or len(last) <= 2):
                last = guess.capitalize()
                return first, last
        else:
            # Try matching first initial + last
            if email_user.startswith(first[0].lower()) and len(email_user) > 2:
                guess = email_user[1:]
                if guess and (last.lower() == guess[0] or len(last) <= 2):
                    last = guess.capitalize()
                    return first, last

    # Also check reverse (smithjohn)
    if len(last) > 1 and len(first) <= 2:
        if last.lower() in email_user:
            guess = email_user.replace(last.lower(), '', 1)
            if guess and (first.lower() == guess[0] or len(first) <= 2):
                first = guess.capitalize()
                return first, last
        else:
            if email_user.endswith(last.lower()):
                guess = email_user.replace(last.lower(), '', 1)
                if guess and guess[0] == last[0].lower():
                    first = guess.capitalize()
                    return first, last

    return first, last


# --- Special Characters ---
def remove_special_chars(text):
    if pd.isna(text): return ''
    text = fix_mojibake(str(text))
    text = unidecode(text)
    text = emoji.replace_emoji(text, replace='')
    text = re.sub(r'[^\w\s\-@\.]', '', text)
    return text.strip()

# --- Main Cleaning Function ---
def clean_data(df):
    if 'First Name' not in df.columns or 'Last Name' not in df.columns or 'Company' not in df.columns:
        st.error("CSV must contain 'First Name', 'Last Name', and 'Company' columns.")
        return df

    for i, row in df.iterrows():
        first = clean_name(remove_special_chars(row.get('First Name', '')))
        last = clean_name(remove_special_chars(row.get('Last Name', '')))
        email = remove_special_chars(row.get('Email', ''))
        first, last = infer_missing_name(first, last, email)

        df.at[i, 'First Name'] = first
        df.at[i, 'Last Name'] = last
        df.at[i, 'Company'] = clean_company(row.get('Company', ''))
        if 'Email' in df.columns:
            df.at[i, 'Email'] = email

    return df

# --- Streamlit UI ---
st.set_page_config(page_title="Excel Cleaner", layout="centered")

st.title("🧼 Excel Cleaner for Mail Merge")
st.markdown("Upload your Cognism or LinkedIn CSV export and get a cleaned version ready for mail merge.")

uploaded_file = st.file_uploader("📤 Upload CSV File", type=["csv"])

if uploaded_file:
    df = pd.read_csv(uploaded_file)

    # Normalize column names
    df.columns = [col.strip().title().replace('_', ' ') for col in df.columns]
    st.write("📋 Detected columns:", df.columns.tolist())

    cleaned_df = clean_data(df.copy())

    st.success("✅ Done! Your data is cleaned and ready to download.")

    st.download_button(
        "📥 Download Cleaned CSV",
        cleaned_df.to_csv(index=False).encode("utf-8"),
        "cleaned_output.csv",
        "text/csv"
    )

    st.subheader("🔍 Before Cleaning")
    st.dataframe(df.head(10))

    st.subheader("✨ After Cleaning")
    st.dataframe(cleaned_df.head(10))
