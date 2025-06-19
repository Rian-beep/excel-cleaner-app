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
        text = text.encode('latin1').decode('utf-8')
    except (UnicodeEncodeError, UnicodeDecodeError):
        pass
    return unidecode(text)

def clean_company(name):
    if pd.isna(name): return ''
    name = str(name)
    name = unidecode(name)  # <- always remove accents
    name = re.sub(r'\b(?:' + '|'.join(COMMON_SUFFIXES) + r')\b', '', name, flags=re.IGNORECASE)
    name = re.sub(r'[^A-Za-z0-9\s\-]', '', name)
    name = name.strip().lower()

    for k, v in NAME_MAP.items():
        if k in name:
            return v.capitalize()

    return name.capitalize()

# --- Name Cleaning ---
def clean_name(name, is_first=True):
    if pd.isna(name): return ''
    name = str(name).strip()
    name = unidecode(name)  # <- always remove accents
    name_parts = name.split()

    if not name_parts:
        return ''

    return name_parts[0].title() if is_first else name_parts[-1].title()

def infer_missing_name(first, last, email):
    if pd.isna(email):
        return first, last

    email_user = email.split('@')[0].lower()
    parts = re.split(r'[._\-]', email_user)

    if len(parts) >= 2:
        if len(first) <= 2:
            first = parts[0].capitalize()
        if len(last) <= 2:
            last = parts[-1].capitalize()
        return first, last

    if len(parts) == 1 and len(first) <= 2 and len(last) <= 2:
        match = re.match(r'^([a-zA-Z])([a-zA-Z]+)$', email_user)
        if match:
            first_guess, last_guess = match.groups()
            if len(first) <= 2:
                first = first_guess.capitalize()
            if len(last) <= 2:
                last = last_guess.capitalize()
            return first, last

    if len(first) > 1 and len(last) <= 2:
        if first.lower() in email_user:
            guess = email_user.replace(first.lower(), '', 1)
            if guess and (last.lower() == guess[0] or len(last) <= 2):
                last = guess.capitalize()
                return first, last
        else:
            if email_user.startswith(first[0].lower()) and len(email_user) > 2:
                guess = email_user[1:]
                if guess and (last.lower() == guess[0] or len(last) <= 2):
                    last = guess.capitalize()
                    return first, last

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
    text = emoji.replace_emoji(text, replace='')
    text = re.sub(r'[^\w\s\-@\.]', '', text)
    return text.strip()

# --- Main Cleaning Function ---
def clean_data(df):
    if 'First Name' not in df.columns or 'Last Name' not in df.columns or 'Company' not in df.columns:
        st.error("CSV must contain 'First Name', 'Last Name', and 'Company' columns.")
        return df, 0.0

    cleaned_df = df.copy()
    rows_changed = 0

    for i, row in df.iterrows():
        original = {
            'First Name': str(row.get('First Name', '')).strip(),
            'Last Name': str(row.get('Last Name', '')).strip(),
            'Company': str(row.get('Company', '')).strip(),
            'Email': str(row.get('Email', '')).strip() if 'Email' in df.columns else ''
        }

        first = clean_name(remove_special_chars(original['First Name']), is_first=True)
        last = clean_name(remove_special_chars(original['Last Name']), is_first=False)
        email = remove_special_chars(original['Email'])
        first, last = infer_missing_name(first, last, email)
        company = clean_company(original['Company'])

        if first != original['First Name'] or last != original['Last Name'] or company != original['Company']:
            rows_changed += 1

        cleaned_df.at[i, 'First Name'] = first
        cleaned_df.at[i, 'Last Name'] = last
        cleaned_df.at[i, 'Company'] = company
        if 'Email' in df.columns:
            cleaned_df.at[i, 'Email'] = email

    percent_cleaned = (rows_changed / len(df)) * 100 if len(df) else 0
    return cleaned_df, percent_cleaned


# --- Streamlit UI ---
st.set_page_config(page_title="Excel Cleaner", layout="centered")

st.title("🧼 Excel Cleaner for Mail Merge")
st.markdown("Upload your Cognism or LinkedIn CSV export and get a cleaned version ready for mail merge.")

uploaded_file = st.file_uploader("📤 Upload CSV File", type=["csv"])

if uploaded_file:
    df = pd.read_csv(uploaded_file)
    df.columns = [col.strip().title().replace('_', ' ') for col in df.columns]

    st.write("📋 Detected columns:", df.columns.tolist())

    cleaned_df, percent_cleaned = clean_data(df.copy())

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
