import streamlit as st
import pandas as pd
import re
from unidecode import unidecode
from ftfy import fix_text
import requests

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

# --- Cleaning Logic ---
COMMON_SUFFIXES = ['ltd', 'inc', 'group', 'brands', 'company', 'companies', 'incorporation']

def clean_company(name):
    if pd.isna(name): return ''
    try:
        name = name.encode('latin1').decode('utf-8')
    except: pass
    name = fix_text(name)
    name = unidecode(str(name))
    name = re.sub(r'\b(?:' + '|'.join(COMMON_SUFFIXES) + r')\b', '', name, flags=re.IGNORECASE)
    name = re.sub(r'[^A-Za-z0-9\s\-]', '', name)
    name = re.sub(r'\s{2,}', ' ', name)
    name = name.strip()
    return name.upper() if len(name) <= 3 else name.title()

def clean_name(name, is_first=True):
    if pd.isna(name): return ''
    try:
        name = name.encode('latin1').decode('utf-8')
    except: pass
    name = fix_text(name)
    name = unidecode(str(name)).strip()
    name_parts = name.split()
    return name_parts[0].title() if is_first and name_parts else name_parts[-1].title() if name_parts else ''

def infer_from_email(first, last, email):
    if pd.isna(email): return first, last
    user = email.split('@')[0].lower()
    if len(last) == 1:
        pattern = re.escape(first.lower()) + r'[._]?([a-z]+)'
        match = re.match(pattern, user)
        if match: return first, match.group(1).title()
        if user.startswith(first[0].lower()):
            guess = user[len(first[0]):]
            return first, guess.title() if guess else last
    return first, last

def clean_data(df):
    cleaned_df = df.copy()
    changes = 0
    for i, row in df.iterrows():
        orig_first, orig_last = str(row.get('First Name', '')).strip(), str(row.get('Last Name', '')).strip()
        orig_company = str(row.get('Company', '')).strip()
        email = str(row.get('Email', '')).strip() if 'Email' in df.columns else ''
        first, last = clean_name(orig_first, True), clean_name(orig_last, False)
        company = clean_company(orig_company)
        first, last = infer_from_email(first, last, email)
        if first != orig_first or last != orig_last or company != orig_company: changes += 1
        cleaned_df.at[i, 'First Name'] = first
        cleaned_df.at[i, 'Last Name'] = last
        cleaned_df.at[i, 'Company'] = company
    pct = (changes / len(df)) * 100 if len(df) else 0
    return cleaned_df, pct

# --- UI Layout ---
st.set_page_config(page_title="Cleanr", layout="centered")

st.markdown('<div class="title-text">Cleanr.</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle-text">Clean your data faster.</div>', unsafe_allow_html=True)

st.markdown('<div class="rounded-box">Upload your Cognism CSV export and get a cleaned version ready for mail merge.</div>', unsafe_allow_html=True)

uploaded_file = st.file_uploader("Upload CSV File", type=["csv"])

if uploaded_file:
    df = pd.read_csv(uploaded_file, encoding='latin1')
    df.columns = [col.strip().title().replace('_', ' ') for col in df.columns]
    df.rename(columns={'Company Name': 'Company'}, inplace=True)

    cleaned_df, percent_cleaned = clean_data(df)

    st.success("✅ Done! Your data is cleaned and ready to download.")
    st.info(f"📊 {percent_cleaned:.1f}% of rows were cleaned or updated.")

    # Log to Google Sheets
    tracking_data = {
        'filename': uploaded_file.name,
        'rows': len(df),
        'cleaned': int((percent_cleaned / 100) * len(df)),
        'percent_cleaned': round(percent_cleaned, 1),
        'time_saved': round((int((percent_cleaned / 100) * len(df)) * 7.5) / 60, 1)  # 7.5s per row manually - 1s tool time
    }
    try:
        requests.post(
            "https://script.google.com/macros/s/AKfycbxS1qSh_ge3DQCbWNhjsvtWa4dvSrx9rBXs9PyJ0sC8P8tYyRGzoNJRAK7tfZvsx_sr1A/exec",
            json={**tracking_data, "type": "usage"}
        )
    except:
        pass

    st.download_button(
        label="📥 Download Cleaned CSV",
        data=cleaned_df.to_csv(index=False).encode("utf-8-sig"),
        file_name=uploaded_file.name.replace('.csv', '_cleaned.csv'),
        mime="text/csv",
        key="download-cleaned",
    )

    st.markdown("<div class='section-header'>Preview</div>", unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("<b>Before Cleaning</b>", unsafe_allow_html=True)
        st.dataframe(df.head(10))
    with col2:
        st.markdown("<b>After Cleaning</b>", unsafe_allow_html=True)
        st.dataframe(cleaned_df.head(10))

# --- Feedback Form ---
st.markdown("<div class='feedback-container'>", unsafe_allow_html=True)
st.markdown("<h4>💬 Leave Feedback</h4>", unsafe_allow_html=True)
with st.form(key="feedback_form"):
    feedback = st.text_area("Have any suggestions, bugs, or feature ideas?", height=120, key="feedback_box")
    submitted = st.form_submit_button("Submit")
    if submitted:
        if feedback.strip():
            try:
                response = requests.post(
                    "https://script.google.com/macros/s/AKfycbxS1qSh_ge3DQCbWNhjsvtWa4dvSrx9rBXs9PyJ0sC8P8tYyRGzoNJRAK7tfZvsx_sr1A/exec",
                    json={"type": "feedback", "timestamp": str(pd.Timestamp.now()), "message": feedback.strip()}
                )
                if response.status_code == 200:
                    st.success("✅ Thanks! Your feedback was submitted.")
                else:
                    st.error("⚠️ Something went wrong. Please try again later.")
            except Exception as e:
                st.error(f"❌ Failed to submit feedback: {e}")
        else:
            st.warning("✏️ Please write something before submitting.")
st.markdown("</div>", unsafe_allow_html=True)
