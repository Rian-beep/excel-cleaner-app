import streamlit as st
import pandas as pd
import re
from unidecode import unidecode
from ftfy import fix_text
import requests
import os
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
import random

# --- Page Config ---
st.set_page_config(page_title="Cleanr", layout="centered")

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

# --- Load Known Company Names ---
company_dict = {}
company_file = "company_directory.csv"
if os.path.exists(company_file):
    known_companies_df = pd.read_csv(company_file)
    company_dict = dict(
        zip(
            known_companies_df["Raw Company"].str.strip().str.lower(),
            known_companies_df["Cleaned Company"].str.strip()
        )
    )

# --- Cleaning Rules ---
COMMON_SUFFIXES = ['ltd', 'inc', 'group', 'brands', 'company', 'companies', 'incorporation', 'corporation']


def clean_company(name):
    """Clean company name, using known mappings if available."""
    if pd.isna(name):
        return ''
    raw_name = name.strip()
    name_key = raw_name.lower()

    if name_key in company_dict:
        return company_dict[name_key]

    try:
        name = name.encode('latin1').decode('utf-8')
    except Exception:
        pass
    name = fix_text(name)
    name = unidecode(str(name))

    name = re.sub(r'\b(?:' + '|'.join(COMMON_SUFFIXES) + r')\b', '', name, flags=re.IGNORECASE)
    name = re.sub(r'[^A-Za-z0-9\s\-]', '', name)
    name = re.sub(r'\s{2,}', ' ', name).strip()

    if len(name) <= 4:
        name = name.upper()
    else:
        name = name.title()

    return name


def clean_first_name(name):
    """Clean first name and take the first token."""
    if pd.isna(name):
        return ''
    try:
        name = name.encode('latin1').decode('utf-8')
    except Exception:
        pass
    name = fix_text(name)
    name = unidecode(str(name)).strip()
    if not name:
        return ''
    parts = name.split()
    first = parts[0]
    return first.title()


def clean_last_name(name):
    """
    Clean last name without changing its structure.
    Keep full string, fix encoding, spacing and casing.
    """
    if pd.isna(name):
        return ''
    try:
        name = name.encode('latin1').decode('utf-8')
    except Exception:
        pass
    name = fix_text(name)
    name = unidecode(str(name)).strip()
    if not name:
        return ''

    parts = name.split()
    processed_parts = []
    for part in parts:
        if part.lower().startswith("mc") and len(part) > 2:
            part = "Mc" + part[2:].title()
        else:
            part = part.title()
        processed_parts.append(part)

    return " ".join(processed_parts)


def infer_last_from_email(first_name, email):
    """
    Infer last name from email ONLY if last name is missing.
    Handles:
      - firstname.lastname@...
      - firstname_lastname@...
      - firstname-lastname@...
      - fLastname@...
    """
    if not email or pd.isna(email) or not first_name:
        return ''

    user = email.split('@')[0].lower()
    first_lower = first_name.lower()

    # firstname[._-]lastname
    m = re.match(rf"{re.escape(first_lower)}[._-]?([a-z]+)", user)
    if m:
        return m.group(1).title()

    # fLastname pattern (e.g. jsmith)
    if user.startswith(first_lower[0]):
        guess = user[len(first_lower[0]):]
        if guess:
            return guess.title()

    return ''


def clean_data(df):
    cleaned_df = df.copy()
    changes = 0
    changed_mask = pd.DataFrame(False, index=df.index, columns=df.columns)

    for i, row in df.iterrows():
        orig_first = str(row.get('First Name', '')).strip()
        orig_last = str(row.get('Last Name', '')).strip()
        orig_company = str(row.get('Company', '')).strip()
        email = str(row.get('Email', '')).strip() if 'Email' in df.columns else ''

        # First name
        first = clean_first_name(orig_first)

        # Last name: clean if present, infer from email if missing
        if orig_last:
            last = clean_last_name(orig_last)
        else:
            inferred_last = infer_last_from_email(first, email)
            last = clean_last_name(inferred_last) if inferred_last else ''

        # Company
        company = clean_company(orig_company)

        if first != orig_first:
            changed_mask.at[i, 'First Name'] = True
            cleaned_df.at[i, 'First Name'] = first
            changes += 1

        if last != orig_last:
            changed_mask.at[i, 'Last Name'] = True
            cleaned_df.at[i, 'Last Name'] = last
            changes += 1

        if company != orig_company:
            changed_mask.at[i, 'Company'] = True
            cleaned_df.at[i, 'Company'] = company
            changes += 1

    pct = (changes / len(df)) * 100 if len(df) else 0
    return cleaned_df, pct, changed_mask


def split_into_lists_by_company(df, max_lists=4):
    """
    Split contacts into up to max_lists lists so that people
    from the same company are spread across different lists.
    Returns a list of lists of row indices.
    """
    if 'Company' not in df.columns or df.empty:
        return [list(df.index)]

    company_groups = {}
    for idx, company in df['Company'].items():
        company_groups.setdefault(company, []).append(idx)

    max_count = max(len(indices) for indices in company_groups.values())
    num_lists = min(max_lists, max_count if max_count > 0 else 1)

    batches = [[] for _ in range(num_lists)]

    for company, indices in company_groups.items():
        random.shuffle(indices)  # randomise order within company
        for i, idx in enumerate(indices):
            batches[i % num_lists].append(idx)

    return [batch for batch in batches if batch]


def generate_highlighted_excel_with_splits(df, mask, split_batches=None, max_lists=4):
    """
    Create an Excel file with:
      - Sheet "All Contacts": full cleaned data with highlights.
      - If split_batches is provided and not empty:
          up to max_lists sheets "List 1" ... "List N": split by company.
    """
    wb = Workbook()
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    # ---- Sheet 1: All Contacts ----
    ws_all = wb.active
    ws_all.title = "All Contacts"

    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
        for c_idx, value in enumerate(row, 1):
            cell = ws_all.cell(row=r_idx, column=c_idx, value=value)
            if r_idx > 1:
                col = df.columns[c_idx - 1]
                orig_idx = df.index[r_idx - 2]
                if (col in mask.columns) and mask.at[orig_idx, col]:
                    cell.fill = yellow_fill

    # ---- Additional sheets: List 1..N (only if splitting enabled) ----
    if split_batches:
        for list_num, indices in enumerate(split_batches[:max_lists], start=1):
            ws = wb.create_sheet(title=f"List {list_num}")
            sub_df = df.loc[indices]
            sub_mask = mask.loc[indices]

            for r_idx, row in enumerate(dataframe_to_rows(sub_df, index=False, header=True), 1):
                for c_idx, value in enumerate(row, 1):
                    cell = ws.cell(row=r_idx, column=c_idx, value=value)
                    if r_idx > 1:
                        col = sub_df.columns[c_idx - 1]
                        orig_idx = sub_df.index[r_idx - 2]
                        if (col in sub_mask.columns) and sub_mask.at[orig_idx, col]:
                            cell.fill = yellow_fill

    output = BytesIO()
    wb.save(output)
    return output.getvalue()


# --- UI Layout ---
st.markdown('<div class="title-text">Cleanr.</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle-text">Clean your data faster.</div>', unsafe_allow_html=True)
st.markdown('<div class="rounded-box">Upload your Cognism CSV export and get a cleaned version ready for mail merge.</div>', unsafe_allow_html=True)

# Toggle for list splitting
split_enabled = st.checkbox(
    "Split contacts from the same company into separate sending lists",
    value=True,
    help="When enabled, contacts from the same organisation are spread across up to four lists so you are not emailing multiple people at the same company seconds apart. This can improve deliverability and reduce the chance of your campaign hitting spam filters."
)

st.caption(
    "Tip: Sending many emails to the same company in a short burst can look suspicious to spam filters. "
    "Splitting contacts across several lists helps smooth out sends and protect deliverability."
)

uploaded_file = st.file_uploader("Upload CSV File", type=["csv"])

if uploaded_file:
    df = pd.read_csv(uploaded_file, encoding='latin1')
    df.columns = [col.strip().title().replace('_', ' ') for col in df.columns]
    df.rename(columns={'Company Name': 'Company'}, inplace=True)

    cleaned_df, percent_cleaned, changed_mask = clean_data(df)

    st.success("✅ Done! Your data is cleaned and ready to download.")
    st.info(f"📊 {percent_cleaned:.1f}% of rows were cleaned or updated.")

    split_batches = None

    if split_enabled:
        split_batches = split_into_lists_by_company(cleaned_df, max_lists=4)
        if len(split_batches) > 1:
            st.info(
                "📧 Splitting is enabled. Your cleaned data has been split into "
                f"{len(split_batches)} sending lists (maximum 4) to help protect deliverability."
            )
            for i, batch in enumerate(split_batches, start=1):
                st.write(f"List {i}: {len(batch)} contacts")
        else:
            st.write("📧 Splitting is enabled, but there are no companies with multiple contacts. All contacts are effectively in a single list.")
    else:
        st.info("📧 Splitting is turned off. All contacts will be kept in a single list.")

    # Send Usage Log
    cleaned_rows = int((percent_cleaned / 100) * len(df))
    usage_data = {
        "type": "usage",
        "sheet": "Usage",
        "filename": uploaded_file.name,
        "rows": len(df),
        "cleaned": cleaned_rows,
        "percent_cleaned": round(percent_cleaned, 1),
        "time_saved": round((cleaned_rows * 7.5) / 60, 1)
    }
    try:
        requests.post(
            "https://script.google.com/macros/s/AKfycbxM7dmZfMIuWcNWiyxAh8nwX69rvuRaioJ6EH_k7Vx9DRu6DdYdMIO3ZbsZmH--Q5q1/exec",
            json=usage_data,
            timeout=2
        )
    except Exception:
        pass

    # Excel with All Contacts + optional split lists
    excel_file = generate_highlighted_excel_with_splits(
        cleaned_df,
        changed_mask,
        split_batches=split_batches if split_enabled else None,
        max_lists=4
    )

    st.download_button(
        label="📥 Download Cleaned File",
        data=excel_file,
        file_name=uploaded_file.name.replace('.csv', '_cleaned.xlsx'),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
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
                requests.post(
                    "https://script.google.com/macros/s/AKfycbxM7dmZfMIuWcNWiyxAh8nwX69rvuRaioJ6EH_k7Vx9DRu6DdYdMIO3ZbsZmH--Q5q1/exec",
                    json={
                        "type": "feedback",
                        "sheet": "Feedback",
                        "timestamp": str(pd.Timestamp.now()),
                        "message": feedback.strip()
                    },
                    timeout=2
                )
                st.success("✅ Thanks! Your feedback was submitted.")
            except Exception as e:
                st.error(f"❌ Failed to submit feedback: {e}")
        else:
            st.warning("✏️ Please write something before submitting.")
st.markdown("</div>", unsafe_allow_html=True)
