# Cleanr - Data Cleaning Tool for Email Outreach

A powerful Streamlit-based data cleaning tool designed to prepare contact data for email outreach campaigns. Clean and standardize contact information, validate emails, detect duplicates, and organize your data for better deliverability.

## Features

### Core Cleaning Features
- **Name Cleaning**: Standardizes first and last names with proper capitalization
- **Company Name Cleaning**: Uses a company directory for known mappings and standardizes company names
- **Email Validation**: Validates email format and flags disposable email domains
- **Phone Number Cleaning**: Standardizes phone numbers to E.164 format (when phonenumbers library is available)
- **Job Title Standardization**: Cleans and expands common job title abbreviations
- **Last Name Inference**: Automatically infers missing last names from email addresses

### Data Quality Features
- **Quality Scoring**: Calculates a 0-100 quality score for each contact based on completeness and validity
- **Duplicate Detection**: Identifies duplicate contacts based on email or name combinations
- **Enhanced Statistics**: Provides detailed metrics on data quality and cleaning results

### Outreach Optimization
- **Company-Based Splitting**: Automatically splits contacts from the same company across multiple sending lists to improve deliverability
- **Configurable List Count**: Choose how many lists to split contacts into (1-10)

### Export Options
- **Excel Export**: Download cleaned data with highlighted changes and quality scores
- **CSV Export**: Standard CSV format for easy import
- **JSON Export**: JSON format for API integrations

### User Experience
- **Auto Column Detection**: Automatically detects email, phone, and job title columns
- **Configurable Options**: Sidebar with toggles for all cleaning features
- **Visual Feedback**: Color-coded highlights in Excel exports showing what was changed
- **Detailed Reports**: View invalid emails, duplicates, and quality statistics

## Installation

1. Install dependencies:
```bash
pip install -r requirements.txt
```

2. Run the application:
```bash
streamlit run app.py
```

## Requirements

- Python 3.7+
- Streamlit
- Pandas
- openpyxl (for Excel export)
- unidecode (for text normalization)
- ftfy (for encoding fixes)
- email-validator (optional, for enhanced email validation)
- phonenumbers (optional, for phone number parsing)

## Usage

1. Upload a CSV file with contact data
2. Configure cleaning options in the sidebar
3. Review statistics and quality metrics
4. Download cleaned data in your preferred format (Excel, CSV, or JSON)

## Column Detection

The tool automatically detects common column names:
- **Email**: email, e-mail, mail, email address
- **Phone**: phone, telephone, tel, mobile, cell
- **Job Title**: title, job title, position, role

## Data Quality Scoring

Each contact receives a quality score (0-100) based on:
- Email validity (30 points)
- First name presence (20 points)
- Last name presence (20 points)
- Company name presence (15 points)
- Phone number validity (15 points)

## Company Directory

The tool uses `company_directory.csv` to map known company name variations to standardized names. Add your own mappings to improve cleaning accuracy.

## Tips for Better Deliverability

- Enable company-based splitting to avoid sending multiple emails to the same company in quick succession
- Review and remove invalid emails before sending
- Check quality scores to identify contacts that need manual review
- Remove duplicates to avoid sending to the same person twice

## Feedback

Found a bug or have a feature request? Use the feedback form in the app to let us know!
