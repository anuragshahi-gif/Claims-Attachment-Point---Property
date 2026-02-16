import streamlit as st
import pandas as pd
import pdfplumber
import docx
import re
import io
from dateutil import parser

# --- CONFIGURATION ---
st.set_page_config(page_title="Attachment Point Assessor", layout="wide")

# --- SESSION STATE & RESET LOGIC ---
# We use a dynamic key for the file uploader. Changing this key forces the uploader to reset.
if "uploader_key" not in st.session_state:
    st.session_state["uploader_key"] = 0

if "text_input" not in st.session_state:
    st.session_state["text_input"] = ""

def reset_all():
    """Resets all inputs and clears the file uploader."""
    st.session_state["uploader_key"] += 1  # Incrementing key resets the file uploader widget
    st.session_state["text_input"] = ""    # Clears the text area
    # We can also clear other session variables if needed here

# --- UTILITY FUNCTIONS ---

def clean_currency(value):
    """Converts string currency to float. Returns 0.0 if failed."""
    if isinstance(value, (int, float)):
        return float(value)
    
    clean_str = str(value).upper().replace('$', '').replace(',', '').replace('USD', '').strip()
    match = re.search(r'-?\d+\.?\d*', clean_str)
    if match:
        try:
            return float(match.group())
        except ValueError:
            return 0.0
    return 0.0

def parse_date_find_year(date_str):
    """Attempts to parse a date string and return the year."""
    try:
        dt = parser.parse(str(date_str), fuzzy=True)
        return dt.year
    except:
        return None

def extract_data_from_file(uploaded_file):
    """Reads Excel, CSV, PDF, or Word and returns a DataFrame."""
    df = pd.DataFrame()
    try:
        filename = uploaded_file.name.lower()
        if filename.endswith(('.xlsx', '.xls')):
            df = pd.read_excel(uploaded_file)
        elif filename.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        elif filename.endswith('.pdf'):
            with pdfplumber.open(uploaded_file) as pdf:
                all_rows = []
                for page in pdf.pages:
                    tables = page.extract_tables()
                    for table in tables:
                        for row in table:
                            if any(row): all_rows.append(row)
                if all_rows:
                    df = pd.DataFrame(all_rows[1:], columns=all_rows[0])

        elif filename
