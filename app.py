import streamlit as st
import pandas as pd
import re
import io
from io import BytesIO
from dateutil import parser

# --- CONFIGURATION ---
st.set_page_config(page_title="Claims Loss Analyzer", layout="wide")

# --- SESSION STATE & RESET LOGIC ---
if "text_input" not in st.session_state:
    st.session_state["text_input"] = ""

def reset_all():
    """Resets the text input."""
    st.session_state["text_input"] = ""

# --- UTILITY FUNCTIONS ---

def clean_currency(value):
    """Converts string currency to float. Returns 0.0 if failed."""
    if isinstance(value, (int, float)):
        return float(value)
    
    clean_str = str(value).upper().replace('$', '').replace(',', '').replace('USD', '').strip()
    if '(' in clean_str and ')' in clean_str:
        clean_str = '-' + clean_str.replace('(', '').replace(')', '')
        
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
        if pd.isna(date_str) or str(date_str).strip() == "":
            return None
        dt = parser.parse(str(date_str), fuzzy=True)
        return dt.year
    except:
        return None

def to_excel(df):
    """Converts a DataFrame to a downloadable Excel byte stream."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()

def parse_pasted_text(text):
    """Parses pasted text assuming it is Tab or Comma separated."""
    try:
        data = io.StringIO(text)
        # 'sep=None' forces pandas to sniff the delimiter (tabs vs commas)
        df = pd.read_csv(data, sep=None, engine='python')
        return df
    except:
        return pd.DataFrame()

# --- FRONTEND UI ---

st.title("📊 Claims Loss Analyzer & Summarizer")
st.markdown("""
* **Mode 1: Large Loss Filter** (Enter an Attachment Point > 0)
* **Mode 2: Loss Summary by Year** (Enter Attachment Point = 0)
""")

# 1. Attachment Point Input (Default is Blank/None)
attachment_point = st.number_input(
    "Enter Attachment Point ($):", 
    min_value=0.0, 
    value=None, 
    step=1000.0,
    format="%.2f",
    help="Leave blank to start. Enter 0 for Summary. Enter amount for Filter."
)

if attachment_point == 0:
    st.info("ℹ️ Attachment Point is 0. **Summary Mode** is active.")
elif attachment_point is not
