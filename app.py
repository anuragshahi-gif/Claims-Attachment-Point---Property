import streamlit as st
import pandas as pd
import pdfplumber
import docx
import re
import io
from io import BytesIO
from dateutil import parser

# --- CONFIGURATION ---
st.set_page_config(page_title="Claims Loss Analyzer", layout="wide")

# --- SESSION STATE & RESET LOGIC ---
if "uploader_key" not in st.session_state:
    st.session_state["uploader_key"] = 0

if "text_input" not in st.session_state:
    st.session_state["text_input"] = ""

def reset_all():
    """Resets all inputs and clears the file uploader."""
    st.session_state["uploader_key"] += 1
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
        elif filename.endswith('.docx'):
            doc = docx.Document(uploaded_file)
            all_rows = []
            for table in doc.tables:
                for row in table.rows:
                    text_row = [cell.text.strip() for cell in row.cells]
                    all_rows.append(text_row)
            if all_rows:
                df = pd.DataFrame(all_rows[1:], columns=all_rows[0])
    except Exception as e:
        st.error(f"Error reading file: {e}")
    return df

def parse_pasted_text(text):
    """Parses pasted text assuming it is Tab or Comma separated."""
    try:
        data = io.StringIO(text)
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

# 1. Attachment Point Input
attachment_point = st.number_input(
    "Enter Attachment Point ($):", 
    min_value=0.0, 
    value=0.0, 
    step=1000.0,
    format="%.2f",
    help="Set to 0 to generate a Summary by Year. Set > 0 to filter specific large claims."
)

if attachment_point == 0:
    st.info("ℹ️ Attachment Point is 0. **Summary Mode** is active.")
else:
    st.info(f"ℹ️ Filtering claims greater than **${attachment_point:,.2f}**.")

# 2. Input Method Selection & Reset Button
col_input, col_reset = st.columns([4, 1])
with col_input:
    input_method = st.radio("Data Source:", ["Upload File", "Paste Text"], horizontal=True)
with col_reset:
    st.write("") 
    st.write("") 
    st.button("🔄 Reset", on_click=reset_all, type="secondary")

df = pd.DataFrame()
policy_col_index = None
incurred_col_index = None
paid_col_index = None

# 3. Logic based on Input Method
if input_method == "Upload File":
    uploaded_file = st.file_uploader(
        "Upload Excel, CSV, PDF, or Word", 
        type=['xlsx', 'csv', 'pdf', 'docx'],
        key=f"file_uploader_{st.session_state['uploader_key']}"
    )
    
    if uploaded_file:
        df = extract_data_from_file(uploaded_file)
        if not df.empty:
            st.success(f"Loaded {len(df)} rows.")
            st.dataframe(df.head(3))
            
            c1, c2, c3 = st.columns(3)
            with c1:
                p_idx = st.number_input("Col # for Policy/Date:", min_value=1, max_value=len(df.columns), value=1)
                policy_col_index = p_idx - 1
            with c2:
                i_idx = st.number_input("Col # for Total Incurred:", min_value=1, max_value=len(df.columns), value=1)
                incurred_
