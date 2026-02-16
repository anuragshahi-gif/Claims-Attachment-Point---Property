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

st.title("🛡️ Attachment Point & Policy Checker")
st.markdown("Checks data against an **Attachment Point** and finds the **Oldest Policy Year**.")

# 1. Attachment Point Input
attachment_point = st.number_input(
    "Enter Attachment Point ($):", 
    min_value=0.0, 
    value=0.0, 
    step=1000.0,
    format="%.2f",
    help="Rows with values GREATER than this will be flagged."
)

# 2. Input Method Selection & Reset Button
col_input, col_reset = st.columns([4, 1])

with col_input:
    input_method = st.radio("Data Source:", ["Upload File", "Paste Text"], horizontal=True)

with col_reset:
    st.write("") # Spacer
    st.write("") # Spacer
    st.button("🔄 Reset All Data", on_click=reset_all, type="secondary")

df = pd.DataFrame()
target_col_index = None
policy_col_index = None

# 3 & 4. Logic based on Input Method
if input_method == "Upload File":
    uploaded_file = st.file_uploader(
        "Upload Excel, CSV, PDF, or Word", 
        type=['xlsx', 'csv', 'pdf', 'docx'],
        key=f"file_uploader_{st.session_state['uploader_key']}"
    )
    
    if uploaded_file:
        df = extract_data_from_file(uploaded_file)
        if not df.empty:
            st.info(f"File loaded: {len(df)} rows, {len(df.columns)} columns.")
            st.dataframe(df.head(3))
            
            c1, c2 = st.columns(2)
            with c1:
                val_col_num = st.number_input("Column # for Value/Claim:", min_value=1, max_value=len(df.columns), value=1)
                target_col_index = val_col_num - 1
            with c2:
                policy_col_num = st.number_input("Column # for Policy/Date:", min_value=1, max_value=len(df.columns), value=1)
                policy_col_index = policy_col_num - 1

elif input_method == "Paste Text":
    raw_text = st.text_area("Paste Data (Copy from Excel/CSV):", height=200, key="text_input")
    
    if raw_text:
        df = parse_pasted_text(raw_text)
        if not df.empty:
            st.success("Text parsed successfully.")
            st.dataframe(df.head(3))
            target_col_name = st.selectbox("Select Value Column:", df.columns)
            policy_col_name = st.selectbox("Select Policy/Date Column:", df.columns)
            if target_col_name: target_col_index = df.columns.get_loc(target_col_name)
            if policy_col_name: policy_col_index = df.columns.get_loc(policy_col_name)

# 5, 6 & 7. Processing Logic
if st.button("Analyze Data", type="primary"):
    if df.empty:
        st.warning("No data found to analyze.")
    elif target_col_index is None:
        st.warning("Please specify columns.")
    else:
        # Filter Logic
        target_col_name = df.columns[target_col_index]
        df['__clean_value__'] = df[target_col_name].apply(clean_currency)
        filtered_df = df[df['__clean_value__'] >= attachment_point].copy()
        
        # Date Logic
        oldest_year = "N/A"
        if policy_col_index is not None:
            policy_col_name = df.columns[policy_col_index]
            df['__parsed_year__'] = df[policy_col_name].apply(parse_date_find_year)
            sorted_by_date = df.sort_values(by='__parsed_year__', ascending=True)
            min_year_row = sorted_by_date[sorted_by_date['__parsed_year__'].notnull()].head(1)
            if not min_year_row.empty:
                oldest_year = int(min_year_row['__parsed_year__'].iloc[0])

        # Display
        st.divider()
        st.subheader("Analysis Results")
        
        m1, m2 = st.columns(2)
        m1.metric("Oldest Policy Year", str(oldest_year))
        m2.metric("Rows Exceeding Limit", len(filtered_df))
        
        if not filtered_df.empty:
            display_df = filtered_df.drop(columns=['__clean_value__', '__parsed_year__'], errors='ignore')
            st.dataframe(display_df)
        else:
            st.info("No records found exceeding the attachment point.")
