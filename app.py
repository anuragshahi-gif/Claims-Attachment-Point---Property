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
elif attachment_point is not None and attachment_point > 0:
    st.info(f"ℹ️ Filtering claims greater than **${attachment_point:,.2f}**.")

# 2. Reset Button
col_spacer, col_reset = st.columns([6, 1])
with col_reset:
    st.button("🔄 Reset Text", on_click=reset_all, type="secondary")

# 3. Text Input Area (ALWAYS VISIBLE)
raw_text = st.text_area(
    "Paste Data (Copy from Excel/CSV):", 
    height=300, 
    key="text_input",
    help="Copy your Excel data (including headers) and paste it here."
)

df = pd.DataFrame()
policy_col_index = None
incurred_col_index = None
paid_col_index = None

# 4. Data Parsing & Column Selection
if raw_text:
    df = parse_pasted_text(raw_text)
    if not df.empty:
        st.success("Data parsed successfully.")
        st.dataframe(df.head(3))
        
        c1, c2, c3 = st.columns(3)
        with c1:
            pol_name = st.selectbox("Policy/Date Col:", df.columns, index=0)
            if pol_name: policy_col_index = df.columns.get_loc(pol_name)
        with c2:
            inc_name = st.selectbox("Total Incurred Col:", df.columns, index=0)
            if inc_name: incurred_col_index = df.columns.get_loc(inc_name)
        with c3:
            paid_name = st.selectbox("Total Paid Col:", df.columns, index=0)
            if paid_name: paid_col_index = df.columns.get_loc(paid_name)

# 5. Processing Logic
if st.button("Analyze Data", type="primary"):
    if df.empty:
        st.warning("Please paste data first.")
    elif attachment_point is None:
        st.warning("Please enter an Attachment Point (0 for Summary, or a specific amount).")
    elif None in [policy_col_index, incurred_col_index, paid_col_index]:
        st.warning("Please specify all three columns (Policy Date, Incurred, Paid).")
    else:
        col_policy = df.columns[policy_col_index]
        col_incurred = df.columns[incurred_col_index]
        col_paid = df.columns[paid_col_index]

        # Clean Data
        df['__Year__'] = df[col_policy].apply(parse_date_find_year)
        df['__Incurred__'] = df[col_incurred].apply(clean_currency)
        df['__Paid__'] = df[col_paid].apply(clean_currency)
        
        df_clean = df.dropna(subset=['__Year__']).copy()
        df_clean['__Year__'] = df_clean['__Year__'].astype(int)

        st.divider()

        # --- MODE 1: SUMMARY BY YEAR (Attachment Point = 0) ---
        if attachment_point == 0:
            st.subheader("📑 Loss Summary by Policy Year")
            
            summary = df_clean.groupby('__Year__').agg(
                Claims_Count=('__Year__', 'count'),
                Total_Incurred=('__Incurred__', 'sum'),
                Total_Paid=('__Paid__', 'sum')
            ).reset_index()
            
            summary.columns = ['Policy Year', '# of Claims', 'Total Incurred', 'Total Paid']
            summary = summary.sort_values(by='Policy Year', ascending=False)
            
            # Display readable version
            display_summary = summary.copy()
            display_summary['Total Incurred'] = display_summary['Total Incurred'].apply(lambda x: f"${x:,.2f}")
            display_summary['Total Paid'] = display_summary['Total Paid'].apply(lambda x: f"${x:,.2f}")
            
            st.dataframe(display_summary, hide_index=True, use_container_width=True)
            
            # DOWNLOAD BUTTON FOR SUMMARY
            st.download_button(
                label="📥 Download Summary to Excel",
                data=to_excel(summary),
                file_name="loss_summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        # --- MODE 2: FILTER BY ATTACHMENT POINT (Attachment Point > 0) ---
        else:
            st.subheader(f"⚠️ Claims Exceeding ${attachment_point:,.2f}")
            
            filtered_df = df_clean[df_clean['__Incurred__'] >= attachment_point].copy()
            
            if not filtered_df.empty:
                st.write(f"Found **{len(filtered_df)}** claims.")
                
                display_cols = [col for col in df.columns if col not in ['__Year__', '__Incurred__', '__Paid__']]
                final_view = filtered_df[display_cols]
                st.dataframe(final_view)

                # DOWNLOAD BUTTON FOR FILTERED LIST
                st.download_button(
                    label="📥 Download Filtered List to Excel",
                    data=to_excel(final_view),
                    file_name="large_losses.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.success("No claims found exceeding this attachment point.")
