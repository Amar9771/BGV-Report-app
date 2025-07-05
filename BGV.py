import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import io
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode

# ----------------------------
# 📅 Public Holidays
# ----------------------------
public_holidays = pd.to_datetime([
    "2025-01-26", "2025-08-15", "2025-10-02", "2025-12-25"
])

# ----------------------------
# 🧠 Utility Functions
# ----------------------------
def is_working_day(date):
    if date.weekday() == 6:
        return False
    if date.weekday() == 5:
        week = (date.day - 1) // 7 + 1
        if week in [2, 4]:
            return False
    if date in public_holidays:
        return False
    return True

def add_working_days(start_date, n):
    date = start_date
    count = 0
    while count < n:
        date += timedelta(days=1)
        if is_working_day(date):
            count += 1
    return date

def calculate_due(row):
    if pd.notnull(row['BGV_Reinitiated']):
        return add_working_days(row['BGV_Reinitiated'], 8)
    elif pd.notnull(row['BGV_Received On']):
        return add_working_days(row['BGV_Received On'], 15)
    return pd.NaT

def calculate_remarks(row):
    dispatch = row['BGV_Final Dispatch']
    due = row['Final TAT Due Date for Report']
    if pd.isnull(dispatch) or pd.isnull(due):
        return "Pending", ""
    diff = (dispatch - due).days
    if diff <= 0:
        return "Within TAT", ""
    return "Exceeded", f"{diff} days Deduction"

def process_report(df):
    for col in df.columns:
        if "Date" in col or "On" in col or "Reinitiated" in col or "Dispatch" in col:
            df[col] = pd.to_datetime(df[col], errors='coerce')

    df['Final TAT Due Date for Report'] = df.apply(calculate_due, axis=1)
    df[['Remarks', 'Due Days']] = df.apply(lambda row: pd.Series(calculate_remarks(row)), axis=1)

    for col in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            df[col] = df[col].dt.strftime('%d-%b-%Y')

    return df

def style_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl', date_format='DD-MMM-YYYY') as writer:
        df.to_excel(writer, sheet_name='BGV_Report', index=False)
        sheet = writer.sheets['BGV_Report']

        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="003366", end_color="003366", fill_type="solid")
        alt_fill = PatternFill(start_color="F7F7F7", end_color="F7F7F7", fill_type="solid")
        center_align = Alignment(horizontal='center', vertical='center')

        green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        yellow_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")

        for col_num, col in enumerate(df.columns, 1):
            cell = sheet.cell(row=1, column=col_num)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center_align

        remarks_col = df.columns.get_loc("Remarks") + 1
        due_col = df.columns.get_loc("Due Days") + 1

        for row in range(2, sheet.max_row + 1):
            remark_value = sheet.cell(row=row, column=remarks_col).value
            row_fill = None
            if remark_value == "Within TAT":
                row_fill = green_fill
            elif remark_value == "Exceeded":
                row_fill = red_fill
            elif remark_value == "Pending":
                row_fill = yellow_fill

            for col in range(1, sheet.max_column + 1):
                cell = sheet.cell(row=row, column=col)
                if row % 2 == 0 and not row_fill:
                    cell.fill = alt_fill
                if col == remarks_col or col == due_col:
                    if row_fill:
                        cell.fill = row_fill
                cell.alignment = center_align

        for col_num, col in enumerate(df.columns, 1):
            max_len = max(df[col].astype(str).map(len).max(), len(col))
            sheet.column_dimensions[get_column_letter(col_num)].width = max_len + 2

    output.seek(0)
    return output

# ----------------------------
# 🌐 Streamlit UI
# ----------------------------
st.set_page_config("BGV Report Generator", layout="wide", page_icon="📋")
st.markdown("""
    <style>
    .main {background-color: #f9f9f9;}
    .block-container {padding-top: 2rem; padding-bottom: 2rem;}
    .css-18e3th9 {background-color: #ffffff; padding: 2rem; border-radius: 12px; box-shadow: 0 0 20px rgba(0,0,0,0.05);}
    .css-1d391kg h1 {color: #003366;}
    </style>
""", unsafe_allow_html=True)

st.title("📋 BGV Final TAT Report Generator")
st.markdown("Generate professional reports with SLA compliance in one click.")

with st.expander("📌 Instructions", expanded=False):
    st.markdown("""
    **Step 1**: Download the Excel template  
    **Step 2**: Fill out the candidate data  
    **Step 3**: Upload the filled file to generate report  
    **Note**: All dates must be in valid Excel date format.
    """)

template_columns = [
    "Sl.No", "CandidateCode", "Candidate Name",
    "BWR_Date of Submission", "BWR_TAT Due On", "BWR_Reinitiated", "BWR_Date of Report Received",
    "BGV_Received On", "BGV_TAT Due On", "BGV_Reinitiated", "BGV_Final Dispatch"
]

st.subheader("📥 Step 1: Download Template")
template_df = pd.DataFrame(columns=template_columns)
template_buf = io.BytesIO()
template_df.to_excel(template_buf, index=False)
st.download_button("⬇️ Download Excel Template", template_buf.getvalue(), file_name="BGV_Template.xlsx")

st.subheader("📤 Step 2: Upload Filled File")
uploaded_file = st.file_uploader("Upload the filled BGV Excel file", type="xlsx")

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        missing_cols = [col for col in template_columns if col not in df.columns]
        if missing_cols:
            st.error(f"❌ Missing columns: {', '.join(missing_cols)}")
        else:
            result_df = process_report(df)
            st.success("✅ Report generated successfully!")

            st.subheader("📊 Step 3: Preview Report")
            gb = GridOptionsBuilder.from_dataframe(result_df)
            gb.configure_default_column(filter=True, sortable=True, resizable=True)
            grid_options = gb.build()
            AgGrid(
                result_df,
                gridOptions=grid_options,
                update_mode=GridUpdateMode.NO_UPDATE,
                fit_columns_on_grid_load=True,
                height=450
            )

            styled_file = style_excel(result_df)
            st.download_button("📁 Download Final Excel Report", styled_file, file_name="BGV_Final_TAT_Report.xlsx")
    except Exception as e:
        st.error(f"⚠️ Error processing file: {e}")
