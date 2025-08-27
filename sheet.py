import streamlit as st
import pandas as pd
import re
from io import BytesIO
import pdfplumber

def convert_units(df, conversion_unit="Lakhs", threshold=1000):
    conversion_factors = {
        "Hundred": 100,
        "Thousand": 1000,
        "Lakhs": 100000,
        "Crore": 10000000
    }
    factor = conversion_factors[conversion_unit]
    unit_name = conversion_unit.lower()
    for col in df.columns:
        new_col = []
        is_large_numeric = False
        for val in df[col]:
            try:
                num = re.sub(r"[^\d.\-]", "", str(val))
                if num.strip() == "" or num == ".":
                    new_col.append(val)
                else:
                    num = float(num)
                    if abs(num) > threshold:
                        new_col.append(round(num / factor, 2))
                        is_large_numeric = True
                    else:
                        new_col.append(num)
            except:
                new_col.append(val)
        df[col] = new_col
        if is_large_numeric:
            df.rename(columns={col: f"{col} (in {conversion_unit})"}, inplace=True)
    return df

def process_excel_all_sheets(file_bytes, conversion_unit):
    xls = pd.ExcelFile(file_bytes)
    sheet_names = xls.sheet_names
    dfs = {}
    for sheet in sheet_names:
        df = pd.read_excel(file_bytes, sheet_name=sheet, header=None)
        df_converted = convert_units(df.copy(), conversion_unit)
        dfs[sheet] = df_converted
    return dfs

def read_pdf(file_bytes):
    text = ""
    with pdfplumber.open(file_bytes) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"
    return text

st.set_page_config(page_title="Excel & PDF Processor", layout="wide")
st.title("📊 Excel & PDF Processor")
st.write("Upload Excel files to convert numbers or read PDF content.")

conversion_unit = st.selectbox(
    "Select conversion unit:",
    ["Crore","Lakhs", "Thousand", "Hundred"],
    index=0
)

uploaded_file = st.file_uploader("Choose an Excel or PDF file", type=["xlsx", "xls", "pdf"])

if uploaded_file:
    file_type = uploaded_file.name.split('.')[-1].lower()
    if file_type in ["xlsx", "xls"]:
        st.success(f"Processing Excel file: {uploaded_file.name}")
        dfs = process_excel_all_sheets(uploaded_file, conversion_unit)
        st.write("### Sheets Preview")
        for sheet, df in dfs.items():
            st.write(f"**Sheet:** {sheet}")
            st.dataframe(df)
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            for sheet, df in dfs.items():
                df.to_excel(writer, index=False, sheet_name=sheet)
        output.seek(0)
        st.download_button(
            label=f"📥 Download Converted Excel (in {conversion_unit})",
            data=output,
            file_name=f"converted_{uploaded_file.name}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    elif file_type == "pdf":
        st.success(f"Reading PDF file: {uploaded_file.name}")
        pdf_text = read_pdf(uploaded_file)
        st.write("### PDF Content")
        st.text_area("Text from PDF:", pdf_text, height=400)
    else:
        st.error("❌ Unsupported file type.")
