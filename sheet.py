import streamlit as st
import pandas as pd
import pdfplumber
import re
from io import BytesIO
import openpyxl

def convert_units_in_cell(val, conversion_unit="Lakhs", threshold=20):
    if isinstance(val, str):
        # Exclude days, months, years, dates from conversion
        days_of_week = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]
        if val.lower() in days_of_week:
            return val
        months = ["january", "february", "march", "april", "may", "june", 
                 "july", "august", "september", "october", "november", "december",
                 "jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"]
        if val.lower() in months:
            return val
        year_pattern = r'^(\d{4}|\d{4}-\d{2}|FY\d{4}|\d{4}-\d{4}|FY\d{2}-\d{2})$'
        if re.match(year_pattern, val, re.IGNORECASE):
            return val
        day_suffix_pattern = r'^\d{1,2}(st|nd|rd|th)$'
        if re.match(day_suffix_pattern, val.lower()):
            return val
        date_patterns = [
            r'^(january|february|march|april|may|june|july|august|september|october|november|december)\s+\d{1,2},\s*\d{4}$',
            r'^(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\s+\d{1,2},\s*\d{4}$',
            r'^\d{1,2}/(\d{1,2})/\d{2,4}$',
            r'^\d{1,2}-(\d{1,2})-\d{2,4}$'
        ]
        for pattern in date_patterns:
            if re.match(pattern, val, re.IGNORECASE):
                return val

    conversion_factors = {
        "Hundred": 100,
        "Thousand": 1000,
        "Lakhs": 100000,
        "Crore": 10000000
    }
    factor = conversion_factors[conversion_unit]

    try:
        if isinstance(val, str):
            if str(val).startswith('='):  # skip formulas
                return val
            if re.search(r'\d{1,2}(?:,\d{2})+(?:\.\d+)?', val):
                num_str = val.replace(',', '')
                num = float(num_str)
            else:
                num_str = re.sub(r"[^\d.\-]", "", str(val))
                if num_str.strip() == "" or num_str == ".":
                    return val
                num = float(num_str)
        else:
            num = float(val)

        if abs(num) > threshold:
            return(num / factor)
        else:
            return num
    except:
        return val


def add_unit_row(df, conversion_unit):
    """Add an extra first row showing units like (in Lakhs) only for columns with converted numeric values."""
    unit_row = []
    for col in df.columns:
        has_converted_numeric = False
        for val in df[col]:
            try:
                if pd.notna(val) and not isinstance(val, str):
                    # Check if this value was converted (is a float with 2 decimal places)
                    if isinstance(val, float) and val == round(val, 2) and abs(val) < 1000:
                        has_converted_numeric = True
                        break
                elif pd.notna(val) and isinstance(val, str):
                    # Check if this string represents a converted number
                    if re.match(r'^-?\d+(\.\d{1,2})?$', str(val)):
                        num = float(val)
                        if abs(num) < 1000:  # Assuming converted values are smaller
                            has_converted_numeric = True
                            break
            except:
                continue
        if has_converted_numeric:
            unit_row.append(f"(in {conversion_unit})")
        else:
            unit_row.append("")
    df_with_units = pd.DataFrame([unit_row], columns=df.columns)
    df_with_units = pd.concat([df_with_units, df], ignore_index=True)
    return df_with_units


def extract_tables_from_pdf(file_bytes, conversion_unit):
    all_tables = {}
    with pdfplumber.open(file_bytes) as pdf:
        for i, page in enumerate(pdf.pages):
            tables = page.extract_tables()
            if tables:
                for j, table in enumerate(tables):
                    df = pd.DataFrame(table)
                    df = df.applymap(lambda x: convert_units_in_cell(x, conversion_unit))
                    df = add_unit_row(df, conversion_unit)
                    all_tables[f"Page_{i+1}_Table_{j+1}"] = df
    return all_tables


def create_preserve_excel(excel_bytes, conversion_unit):
    wb_data = openpyxl.load_workbook(excel_bytes, data_only=True)
    wb = openpyxl.load_workbook(excel_bytes, data_only=False)

    for ws_name in wb.sheetnames:
        ws_data = wb_data[ws_name]
        ws = wb[ws_name]

        # convert cell values
        for row in range(1, ws.max_row + 1):
            for col in range(1, ws.max_column + 1):
                cell = ws.cell(row=row, column=col)
                cell_data = ws_data.cell(row=row, column=col)
                if cell.value is None:
                    continue
                if cell.data_type == "f":
                    try:
                        calculated_value = cell_data.value
                        if calculated_value is not None:
                            converted_value = convert_units_in_cell(calculated_value, conversion_unit)
                            cell.value = converted_value
                    except:
                        pass
                else:
                    cell.value = convert_units_in_cell(cell.value, conversion_unit)

        # add top row with units
        unit_row = []
        for col in range(1, ws.max_column + 1):
            has_numeric = False
            for row in range(2, ws.max_row + 1):
                val = ws.cell(row=row, column=col).value
                try:
                    if val is not None:
                        float(val)
                        has_numeric = True
                        break
                except:
                    continue
            if has_numeric:
                unit_row.append(f"(in {conversion_unit})")
            else:
                unit_row.append("")
        ws.insert_rows(1)
        for idx, val in enumerate(unit_row, start=1):
            ws.cell(row=1, column=idx).value = val

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# Streamlit UI
st.set_page_config(page_title="Balance Sheet Converter", layout="wide")
st.title("📊 Balance Sheet Converter")

conversion_unit = st.selectbox(
    "Select conversion unit:",
    ["Crore", "Lakhs", "Thousand", "Hundred"],
    index=1
)

uploaded_file = st.file_uploader("Choose an Excel or PDF file", type=["xlsx", "xls", "pdf"])

if uploaded_file is not None:
    file_type = uploaded_file.name.split('.')[-1].lower()
    file_bytes = uploaded_file.read()

    if file_type in ["xlsx", "xls"]:
        st.success(f"Processing Excel file: {uploaded_file.name}")

        st.subheader("Original Values Preview")
        try:
            original_df = pd.read_excel(BytesIO(file_bytes))
            st.dataframe(original_df)
        except Exception as e:
            st.warning(f"Could not display original file preview: {e}")

        excel_output = create_preserve_excel(BytesIO(file_bytes), conversion_unit)

        st.subheader("Converted Values Preview")
        try:
            converted_df = pd.read_excel(excel_output)
            st.dataframe(converted_df)
        except Exception as e:
            st.warning(f"Could not display converted file preview: {e}")

        st.download_button(
            label=f"📥 Download Converted Excel (values converted to {conversion_unit})",
            data=excel_output,
            file_name=f"converted_{uploaded_file.name}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    elif file_type == "pdf":
        st.success(f"Processing PDF file: {uploaded_file.name}")
        tables_dict = extract_tables_from_pdf(BytesIO(file_bytes), conversion_unit)

        if tables_dict:
            st.write("### Preview Extracted Tables")
            for sheet_name, df in tables_dict.items():
                st.write(f"**{sheet_name}**")
                st.dataframe(df)

            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                for sheet_name, df in tables_dict.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
            output.seek(0)

            st.download_button(
                label=f"📥 Download Excel (values only, {conversion_unit})",
                data=output,
                file_name=f"converted_{uploaded_file.name.split('.')[0]}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("❌ No tables found in this PDF.")
