import streamlit as st
import pdfplumber
import pandas as pd
import io
from collections import defaultdict

st.set_page_config(page_title="PDF Table Extractor", layout="wide")
st.title("📄 PDF to Excel Converter (Precise Layout Handling)")

uploaded_file = st.file_uploader("Upload a PDF", type="pdf")
buffer = None

def extract_table_by_position(pdf):
    tables = []
    for page in pdf.pages:
        chars = page.chars  # individual characters
        if not chars:
            continue

        # Step 1: Group characters into rows
        rows = defaultdict(list)
        for char in chars:
            row_key = round(char['top'], 1)  # adjust tolerance here
            rows[row_key].append(char)

        # Step 2: Sort and cluster by x-position to approximate columns
        structured_rows = []
        for top in sorted(rows.keys()):
            line = sorted(rows[top], key=lambda c: c['x0'])
            line_text = [c['text'] for c in line]
            joined = "".join(line_text)
            words = joined.split()
            structured_rows.append(words)

        df = pd.DataFrame(structured_rows)
        tables.append(df)

    return tables

def create_excel_file(tables):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for idx, df in enumerate(tables):
            df.to_excel(writer, sheet_name=f'Table{idx + 1}', index=False, header=False)
    output.seek(0)
    return output

if uploaded_file:
    file_bytes = uploaded_file.read()
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        all_tables = extract_table_by_position(pdf)

    if all_tables:
        st.success(f"✅ Extracted {len(all_tables)} table(s) from PDF.")
        for idx, df in enumerate(all_tables):
            st.markdown(f"---\n### 📑 Table {idx + 1}")
            st.dataframe(df)

        if st.button("📥 Download Excel"):
            buffer = create_excel_file(all_tables)
    else:
        st.error("❌ No usable table-like structures found.")

    if buffer:
        st.download_button("📥 Download Excel", buffer, file_name="tables_from_pdf.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
