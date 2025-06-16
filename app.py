import streamlit as st
import pdfplumber
import pandas as pd
import io
from collections import defaultdict

st.set_page_config(page_title="PDF Table Extractor", layout="wide")
st.title("📄 PDF to Excel Converter")

uploaded_file = st.file_uploader("Upload a PDF", type="pdf")
buffer = None

def group_words_to_rows(words, row_tolerance=3):
    """Group extracted words into rows by vertical position."""
    rows = defaultdict(list)
    for word in words:
        top_key = round(word['top'] / row_tolerance) * row_tolerance
        rows[top_key].append((word['x0'], word['text']))
    return rows

def extract_table_from_words(words):
    """Organize words into a structured table based on position."""
    rows = group_words_to_rows(words)
    structured_rows = []
    for top in sorted(rows.keys()):
        row = [text for _, text in sorted(rows[top])]
        structured_rows.append(row)
    return pd.DataFrame(structured_rows)

def create_excel_file(tables):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for idx, df in enumerate(tables):
            df.to_excel(writer, sheet_name=f'Table{idx + 1}', index=False, header=False)
    output.seek(0)
    return output

if uploaded_file:
    file_bytes = uploaded_file.read()
    all_tables = []

    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page_num, page in enumerate(pdf.pages):
            words = page.extract_words(x_tolerance=2, y_tolerance=2)
            if words:
                df = extract_table_from_words(words)
                if not df.empty:
                    all_tables.append(df)

    if all_tables:
        st.success(f"✅ Extracted {len(all_tables)} table(s) from PDF.")
        for idx, df in enumerate(all_tables):
            st.markdown(f"---\n### 📑 Table {idx + 1}")
            st.dataframe(df)

        if st.button("📥 Download Excel"):
            buffer = create_excel_file(all_tables)
    else:
        st.error("❌ No usable table data found.")

    if buffer:
        st.download_button("📥 Download Excel", buffer, file_name="tables_from_pdf.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
