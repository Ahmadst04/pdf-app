import streamlit as st
import pandas as pd
import io
import fitz  # PyMuPDF
from collections import defaultdict

st.set_page_config(page_title="PDF Table Extractor with PyMuPDF", layout="wide")
st.title("📄 PDF to Excel Converter (via PyMuPDF)")

uploaded_file = st.file_uploader("Upload a PDF", type="pdf")
buffer = None

def group_blocks_to_rows(blocks, row_tolerance=5):
    """Group text blocks by approximate Y coordinate."""
    rows = defaultdict(list)
    for block in blocks:
        y = round(block["top"] / row_tolerance) * row_tolerance
        rows[y].append((block["x0"], block["text"]))
    return rows

def extract_table_from_pymupdf(doc):
    tables = []
    for page in doc:
        words = page.get_text("dict")["blocks"]
        text_blocks = []
        for block in words:
            if "lines" in block:
                for line in block["lines"]:
                    for span in line["spans"]:
                        text_blocks.append({
                            "text": span["text"],
                            "x0": span["bbox"][0],
                            "top": span["bbox"][1]
                        })
        if not text_blocks:
            continue
        rows = group_blocks_to_rows(text_blocks)
        structured_rows = []
        for y in sorted(rows.keys()):
            row = [text for _, text in sorted(rows[y])]
            structured_rows.append(row)
        if structured_rows:
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
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    all_tables = extract_table_from_pymupdf(doc)

    if all_tables:
        st.success(f"✅ Extracted {len(all_tables)} table(s).")
        for idx, df in enumerate(all_tables):
            st.markdown(f"---\n### 📑 Table {idx + 1}")
            st.dataframe(df)

        if st.button("📥 Download Excel"):
            buffer = create_excel_file(all_tables)
    else:
        st.error("❌ No table-like text detected.")

    if buffer:
        st.download_button("📥 Download Excel", buffer, file_name="tables_from_pdf.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
