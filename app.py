import streamlit as st
import pdfplumber
import pandas as pd
import io

st.title("PDF to Excel Converter")

uploaded_file = st.file_uploader("Upload a PDF", type="pdf")

if uploaded_file:
    with pdfplumber.open(uploaded_file) as pdf:
        all_tables = []
        for i, page in enumerate(pdf.pages):
            table = page.extract_table()
            if table:
                df = pd.DataFrame(table[1:], columns=table[0])
                all_tables.append(df)

    if all_tables:
        st.success(f"{len(all_tables)} table(s) found.")

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            for i, table in enumerate(all_tables):
                table.to_excel(writer, sheet_name=f'Table{i+1}', index=False)
        buffer.seek(0)

        st.download_button("Download Excel", buffer, file_name="converted.xlsx")
    else:
        st.warning("No tables found in the PDF.")
