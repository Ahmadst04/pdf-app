import streamlit as st
import pdfplumber
import io
import pytesseract
from pdf2image import convert_from_bytes
import pandas as pd

st.set_page_config(page_title="PDF to Excel - Raw Text Export", layout="wide")
st.title("📄 PDF to Excel Converter - Raw Text Export with OCR Fallback")1

uploaded_file = st.file_uploader("Upload a PDF", type="pdf")
buffer = None

def extract_text_from_pdf(file_bytes):
    texts = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text and text.strip():
                texts.append(text)
            else:
                texts.append("")  # blank if no text
    return texts

def extract_text_with_ocr(file_bytes):
    images = convert_from_bytes(file_bytes)
    extracted = []
    for img in images:
        text = pytesseract.image_to_string(img)
        extracted.append(text)
    return extracted

def create_excel_with_text(pages_text):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for idx, text in enumerate(pages_text):
            # Convert raw text to a dataframe with one line per row
            lines = text.split('\n')
            df = pd.DataFrame(lines, columns=["Text"])
            sheet_name = f"Page{idx+1}"
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    output.seek(0)
    return output

if uploaded_file:
    file_bytes = uploaded_file.read()

    # Step 1: Extract text with pdfplumber
    texts = extract_text_from_pdf(file_bytes)

    # Step 2: Check if any page has text
    if all(not t.strip() for t in texts):
        st.warning("⚠️ No text found in PDF pages. Trying OCR fallback...")
        texts = extract_text_with_ocr(file_bytes)

    if any(t.strip() for t in texts):
        st.success(f"✅ Extracted text from {len(texts)} pages.")

        if st.button("📥 Generate Excel with Raw Text"):
            buffer = create_excel_with_text(texts)

        if buffer:
            st.download_button("📥 Download Excel", buffer,
                               file_name="pdf_raw_text.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.error("❌ No text found in PDF even after OCR fallback.")
