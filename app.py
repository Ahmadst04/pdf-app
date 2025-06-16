import streamlit as st
import pdfplumber
import pandas as pd
import io
import pytesseract
from pdf2image import convert_from_bytes
from PIL import Image
import cv2
import numpy as np

st.set_page_config(page_title="PDF to Excel (Tables Only)", layout="wide")
st.title("📄 PDF to Excel Converter with OCR (No Chart)")

uploaded_file = st.file_uploader("Upload a PDF", type="pdf")
buffer = None

def clean_data(df):
    for col in df.columns:
        df[col] = df[col].replace(r'[^\d.-]', '', regex=True)
        df[col] = pd.to_numeric(df[col], errors='ignore')
    return df

def text_to_table(text):
    lines = text.split("\n")
    lines = [line.strip() for line in lines if line.strip()]
    if len(lines) < 2:
        return None

    headers = lines[0].split()
    data = []

    for line in lines[1:]:
        row = line.split()
        if len(row) < len(headers):
            row += [''] * (len(headers) - len(row))
        elif len(row) > len(headers):
            row = row[:len(headers)]
        data.append(row)

    if not data:
        return None

    return pd.DataFrame(data, columns=headers)

def preprocess_image_for_ocr(pil_image):
    img = np.array(pil_image)
    if img.ndim == 3:
        img = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)
    _, img_bin = cv2.threshold(img, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    img_bin = cv2.medianBlur(img_bin, 3)
    return img_bin

def extract_text_with_ocr(file_bytes):
    images = convert_from_bytes(file_bytes)
    extracted = []
    config = r'--oem 3 --psm 6'
    for img in images:
        pre_img = preprocess_image_for_ocr(img)
        text = pytesseract.image_to_string(pre_img, config=config)
        extracted.append(text)
    return extracted

def create_excel_file(tables):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for idx, df in enumerate(tables):
            df.to_excel(writer, sheet_name=f'Table{idx + 1}', index=False)
    output.seek(0)
    return output

if uploaded_file:
    file_bytes = uploaded_file.read()
    all_tables = []
    raw_text_tables = []

    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for i, page in enumerate(pdf.pages):
            table = page.extract_table()
            if table:
                df = pd.DataFrame(table[1:], columns=table[0])
                df = clean_data(df)
                all_tables.append(df)
            else:
                text = page.extract_text()
                if text:
                    df_text = text_to_table(text)
                    if df_text is not None:
                        df_text = clean_data(df_text)
                        raw_text_tables.append(df_text)

    all_tables.extend(raw_text_tables)

    if not all_tables:
        st.warning("⚠️ No tables or text found. Trying OCR (scanned document fallback)...")
        ocr_text_pages = extract_text_with_ocr(file_bytes)
        for text in ocr_text_pages:
            df = text_to_table(text)
            if df is not None:
                df = clean_data(df)
                all_tables.append(df)

    if all_tables:
        st.success(f"✅ Extracted {len(all_tables)} usable table(s).")
        for idx, df in enumerate(all_tables):
            st.markdown(f"---\n### 📑 Table {idx + 1}")
            st.dataframe(df)

        if st.button("📥 Download Excel"):
            buffer = create_excel_file(all_tables)
    else:
        st.error("❌ No usable data found even after OCR.")

    if buffer:
        st.download_button("📥 Download Excel", buffer, file_name="tables_only.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
