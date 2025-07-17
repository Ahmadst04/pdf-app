import streamlit as st
import pandas as pd
import pytesseract
from pdf2image import convert_from_bytes
import io
import cv2
import numpy as np
from PIL import Image
import re

st.set_page_config(page_title="PDF OCR Table Extractor", layout="wide")
st.title("📸 Image-based PDF to Excel Converter (with OCR + Smart Text & Numbers)")

uploaded_file = st.file_uploader("Upload a scanned or layout-tricky PDF", type="pdf")
buffer = None

def preprocess_image(pil_image):
    """Convert PIL image to pure black and white for better OCR."""
    img = np.array(pil_image)
    gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)
    _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    kernel = np.ones((1, 1), np.uint8)
    cleaned = cv2.morphologyEx(binary, cv2.MORPH_OPEN, kernel)
    return cleaned

def is_numeric_string(s):
    cleaned = s.replace(',', '').replace('RM', '').replace('$', '').strip()
    return bool(re.match(r'^-?\d+(\.\d+)?$', cleaned))

def convert_if_number(s):
    try:
        cleaned = s.replace(',', '').replace('RM', '').replace('$', '').strip()
        return float(cleaned) if '.' in cleaned else int(cleaned)
    except:
        return s.strip()

def ocr_image_to_table(image):
    """Perform OCR and return lines with smart number + text detection."""
    custom_config = r'--psm 6'
    data = pytesseract.image_to_data(image, config=custom_config, output_type=pytesseract.Output.DATAFRAME)
    data = data.dropna().query('text.str.strip() != ""', engine='python')

    grouped = data.groupby(['page_num', 'block_num', 'par_num', 'line_num'])

    rows = []
    for _, group in grouped:
        words = group.sort_values('left')['text'].tolist()
        parsed = [convert_if_number(word) for word in words]
        line = " ".join([str(w) for w in parsed])
        rows.append([line])

    df = pd.DataFrame(rows, columns=["Full Text"])
    return df

def create_excel_file(tables):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for idx, df in enumerate(tables):
            df.to_excel(writer, sheet_name=f'Table{idx + 1}', index=False)
    output.seek(0)
    return output

if uploaded_file:
    file_bytes = uploaded_file.read()
    images = convert_from_bytes(file_bytes, dpi=300)
    all_tables = []

    for page_num, img in enumerate(images):
        st.markdown(f"---\n### 🖼 Page {page_num + 1}")
        st.image(img, caption=f"Original Page {page_num + 1}", use_column_width=True)

        processed = preprocess_image(img)
        st.image(processed, caption="🧼 Preprocessed (Black & White)", use_column_width=True, channels="GRAY")

        df = ocr_image_to_table(processed)
        if not df.empty:
            all_tables.append(df)
            st.dataframe(df)
        else:
            st.warning("⚠️ No text detected on this page.")

    if all_tables:
        if st.button("📥 Generate Excel"):
            buffer = create_excel_file(all_tables)
    else:
        st.error("❌ No usable text extracted via OCR.")

    if buffer:
        st.download_button("📥 Download Excel", buffer, file_name="ocr_smart_data.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
