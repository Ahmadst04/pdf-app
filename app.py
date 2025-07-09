import streamlit as st
import pandas as pd
import pytesseract
from pdf2image import convert_from_bytes
import io
import cv2
import numpy as np
from PIL import Image
import re

st.set_page_config(page_title="Smart PDF OCR to Excel", layout="wide")
st.title("🧠 Smart PDF OCR to Excel Converter")

uploaded_file = st.file_uploader("Upload a scanned or tricky-layout PDF", type="pdf")
buffer = None

def preprocess_image(pil_image):
    """Convert image to black and white for OCR."""
    img = np.array(pil_image)
    gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)
    _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    cleaned = cv2.morphologyEx(binary, cv2.MORPH_OPEN, np.ones((1, 1), np.uint8))
    return cleaned

def is_numeric_string(s):
    s = s.replace(',', '').replace('RM', '').replace('$', '').strip()
    return bool(re.match(r'^-?\d+(\.\d+)?$', s))

def format_numerics(row):
    return [float(cell.replace(',', '').replace('RM', '').replace('$', '').strip()) if is_numeric_string(cell) else cell for cell in row]

def detect_table_from_layout(image):
    custom_config = r'--psm 6'
    data = pytesseract.image_to_data(image, config=custom_config, output_type=pytesseract.Output.DATAFRAME)
    data = data.dropna(subset=['text'])
    data = data[data['text'].str.strip() != '']
    
    if data.empty:
        return pd.DataFrame()

    rows = []
    current_line_top = -1
    current_line = []

    # Sort by vertical position first
    data = data.sort_values(['top', 'left'])

    for _, row in data.iterrows():
        if abs(row['top'] - current_line_top) > 10:
            # New line detected
            if current_line:
                rows.append(format_numerics(current_line))
            current_line = [row['text']]
            current_line_top = row['top']
        else:
            current_line.append(row['text'])

    if current_line:
        rows.append(format_numerics(current_line))

    # Convert to DataFrame with variable columns
    max_len = max(len(r) for r in rows)
    for r in rows:
        r.extend([""] * (max_len - len(r)))
    df = pd.DataFrame(rows)
    return df

def create_excel_file(tables):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for idx, df in enumerate(tables):
            df.to_excel(writer, sheet_name=f'Table{idx + 1}', index=False, header=False)
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

        df = detect_table_from_layout(processed)
        if not df.empty:
            all_tables.append(df)
            st.dataframe(df)
        else:
            st.warning("⚠️ No table-like structure detected on this page.")

    if all_tables:
        if st.button("📥 Generate Excel"):
            buffer = create_excel_file(all_tables)
    else:
        st.error("❌ No usable data extracted.")

    if buffer:
        st.download_button("📥 Download Excel", buffer, file_name="smart_ocr_tables.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


    if buffer:
        st.download_button("📥 Download Excel", buffer, file_name="ocr_lines.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
