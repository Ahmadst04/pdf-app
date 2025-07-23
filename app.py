import streamlit as st
import pandas as pd
import pytesseract
from pdf2image import convert_from_bytes
import io
import cv2
import numpy as np
from PIL import Image
from collections import defaultdict

st.set_page_config(page_title="PDF OCR Table Extractor", layout="wide")
st.title("📸 Image-based PDF to Excel Converter (with OCR)")

uploaded_file = st.file_uploader("Upload a scanned or layout-tricky PDF", type="pdf")
buffer = None

def preprocess_image(pil_image):
    """Convert PIL image to pure black and white for better OCR."""
    img = np.array(pil_image)
    gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)
    
    # Apply Otsu thresholding to binarize
    _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

    # Optional: Morphological clean-up to remove noise
    kernel = np.ones((1, 1), np.uint8)
    cleaned = cv2.morphologyEx(binary, cv2.MORPH_OPEN, kernel)

    return cleaned

def ocr_image_to_table(image):
    custom_config = r'--psm 6'  # Assume uniform block of text
    data = pytesseract.image_to_data(image, config=custom_config, output_type=pytesseract.Output.DATAFRAME)
    data = data.dropna().query('text.str.strip() != ""', engine='python')

    # Group by line number
    grouped = data.groupby(['page_num', 'block_num', 'par_num', 'line_num'])

    rows = []
    for _, group in grouped:
        line = group.sort_values('left')['text'].tolist()
        rows.append(line)

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
        # Show preprocessed black & white version
        st.image(processed, caption="🧼 Preprocessed (Black & White)", use_column_width=True, channels="GRAY")

        df = ocr_image_to_table(processed)
        if not df.empty:
            all_tables.append(df)
            st.dataframe(df)
        else:
            st.warning("⚠️ No text detected on this page.")

    if all_tables:
        if st.button("📥 Download as Excel"):
            buffer = create_excel_file(all_tables)
    else:
        st.error("❌ No usable tables extracted via OCR.")

    if buffer:
        st.download_button("📥 Download Excel", buffer, file_name="ocr_tables.xlsx", mult, `11` for sparse text). This would require modifying the `custom_config` variable.
""")
