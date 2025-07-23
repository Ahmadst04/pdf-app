import streamlit as st
import pandas as pd
import pytesseract
from pdf2image import convert_from_bytes
import io
import cv2
import numpy as np
from PIL import Image

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
    """Run OCR and return structured table with multiple columns."""
    custom_config = r'--psm 6'
    data = pytesseract.image_to_data(image, config=custom_config, output_type=pytesseract.Output.DATAFRAME)
    data = data.dropna().query('text.str.strip() != ""', engine='python')

    # Group by line number to reconstruct rows
    grouped = data.groupby(['page_num', 'block_num', 'par_num', 'line_num'])

    rows = []
    for _, group in grouped:
        # Sort words left to right
        line_data = group.sort_values('left')['text'].tolist()
        rows.append(line_data)

    # Normalize number of columns in each row
    max_cols = max((len(r) for r in rows), default=0)
    padded_rows = [r + [''] * (max_cols - len(r)) for r in rows]

    df = pd.DataFrame(padded_rows)
    return df

def create_excel_file(tables):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for idx, df in enumerate(tables):
            df.to_excel(writer, sheet_name=f'Page{idx + 1}', index=False, header=False)
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
        if st.button("📥 Download as Excel"):
            buffer = create_excel_file(all_tables)
    else:
        st.error("❌ No usable text extracted via OCR.")

    if buffer:
        st.download_button(
            "📥 Download Excel",
            buffer,
            file_name="ocr_table_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

