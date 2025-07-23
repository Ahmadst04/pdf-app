import streamlit as st
import pytesseract
from pdf2image import convert_from_bytes
from PIL import Image
import pandas as pd
import io
import cv2
import numpy as np

st.set_page_config(page_title="PDF to Excel OCR", layout="centered")

st.title("📄 PDF to Excel OCR Converter")

uploaded_file = st.file_uploader("Upload PDF file", type=["pdf"])

def ocr_image_to_table(image):
    """Run OCR and return structured table while keeping full cell content in one Excel column."""
    custom_config = r'--psm 6'
    data = pytesseract.image_to_data(image, config=custom_config, output_type=pytesseract.Output.DATAFRAME)
    data = data.dropna().query('text.str.strip() != ""', engine='python')

    # Group by line
    grouped = data.groupby(['page_num', 'block_num', 'par_num', 'line_num'])

    rows = []
    for _, group in grouped:
        line = group.sort_values('left')

        # Cluster words into columns based on X position
        threshold = 50  # adjust spacing sensitivity here
        columns = []
        current_col = [line.iloc[0]['text']]
        last_left = line.iloc[0]['left']

        for i in range(1, len(line)):
            word = line.iloc[i]
            if word['left'] - last_left > threshold:
                columns.append(" ".join(current_col))
                current_col = [word['text']]
            else:
                current_col.append(word['text'])
            last_left = word['left']

        columns.append(" ".join(current_col))  # append last column
        rows.append(columns)

    # Pad rows to same length
    max_cols = max((len(r) for r in rows), default=0)
    padded_rows = [r + [''] * (max_cols - len(r)) for r in rows]

    df = pd.DataFrame(padded_rows)
    return df

if uploaded_file is not None:
    with st.spinner("Converting PDF to images..."):
        images = convert_from_bytes(uploaded_file.read())

    final_df = pd.DataFrame()

    with st.spinner("Running OCR and processing..."):
        for page_num, image in enumerate(images):
            # Convert to grayscale for better OCR accuracy
            gray = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2GRAY)
            pil_image = Image.fromarray(gray)

            df = ocr_image_to_table(pil_image)
            final_df = pd.concat([final_df, df, pd.DataFrame([[]])], ignore_index=True)  # blank row between pages

    st.success("OCR completed and table extracted!")

    st.dataframe(final_df)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        final_df.to_excel(writer, index=False, header=False)
        writer.save()

    st.download_button("📥 Download Excel", buffer, file_name="ocr_text_single_column.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
