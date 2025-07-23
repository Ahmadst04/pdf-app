import streamlit as st
import pytesseract
import fitz  # PyMuPDF
import pandas as pd
import io
import cv2
import numpy as np
from PIL import Image

def extract_text_from_pdf(pdf_file):
    pdf_document = fitz.open(stream=pdf_file.read(), filetype="pdf")
    text_blocks = []

    for page_num in range(len(pdf_document)):
        page = pdf_document.load_page(page_num)
        pix = page.get_pixmap(dpi=300)
        img_data = pix.tobytes("png")
        image = Image.open(io.BytesIO(img_data))

        open_cv_image = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)

        # Gunakan OCR
        d = pytesseract.image_to_data(open_cv_image, output_type=pytesseract.Output.DICT)

        num_boxes = len(d['level'])
        rows = []
        current_row = []
        last_top = None

        for i in range(num_boxes):
            word = d['text'][i].strip()
            if word:
                top = d['top'][i]
                if last_top is None:
                    last_top = top

                # Jika perbezaan posisi 'top' besar, anggap baris baru
                if abs(top - last_top) > 10:
                    if current_row:
                        rows.append(current_row)
                    current_row = [word]
                    last_top = top
                else:
                    current_row.append(word)

        if current_row:
            rows.append(current_row)

        # Simpan setiap baris sebagai senarai dalam text_blocks
        for row in rows:
            text_blocks.append(row)

    return text_blocks

def format_to_excel_data(blocks):
    formatted_data = []
    for block in blocks:
        formatted_data.append(block)
    return formatted_data

def main():
    st.title("📄 PDF Table OCR to Excel Converter")
    uploaded_file = st.file_uploader("Upload PDF", type=["pdf"])

    if uploaded_file is not None:
        with st.spinner("Processing..."):
            extracted_blocks = extract_text_from_pdf(uploaded_file)
            data = format_to_excel_data(extracted_blocks)
            df = pd.DataFrame(data)

            # Buat buffer untuk simpan Excel
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, header=False)

            st.success("Done extracting!")
            st.download_button(
                "📥 Download Excel",
                buffer,
                file_name="ocr_text_single_column.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    main()
